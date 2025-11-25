"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

import re
import pyodbc
import pandas as pd

from vejman import ProcessCases, map_materiel_to_equipment_types, FetchVejmanPermissions, FetchPricebookData
from datetime import datetime

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""

    orchestrator_connection.log_trace("Running process.")

    # Credentials / constants
    token = orchestrator_connection.get_credential("VejmanToken").password
    pricebook_map = FetchPricebookData(token)

    sql_server = orchestrator_connection.get_constant("SqlServer")
    conn_string = "DRIVER={SQL Server};" + f"SERVER={sql_server.value};DATABASE=VejmanKassen;Trusted_Connection=yes;"
    conn = pyodbc.connect(conn_string)
    cursor = conn.cursor()

    # ------------------------------
    # 1) Load materiel / fakturalinjer config
    # ------------------------------
    materiel_query = """
        SELECT 
            MaterielIDVejman, 
            STRING_AGG(Fakturalinje, ',') AS Fakturalinjer, 
            MIN(FraStartdato) AS EarliestStartDate,
            MIN(FraSlutdato) AS EarliestSlutDate
        FROM [dbo].[VejmanFakturaTekster]
        GROUP BY MaterielIDVejman
    """
    cursor.execute(materiel_query)
    materiel_rows = cursor.fetchall()

    # Build a config dict: materiel_id -> {fakturalinjer, date range}
    materiel_config: dict[int, dict] = {}
    for r in materiel_rows:
        materiel_id = r.MaterielIDVejman
        fakturalinjer_list = [
            f.strip()
            for f in (r.Fakturalinjer or "").split(",")
            if f and f.strip()
        ]
        materiel_config[materiel_id] = {
            "materiel_id": materiel_id,
            "fakturalinjer": fakturalinjer_list,
            "start_date": r.EarliestStartDate,
            "slut_date": r.EarliestSlutDate,
        }

    # ------------------------------
    # 2) Load current faktura rows into a dict for quick lookup
    # ------------------------------
    cursor.execute("""SELECT * FROM [dbo].[VejmanFakturering]""")
    faktura_db = cursor.fetchall()

    faktura_db_by_vejman_id = {}
    for row in faktura_db:
        if getattr(row, "VejmanFakturaID", None) is not None:
            faktura_db_by_vejman_id[row.VejmanFakturaID] = row

    # ------------------------------
    # 3) Collect unique cases across all materiel/equipmentTypes
    # ------------------------------
    cases_by_id: dict[str, pd.Series] = {}
    case_materiel_ids: dict[str, set[int]] = {}

    for materiel_id, conf in materiel_config.items():
        start_date = conf["start_date"]
        from_end_date = conf["slut_date"]

        # Map materiel -> equipmentType(s) as before
        equipment_types = map_materiel_to_equipment_types(materiel_id)

        for equipment_type in equipment_types:
            df = FetchVejmanPermissions(
                token,
                equipment_type,
                start_date,
                from_end_date,
                orchestrator_connection,
            )

            if df.empty:
                orchestrator_connection.log_info(
                    f'Ingen rækker for equipmentType {equipment_type} (MaterielID {materiel_id}) '
                    f'fra startdato {start_date} og fra slutdato {from_end_date}'
                )
                continue

            # Clean authority_reference_number column
            df["cleaned_authority_reference_number"] = df["authority_reference_number"].apply(
                lambda x: re.sub(r"[^\x20-\x7E]", "", str(x).strip().lower()) if pd.notnull(x) else ""
            )

            # Filter rows based on substring checks for 'faktura sendt', 'faktureres ikke', 'annulleret', and 'fak'
            filtered_rows = df[
                ~(
                    df["cleaned_authority_reference_number"].str.contains("faktura sendt")
                    | df["cleaned_authority_reference_number"].str.contains("faktureres ikke")
                    | df["cleaned_authority_reference_number"].str.contains("annulleret")
                    | (df["cleaned_authority_reference_number"] == "fak")
                )
                & (df["initials"] != "JADT")
            ]

            for _, case_row in filtered_rows.iterrows():
                case_id = case_row["case_id"]
                # Save one representative row per case_id
                cases_by_id[case_id] = case_row
                # Track which materiel_ids this case belongs to
                case_materiel_ids.setdefault(case_id, set()).add(materiel_id)

    # ------------------------------
    # 4) Process each unique case, its invoices and DB updates
    # ------------------------------
    ProcessCases(
        cases_by_id=cases_by_id,
        case_materiel_ids=case_materiel_ids,
        materiel_config=materiel_config,
        token=token,
        pricebook_map=pricebook_map,
        conn=conn,
        faktura_db_by_vejman_id=faktura_db_by_vejman_id,
        orchestrator_connection=orchestrator_connection,
    )

    # ------------------------------
    # 5) Update sync history table
    # ------------------------------
    with conn.cursor() as cursor:
        cursor.execute("DELETE FROM [dbo].[VejmanKassenSyncHistory]")
        cursor.execute(
            "INSERT INTO [dbo].[VejmanKassenSyncHistory] (SyncedAt) VALUES (?)",
            datetime.now(),
        )
        conn.commit()

    orchestrator_connection.log_info("VejmanKassen sync timestamp updated.")

