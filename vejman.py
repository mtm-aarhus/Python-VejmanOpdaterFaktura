from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

import requests
import time
import locale
from datetime import datetime
import pandas as pd
import re


def map_materiel_to_equipment_types(materiel_id: int) -> list[int]:
    """
    Reuse your mapping logic:
    - MaterielID 1 -> equipmentTypes [1, 9]
    - MaterielID 2 -> equipmentTypes [2, 7]
    - Other -> [materiel_id]
    """
    if materiel_id == 1:
        return [1, 9]
    if materiel_id == 2:
        return [2, 7]
    return [materiel_id]


def FetchVejmanPermissions(token, equipment_type, fra_startdato, fra_slutdato,
                           orchestrator_connection: OrchestratorConnection):
    combined_cases = []

    with requests.Session() as client:
        url = (
            "https://vejman.vd.dk/permissions/getcases"
            f"?pmCaseStates=8"
            "&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number"
            "%2Cstart_date%2Cstreet_name%2Ccvr_number%2Capplicant%2Cend_date"
            "%2Ccompletion_date%2Cauto_completedcontractor%2Cinitials"
            "&pmCaseWorker=all"
            "&pmCaseTypes=%27rovm%27"
            "&pmCaseVariant=all"
            "&pmCaseTags=ignorerTags"
            "&pmCaseTagShow="
            "&pmCaseShowAttachments=false"
            "&pmAllStates="
            "&dontincludemap=1"
            "&authority=751"
            "&cse="
            f"&equipmentType={equipment_type}"
            f"&startDateFrom={fra_startdato}"
            f"&startDateTo={datetime.today().strftime('%Y-%m-%d')}"
            f"&endDateFrom={fra_slutdato}"
            "&policeDistrictShow="
            f"&_={int(time.time() * 1000)}"
            f"&token={token}"
        )

        response = client.get(url, timeout=500)
        response.raise_for_status()
        json_object = response.json()

        cases = json_object.get("cases", [])
        combined_cases.extend(cases)

    if combined_cases:
        return pd.DataFrame(combined_cases)

    orchestrator_connection.log_info("No new permissions")
    return pd.DataFrame()


def FetchPricebookData(token):
    url = f"https://vejman.vd.dk/services/data.do?table=v_h_pm_pricebook&token={token}"
    response = requests.get(url)
    response.raise_for_status()
    pricebook_data = response.json().get("data", [])
    pricebook_map = {item["text"]: item for item in pricebook_data}
    return pricebook_map


def upsert_issue(conn,
                 *,
                 case_id: str,
                 invoice_id: str,
                 issue_type: str,
                 fakturalinje: str,
                 description: str,
                 fix: str,
                 caseworker_email: str | None,
                 inserted_to_kassen: str):

    inv_id_str = int(invoice_id) if invoice_id else None

    # --------------------------------------------------------
    # 1. CHECK IF ISSUE ALREADY EXISTS AND IS RESOLVED BY USER
    # --------------------------------------------------------
    check_sql = """
        SELECT Status
        FROM dbo.InvoiceIssues
        WHERE CaseID = ? AND InvoiceID = ? AND IssueType = ?
    """

    cur = conn.cursor()
    cur.execute(check_sql, (case_id, inv_id_str, issue_type))
    row = cur.fetchone()

    if row and row.Status in ("UserAccepted", "AutoResolved"):
        # User or auto-resolve has handled this. Robot must not touch it.
        print(f'Already resolved or accepted {case_id} {invoice_id} {issue_type}')
        return


    # --------------------------------------------------------
    # 2. NORMAL UPSERT FOR OPEN OR NON-EXISTING ISSUES
    # --------------------------------------------------------
    merge_sql = """
    MERGE dbo.InvoiceIssues AS target
    USING (
        SELECT ? AS CaseID,
               ? AS InvoiceID,
               ? AS IssueType
    ) AS source
        ON target.CaseID = source.CaseID
       AND target.InvoiceID = source.InvoiceID
       AND target.IssueType = source.IssueType
    WHEN MATCHED THEN
        UPDATE SET
            Fakturalinje     = ?,
            IssueDescription = ?,
            SuggestedFix     = ?,
            CaseworkerEmail  = ?,
            InsertedToKassen = ?,
            Status           = 'Open',
            UpdatedAt        = GETDATE()
    WHEN NOT MATCHED THEN
        INSERT (
            CaseID,
            InvoiceID,
            IssueType,
            Fakturalinje,
            IssueDescription,
            SuggestedFix,
            CaseworkerEmail,
            InsertedToKassen,
            Status,
            CreatedAt,
            UpdatedAt
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Open', GETDATE(), GETDATE());
    """

    params = (
        # USING
        case_id,
        inv_id_str,
        issue_type,
        # UPDATE
        fakturalinje,
        description,
        fix,
        caseworker_email,
        inserted_to_kassen,
        # INSERT
        case_id,
        inv_id_str,
        issue_type,
        fakturalinje,
        description,
        fix,
        caseworker_email,
        inserted_to_kassen,
    )

    cur.execute(merge_sql, params)
    # Commit done by caller


def auto_resolve_missing_issues(conn, case_id: str, current_issue_keys: set[tuple[str, str]]):
    """
    For a given case, mark any previously 'Open' issues as 'AutoResolved'
    if they are NOT in current_issue_keys.

    current_issue_keys contains tuples of (InvoiceID, IssueType) that were
    detected in the current run for this case.
    """
    select_sql = """
    SELECT InvoiceID, IssueType
    FROM dbo.InvoiceIssues
    WHERE CaseID = ? AND Status = 'Open';
    """

    cur = conn.cursor()
    cur.execute(select_sql, (case_id,))
    rows = cur.fetchall()

    for row in rows:
        invoice_id_db = str(row.InvoiceID)
        issue_type_db = row.IssueType
        key = (invoice_id_db, issue_type_db)

        if key not in current_issue_keys:
            update_sql = """
            UPDATE dbo.InvoiceIssues
            SET Status     = 'AutoResolved',
                ResolvedBy = NULL,       -- resolved by data change; user unknown
                ResolvedAt = GETDATE(),
                UpdatedAt  = GETDATE()
            WHERE CaseID   = ?
              AND InvoiceID = ?
              AND IssueType = ?
              AND Status    = 'Open';
            """
            cur.execute(update_sql, (case_id, invoice_id_db, issue_type_db))

def ProcessCases(
    cases_by_id: dict[str, pd.Series],
    case_materiel_ids: dict[str, set[int]],
    materiel_config: dict[int, dict],
    token,
    pricebook_map,
    conn, 
    faktura_db_by_vejman_id: dict,
    orchestrator_connection: OrchestratorConnection
):
    # Danish numeric formatting (if available on the OS)
    locale.setlocale(locale.LC_NUMERIC, "da_DK")

    # Map fakturalinje text -> materiel_id
    fakturalinje_to_materiel: dict[str, int] = {}

    for mid, conf in materiel_config.items():
        for fl in conf["fakturalinjer"]:
            f_clean = fl.strip().lower()
            fakturalinje_to_materiel[f_clean] = mid

    with requests.Session() as client:
        for case_id, case_row in cases_by_id.items():
            case_number = case_row["case_number"]

            current_issue_keys: set[tuple[str, str]] = set()

            # Base data from the list cases API
            start_date = datetime.strptime(case_row.get("start_date", ""), "%d-%m-%Y")
            end_date = datetime.strptime(case_row.get("end_date", ""), "%d-%m-%Y")
            completion_date = datetime.strptime(case_row.get("completion_date", ""), "%d-%m-%Y")
            auto_completed = case_row.get("auto_completed")
            applicant = case_row.get("applicant")
            address = case_row["street_name"]

            # Fetch detailed case data
            response = client.get(
                f"https://vejman.vd.dk/permissions/getcase?caseid={case_id}&token={token}",
                timeout=500,
            )
            response.raise_for_status()
            json_object = response.json().get("data")
            caseworker_email = json_object.get("authEmail")

            # Invoice block
            invoice_data = json_object.get("invoice")
            if not invoice_data:
                # No invoice block at all → nothing to do
                continue

            invoice_details = invoice_data.get("details", [])
            if not invoice_details:
                # Invoice exists but contains no invoice lines → nothing to do
                continue

            # Determine invoice contact (ATT)
            invoice_role_id = invoice_data.get("role", {}).get("id", 1)
            att = "Intet navn angivet"
            contacts = json_object.get("contacts", [])

            cvr_number = None
            contact_type = None

            for contact in contacts:
                roles = contact.get("roles", [])
                if any(role.get("role", {}).get("id") == invoice_role_id for role in roles):
                    name_parts = [
                        contact.get("given_name", ""),
                        contact.get("middle_name", ""),
                        contact.get("surname", ""),
                    ]
                    combined_name = " ".join(part for part in name_parts if part)
                    if combined_name:
                        att = f"Att: {combined_name}"
                    applicant = contact.get("company_name")
                    cvr_number = contact.get("cvr_number")
                    contact_type = contact.get("type", {}).get("code")
                    break

            # If it's a person (P) and no CVR at all => skip case
            if contact_type == "P" and (not cvr_number or str(cvr_number).strip() == ""):
                continue

            tilladelse_nr = case_number

            # ------------------------
            # CVR-related issues
            # ------------------------
            # These are "per case" issues, so we use a synthetic InvoiceID.

            if not cvr_number or not re.fullmatch(r"\d{8}", str(cvr_number)):
                # CVR missing or not 8 digits
                upsert_issue(
                    conn=conn,
                    case_id=case_id,
                    invoice_id=case_number,
                    issue_type="CVR mangler eller ugyldigt.",
                    fakturalinje="CVR ugyldigt",
                    description=(
                        "CVR-nummeret mangler eller består ikke af præcis 8 cifre."
                    ),
                    fix=(
                        "Ret CVR-nummeret i kontaktoplysningerne i Vejman på den "
                        "interessent, der bliver faktureret. Robotten har midlertidigt "
                        "sat CVR = 00000000."
                    ),
                    caseworker_email=caseworker_email,
                    inserted_to_kassen="No",  # not tied to a specific fakturalinje
                )
                current_issue_keys.add((case_number, "CVR mangler eller ugyldigt."))
                cvr_number = "00000000"
            else:
                cvr_str = str(cvr_number)
                if not is_valid_cvr_mod11(cvr_str):
                    upsert_issue(
                        conn=conn,
                        case_id=case_id,
                        invoice_id=case_number,
                        issue_type="CVR modulus 11 fejlede",
                        fakturalinje="CVR ugyldigt (modulus 11)",
                        description=(
                            f"CVR-nummeret {cvr_str} består af 8 cifre, men opfylder "
                            "ikke modulus-11 kontrollen og er derfor ikke et gyldigt CVR."
                        ),
                        fix=(
                            "Ret CVR-nummeret i kontaktoplysningerne i Vejman."
                        ),
                        caseworker_email=caseworker_email,
                        inserted_to_kassen="No",
                    )
                    current_issue_keys.add((case_number, "CVR modulus 11 fejlede"))

            # Pre-calc allowed fakturalinjer + materiel IDs for this case
            materiel_ids_for_case = sorted(case_materiel_ids.get(case_id, set()))
            allowed_fakturalinjer: list[str] = []
            for mid in materiel_ids_for_case:
                conf = materiel_config.get(mid)
                if conf:
                    allowed_fakturalinjer.extend(conf["fakturalinjer"])

            # ------------------------
            # Process each invoice line
            # ------------------------
            for detail in invoice_details:
                detail_text = detail.get("text", "") or ""
                vejman_faktura_id = detail.get("id")
                invoice_id_str = str(vejman_faktura_id) if vejman_faktura_id is not None else ""

                # Lookup existing faktura row for this invoice line
                matching_row = faktura_db_by_vejman_id.get(vejman_faktura_id)

                # If row exists and is not 'Ny', skip completely
                if matching_row and matching_row.FakturaStatus != "Ny":
                    print(f"{case_number} - {detail_text} - Already exists and not Ny")
                    continue

                detail_text_clean = detail_text.strip().lower()

                matched_fakturalinje = None
                matched_materiel_id = None

                for fl, mid in fakturalinje_to_materiel.items():
                    if fl in detail_text_clean:
                        matched_fakturalinje = fl
                        matched_materiel_id = mid
                        break

                text_is_known = matched_fakturalinje is not None
                materiel_is_allowed = matched_materiel_id in materiel_ids_for_case

                # ------------------------
                # Fakturalinje not known at all
                # ------------------------
                if not text_is_known:
                    # This line is never inserted into Vejmankassen (blocked).
                    upsert_issue(
                        conn=conn,
                        case_id=case_id,
                        invoice_id=invoice_id_str,
                        issue_type="Fakturalinje uden match",
                        fakturalinje="Ingen match",
                        description=(
                            "Fakturalinjen er ikke godkendt til brug i Vejmankassen. "
                            "Robotten kan derfor ikke behandle denne linje."
                        ),
                        fix=(
                            "Ret fakturalinjen i Vejman til en gyldig fakturerbar linje. "
                            "Denne fakturalinje bliver ikke tilføjet til Vejmankassen."
                        ),
                        caseworker_email=caseworker_email,
                        inserted_to_kassen="No",
                    )
                    current_issue_keys.add((invoice_id_str, "Fakturalinje uden match"))
                    # Skip DB insertion completely for this line
                    continue

                # ------------------------
                # Fakturalinje known, but materiel not attached on the case
                # ------------------------
                if text_is_known and not materiel_is_allowed:
                    upsert_issue(
                        conn=conn,
                        case_id=case_id,
                        invoice_id=invoice_id_str,
                        issue_type="Manglende/forkert materiel",
                        fakturalinje=matched_fakturalinje,
                        description=(
                            f"Fakturalinjen matcher materiel nummer {matched_materiel_id}, men "
                            "denne materieltype er ikke tilknyttet tilladelsen i Vejman."                        ),
                        fix=(
                            f"Tilknyt materiel nummer {matched_materiel_id} i Vejman."
                        ),
                        caseworker_email=caseworker_email,
                        inserted_to_kassen="Yes",  # line still goes to Vejmankassen
                    )
                    current_issue_keys.add((invoice_id_str, "Manglende/forkert materiel"))

                # ------------------------
                # Determine dates / chosen_end_date
                # ------------------------
                if matching_row:
                    start_for_calc = datetime.strptime(matching_row.Startdato, "%Y-%m-%d")
                    chosen_end_date = datetime.strptime(matching_row.Slutdato, "%Y-%m-%d")
                else:
                    chosen_end_date = (
                        end_date
                        if auto_completed == "AF"
                        else min(completion_date, end_date)
                        if completion_date and end_date
                        else end_date
                    )
                    start_for_calc = start_date

                if start_date and chosen_end_date:
                    days_difference = (chosen_end_date.date() - start_date.date()).days
                    days_period = days_difference + 1  # inclusive
                else:
                    days_period = None

                # ------------------------
                # Unit price + length / price calc
                # ------------------------
                raw_detail_unit_price = detail.get("unit_price", 0)
                if isinstance(raw_detail_unit_price, str):
                    raw_detail_unit_price = float(raw_detail_unit_price.replace(",", "."))
                detail_unit_price = raw_detail_unit_price

                pricebook_entry = pricebook_map.get(detail_text, {})
                unit_price = float(pricebook_entry.get("unit_price", 0))

                try:
                    match_len = re.search(
                        r"\d+(\.\d+)?",
                        str(json_object.get("connected_case", "")).replace(",", "."),
                    )
                    length = float(match_len.group()) if match_len else 0
                except Exception:
                    length = 0

                total_calculated_price = (
                    round(days_period * (unit_price * length), 2)
                    if days_period is not None and unit_price and length
                    else None
                )

                actual_price = detail.get("price", 0)

                price_match_status = (
                    "MATCH"
                    if total_calculated_price is not None
                    and abs(total_calculated_price - actual_price) <= 0.01
                    else "MISMATCH"
                )


                if price_match_status == "MISMATCH":
                    days_written = detail.get("units")
                    calculated_length = (
                        round(detail_unit_price / unit_price, 2) if unit_price else 0
                    )

                    # Længde/m2 issue
                    if length != calculated_length:
                        upsert_issue(
                            conn=conn,
                            case_id=case_id,
                            invoice_id=invoice_id_str,
                            issue_type="Længde/m2 stemmer ikke",
                            fakturalinje=matched_fakturalinje,
                            description=(
                                f"Længden/m2 er opgivet til {length}, men ud fra fakturalinjen "
                                f"udregnes længden/m2 til at være {calculated_length} hvis "
                                f"enhedsprisen er {unit_price}."
                            ),
                            fix=(
                                'Ret længden/m2 i feltet "Relateret sag", eller justér fakturalinjen '
                                "så den matcher den korrekte længde/m2. Sørg for kun at have selve længden "
                                "eller m2-værdien stående - ikke udregninger."
                            ),
                            caseworker_email=caseworker_email,
                            inserted_to_kassen="Yes",  # line will be inserted i Vejmankassen
                        )
                        current_issue_keys.add((invoice_id_str, "Længde/m2 stemmer ikke"))

                    # Antal dage issue
                    if days_period != days_written:
                        if chosen_end_date != end_date:
                            desc = (
                                f"Antal dage er angivet til {days_written} i fakturalinjen, "
                                f"men ud fra startdato {start_for_calc.strftime('%d-%m-%Y')} "
                                f"og færdigmeldingsdato {chosen_end_date.strftime('%d-%m-%Y')} "
                                f"udregnes {days_period} dage. Færdigmeldingsdatoen benyttes, "
                                f"da den er før slutdatoen {end_date.strftime('%d-%m-%Y')}."
                            )
                        else:
                            desc = (
                                f"Antal dage er angivet til {days_written}, men ud fra startdato "
                                f"{start_for_calc.strftime('%d-%m-%Y')} og slutdato "
                                f"{end_date.strftime('%d-%m-%Y')} udregnes {days_period} dage."
                            )

                        upsert_issue(
                            conn=conn,
                            case_id=case_id,
                            invoice_id=invoice_id_str,
                            issue_type="Antal dage stemmer ikke",
                            fakturalinje=matched_fakturalinje,
                            description=desc,
                            fix=(
                                "Ret antal dage i fakturalinjen og/eller ret datoer i Vejmankassen."
                            ),
                            caseworker_email=caseworker_email,
                            inserted_to_kassen="Yes",
                        )
                        current_issue_keys.add((invoice_id_str, "Antal dage stemmer ikke"))

                # ------------------------
                # Determine TilladelsesType to store
                # ------------------------
                tilladelsestype_value = matched_fakturalinje

                # Dates for DB insert/update
                short_start_date = (
                    start_for_calc.strftime("%Y-%m-%d") if start_for_calc else None
                )
                short_end_date = (
                    chosen_end_date.strftime("%Y-%m-%d") if chosen_end_date else None
                )

                # ------------------------
                # MERGE into VejmanFakturering
                # (still commented out so you can toggle it manually)
                # ------------------------
                merge_query = """
                MERGE INTO [dbo].[VejmanFakturering] AS target
                USING (SELECT ? AS VejmanFakturaID) AS source
                ON target.VejmanFakturaID = source.VejmanFakturaID
                WHEN MATCHED THEN
                    UPDATE SET 
                        Ansøger         = ?, 
                        FørsteSted      = ?, 
                        Tilladelsesnr   = ?, 
                        CvrNr           = ?, 
                        TilladelsesType = ?,
                        Enhedspris      = ?, 
                        Meter           = ?, 
                        Startdato       = ?, 
                        Slutdato        = ?,
                        ATT             = ?
                WHEN NOT MATCHED THEN
                    INSERT (
                        VejmanID, Ansøger, FørsteSted, Tilladelsesnr, CvrNr,
                        TilladelsesType, Enhedspris, Meter, Startdato, Slutdato,
                        VejmanFakturaID, ATT, FakturaStatus
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                """

                with conn.cursor() as cur:
                    cur.execute(
                        merge_query,
                        (
                            # Source (for MATCH test)
                            vejman_faktura_id,
                            # UPDATE (if exists)
                            applicant,
                            address,
                            tilladelse_nr,
                            cvr_number,
                            tilladelsestype_value,
                            unit_price,
                            length,
                            short_start_date,
                            short_end_date,
                            att,
                            # INSERT (if not exists)
                            case_id,
                            applicant,
                            address,
                            tilladelse_nr,
                            cvr_number,
                            tilladelsestype_value,
                            unit_price,
                            length,
                            short_start_date,
                            short_end_date,
                            vejman_faktura_id,
                            att,
                            "Ny",
                        ),
                    )


            # ------------------------
            # Auto-resolve issues that have disappeared for this case
            # ------------------------
            auto_resolve_missing_issues(conn, case_id, current_issue_keys)
            conn.commit()

        orchestrator_connection.log_info("Finished processing all cases and updating InvoiceIssues.")


def is_valid_cvr_mod11(cvr: str) -> bool:
    if not re.fullmatch(r"\d{8}", cvr):
        return False

    weights = [2, 7, 6, 5, 4, 3, 2, 1]
    total = sum(int(cvr[i]) * weights[i] for i in range(8))
    return total % 11 == 0
