from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from old_email_functions import build_case_email_body, SendEmail

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


def FetchVejmanPermissions(token, equipment_type, fra_startdato, fra_slutdato, orchestrator_connection: OrchestratorConnection):
    combined_cases = []
    
    with requests.Session() as client:
        url = (
            "https://vejman.vd.dk/permissions/getcases"
            f"?pmCaseStates=8"
            "&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cstart_date%2Cstreet_name%2Ccvr_number"
            "%2Capplicant%2Cend_date%2Ccompletion_date%2Cauto_completedcontractor%2Cinitials"
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



# ------------------------------
# Core: process all unique cases
# ------------------------------


def ProcessCases(
    cases_by_id: dict[str, pd.Series],
    case_materiel_ids: dict[str, set[int]],
    materiel_config: dict[int, dict],
    token,
    pricebook_map,
    conn,
    faktura_db_by_vejman_id: dict,
    developer_email: str,
    orchestrator_connection: OrchestratorConnection,
):
    locale.setlocale(locale.LC_NUMERIC, "da_DK")
    fakturalinje_to_materiel = {}
    for mid, conf in materiel_config.items():
        for fl in conf["fakturalinjer"]:
            f_clean = fl.strip().lower()
            fakturalinje_to_materiel[f_clean] = mid

    triggered_urls = []          # list[str]  → for TXT file
    issues_for_excel = []   
    with requests.Session() as client:
        for case_id, case_row in cases_by_id.items():
            case_number = case_row["case_number"]

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
            caseworker_email = json_object["authEmail"]

            # Invoice block
            invoice_data = json_object.get("invoice")
            if not invoice_data:
                # No invoice block at all → skip case
                continue

            invoice_details = invoice_data.get("details", [])
            if not invoice_details:
                # Invoice exists but contains no invoice lines → skip case
                continue

            # Determine invoice contact (ATT)
            invoice_role_id = invoice_data.get("role", {}).get("id", 1)
            att = "Intet navn angivet"
            contacts = json_object.get("contacts", [])


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
            if contact_type == "P" and (not cvr_number or str(cvr_number).strip() == ""):
                continue
            # Lists for email aggregation per case
            mismatch_issues = []        # [{...}, ...]
            unmatched_lines_email = []  # [{...}, ...]

            if not cvr_number or not re.fullmatch(r"\d{8}", str(cvr_number)):
                mismatch_issues.append({
                    "detail_text": "CVR-nummer",
                    "fakturalinje": "CVR ugyldigt",
                    "issues": [{
                        "type": "CVR mangler eller ugyldigt.",
                        "description": (
                            "CVR-nummeret mangler eller består ikke af præcis 8 cifre."
                        ),
                        "fix": (
                            "Ret CVR-nummeret i kontaktoplysningerne i Vejman på den interessent der bliver faktureret. "
                            "Robotten har midlertidigt sat CVR = 00000000."
                        ),
                    }],
                })
                cvr_number = "00000000"

            else:
                cvr_str = str(cvr_number)
                if not is_valid_cvr_mod11(cvr_str):
                    mismatch_issues.append({
                        "detail_text": "CVR-nummer",
                        "fakturalinje": "CVR ugyldigt (modulus 11)",
                        "issues": [{
                            "type": "CVR modulus 11 fejlede",
                            "description": (
                                f"CVR-nummeret {cvr_str} består af 8 cifre, "
                                "men opfylder ikke modulus-11 kontrollen og er derfor ikke et gyldigt CVR."
                            ),
                            "fix": (
                                "Ret CVR-nummeret i kontaktoplysningerne i Vejman. "
                            ),
                        }],
                    })

            tilladelse_nr = case_number

            # Pre-calc allowed fakturalinjer + materiel IDs for this case
            materiel_ids_for_case = sorted(case_materiel_ids.get(case_id, set()))
            allowed_fakturalinjer = []
            for mid in materiel_ids_for_case:
                conf = materiel_config.get(mid)
                if conf:
                    allowed_fakturalinjer.extend(conf["fakturalinjer"])

            # Process each invoice line for this case
            for detail in invoice_details:
                detail_text = detail.get("text", "") or ""
                VejmanFakturaID = detail.get("id")

                # Lookup existing faktura row for this invoice line
                matching_row = faktura_db_by_vejman_id.get(VejmanFakturaID)

                # If row exists and is not 'Ny', skip completely
                if matching_row and getattr(matching_row, "FakturaStatus", None) != "Ny":
                    continue

                matched_fakturalinje = None
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

                if not text_is_known:
                    mismatch_issues.append({
                        "detail_text": detail_text,
                        "fakturalinje": "Ingen match",
                        "issues": [{
                            "type": "Fakturalinje uden match",
                            "description": (
                                "Fakturalinjen er ikke godkendt til brug i Vejmankassen. "
                                "Robotten kan derfor ikke behandle denne linje."
                            ),
                            "fix": (
                                "Ret fakturalinjen i Vejman til en gyldig fakturerbar linje. "
                                "Denne fakturalinje bliver *ikke* tilføjet til Vejmankassen, og "
                                "du vil fortsætte med at modtage denne mail indtil det er rettet, "
                                "medmindre du sætter ESDH-journalnr. til faktureres ikke"
                            ),
                        }],
                    })

                    # Skip DB insertion completely
                    continue
                if text_is_known and not materiel_is_allowed:
                    mismatch_issues.append({
                        "detail_text": detail_text,
                        "fakturalinje": matched_fakturalinje,
                        "issues": [{
                            "type": "Manglende/forkert materiel",
                            "description": (
                                f"Fakturalinjen matcher materiel id {matched_materiel_id}, men "
                                "denne materieltype er ikke tilknyttet tilladelsen i Vejman. "
                                "Fakturalinjen er lagt ind i vejmankassen."
                            ),
                            "fix": (
                                f"Tilknyt materiel med ID {matched_materiel_id} i Vejman."
                            ),
                        }],
                    })

                # Determine dates / chosen_end_date
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
                    # Convert both to date objects, ignoring time part
                    days_difference = (chosen_end_date.date() - start_date.date()).days
                    
                    # Add 1 to count both start and end dates (as in the example provided)
                    days_period = days_difference + 1
                else:
                    days_period = None

                # Unit price + length / price calc
                raw_detail_unit_price = detail.get("unit_price", 0)
                if isinstance(raw_detail_unit_price, str):
                    raw_detail_unit_price = float(raw_detail_unit_price.replace(",", "."))

                detail_unit_price = raw_detail_unit_price

                pricebook_entry = pricebook_map.get(detail_text, {})
                unit_price = float(pricebook_entry.get("unit_price", 0))

                try:
                    match_len = re.search(r"\d+(\.\d+)?", str(json_object.get("connected_case", "")).replace(",", "."))
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

                # Build data for mismatch table if we have a match and mismatch and row is new (not yet in DB)
                if price_match_status == "MISMATCH" and not matching_row:
                    days_written = detail.get("units")
                    calculated_length = round(detail_unit_price / unit_price, 2) if unit_price else 0

                    issues = []
                    if length != calculated_length:
                        issues.append({
                            "type": "Længde/m2 stemmer ikke",
                            "description": (
                                f"Længden/m2 er opgivet til {length}, men ud fra fakturalinjen "
                                f"udregnes længden/m2 til at være {calculated_length} hvis enhedsprisen er {unit_price}."
                            ),
                            "fix": (
                                'Ret længden/m2 i feltet "Relateret sag", eller justér fakturalinjen '
                                "så den matcher den korrekte længde/m2. Sørg for kun at have selve længden "
                                "eller m2-værdien stående – ikke udregninger."
                            ),
                        })
                    if days_period != days_written:
                        if chosen_end_date != end_date:
                            issues.append({
                                "type": "Antal dage stemmer ikke",
                                "description": (
                                    f"Antal dage er angivet til {days_written} i fakturalinjen, men ud fra startdato "
                                    f"{start_for_calc.strftime('%d-%m-%Y')} og færdigmeldingsdato "
                                    f"{chosen_end_date.strftime('%d-%m-%Y')} udregnes {days_period} dage. "
                                    f"Færdigmeldingsdatoen benyttes, da den er før slutdatoen {end_date.strftime('%d-%m-%Y')}."
                                ),
                                "fix": (
                                    "Ret antal dage i fakturalinjen og/eller ret datoer i Vejmankassen"
                                ),
                            })
                        else:
                            issues.append({
                                "type": "Antal dage stemmer ikke",
                                "description": (
                                    f"Antal dage er angivet til {days_written}, men ud fra startdato "
                                    f"{start_for_calc.strftime('%d-%m-%Y')} og slutdato "
                                    f"{end_date.strftime('%d-%m-%Y')} udregnes {days_period} dage."
                                ),
                                "fix": (
                                    "Ret antal dage i fakturalinjen og/eller ret datoer i Vejmankassen"
                                ),
                            })
                    if issues:
                        mismatch_issues.append({
                            "detail_text": detail_text,
                            "fakturalinje": matched_fakturalinje,
                            "issues": issues,
                        })

                # Determine TilladelsesType to store
                tilladelsestype_value = matched_fakturalinje

                # Prepare dates for DB insert/update
                short_start_date = start_for_calc.strftime("%Y-%m-%d") if start_for_calc else None
                short_end_date = chosen_end_date.strftime("%Y-%m-%d") if chosen_end_date else None

                # Insert unmatched lines into DB anyway (if not present) and collect for email
                if not matching_row:
                    line_clean = detail_text.strip().lower()
                    suggested_mids = [
                        mid for fl, mid in fakturalinje_to_materiel.items()
                        if fl in line_clean
                    ]
                    unmatched_lines_email.append({
                        "detail_text": detail_text,
                        "vejman_faktura_id": VejmanFakturaID,
                        "suggested_materiel_ids": suggested_mids or [],
                        "price": actual_price,
                        "units": detail.get("units"),
                    })

                # MERGE into VejmanFakturering if either:
                # - row doesn't exist, or
                # - row exists but FakturaStatus == 'Ny'
                if (not matching_row) or (getattr(matching_row, "FakturaStatus", None) == "Ny"):
                    print("test")
                    merge_query = """
                    MERGE INTO [dbo].[VejmanFakturering] AS target
                    USING (SELECT ? AS VejmanFakturaID) AS source
                    ON target.VejmanFakturaID = source.VejmanFakturaID
                    WHEN MATCHED THEN
                        UPDATE SET 
                            Ansøger = ?, 
                            FørsteSted = ?, 
                            Tilladelsesnr = ?, 
                            CvrNr = ?, 
                            TilladelsesType = ?,
                            Enhedspris = ?, 
                            Meter = ?, 
                            Startdato = ?, 
                            Slutdato = ?,
                            ATT = ?
                    WHEN NOT MATCHED THEN
                        INSERT (
                            VejmanID, Ansøger, FørsteSted, Tilladelsesnr, CvrNr, TilladelsesType, 
                            Enhedspris, Meter, Startdato, Slutdato, VejmanFakturaID, ATT, FakturaStatus
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                    """

                    # with conn.cursor() as cur:
                    #     cur.execute(
                    #         merge_query,
                    #         (
                    #             # Source (for MATCH test)
                    #             VejmanFakturaID,
                    #             # UPDATE (if exists)
                    #             applicant,
                    #             address,
                    #             tilladelse_nr,
                    #             cvr_number,
                    #             tilladelsestype_value,
                    #             unit_price,
                    #             length,
                    #             short_start_date,
                    #             short_end_date,
                    #             att,
                    #             # INSERT (if not exists)
                    #             case_id,
                    #             applicant,
                    #             address,
                    #             tilladelse_nr,
                    #             cvr_number,
                    #             tilladelsestype_value,
                    #             unit_price,
                    #             length,
                    #             short_start_date,
                    #             short_end_date,
                    #             VejmanFakturaID,
                    #             att,
                    #             "Ny",
                    #         ),
                    #     )
                    #     conn.commit()

            # After all invoice lines for this case: send a single email (if needed)
            case_url = f"https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}"

            if (mismatch_issues or unmatched_lines_email):
                  # 1) Collect URL for TXT export
                triggered_urls.append(case_url)

                # 2) Collect issues for Excel export
                # Mismatch issues
                for issue in mismatch_issues:
                    for i in issue["issues"]:
                        issues_for_excel.append({
                            "TilladelseNr": tilladelse_nr,
                            "CaseID": case_id,
                            "URL": case_url,
                            "Fakturalinje": issue["fakturalinje"],
                            "IssueType": i["type"],
                            "Description": i["description"],
                            "Fix": i["fix"],
                        })


                mail_body = build_case_email_body(
                    case_id=case_id,
                    case_number=tilladelse_nr,
                    mismatch_issues=mismatch_issues,
                )
                subject = f"Uoverensstemmelser for fakturering på tilladelse {tilladelse_nr}"
                SendEmail(developer_email, subject, mail_body, developer_email)
    with open("triggered_tilladelser.txt", "w", encoding="utf-8") as f:
        for url in triggered_urls:
            f.write(url + "\n")

    # 2) Export EXCEL with issues
    if issues_for_excel:
        df = pd.DataFrame(issues_for_excel)
        df.to_excel("tilladelses_issues.xlsx", index=False)
        orchestrator_connection.log_info("Generated tilladelses_issues.xlsx and triggered_tilladelser.txt")
    else:
        orchestrator_connection.log_info("No issues found; no export files generated.")



def is_valid_cvr_mod11(cvr: str) -> bool:
    if not re.fullmatch(r"\d{8}", cvr):
        return False

    weights = [2, 7, 6, 5, 4, 3, 2, 1]
    total = sum(int(cvr[i]) * weights[i] for i in range(8))
    return total % 11 == 0
