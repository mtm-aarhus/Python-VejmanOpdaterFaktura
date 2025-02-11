# Debug in sandbox.py, make sure you set environment variables in CMD (ADMIN) for OpenOrchestrator:
# setx OpenOrchestratorKey key /M
# setx OpenOrchestratorSQL connectionstring /M

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import requests
import os
import random
import re
import time
import string
import pandas as pd
import locale
import smtplib
from email.message import EmailMessage
from datetime import datetime
from robot_framework import config
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import pyodbc

def FetchVejmanPermissions(token, equipment_type, fra_startdato, fra_slutdato, orchestrator_connection: OrchestratorConnection):

    combined_cases = []

    # Initialize an HTTP client
    with requests.Session() as client:
        # Modify the URL for each equipment type
        url = f"https://vejman.vd.dk/permissions/getcases?pmCaseStates=8&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Cstart_date%2Cstreet_name%2Ccvr_number%2Capplicant%2Cend_date%2Ccompletion_date%2Cauto_completedcontractor%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27rovm%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseTagShow=&pmCaseShowAttachments=false&pmAllStates=&dontincludemap=1&authority=751&cse=&equipmentType={equipment_type}&startDateFrom={fra_startdato}&startDateTo={datetime.today().strftime('%Y-%m-%d')}&endDateFrom={fra_slutdato}&policeDistrictShow=&_={int(time.time() * 1000)}&token={token}"

        response = client.get(url, timeout = 500)
        response.raise_for_status()
        json_object = response.json()

        cases = json_object["cases"]

        # Combine the cases into one list
        combined_cases.extend(cases)

    # Assuming combined_cases is not empty
    if combined_cases:
        # Create a DataFrame from the combined cases
        data_frame = pd.DataFrame(combined_cases)
    else:
        print("No new permissions")
        data_frame = pd.DataFrame()
    return data_frame

def FetchPricebookData(token):
    url = f"https://vejman.vd.dk/services/data.do?table=v_h_pm_pricebook&token={token}"
    response = requests.get(url)
    response.raise_for_status()
    pricebook_data = response.json().get('data', [])
    pricebook_map = {item['text']: item for item in pricebook_data}
    return pricebook_map


def append_to_mail_body(mail_body, append_text):
    """
    Appends text to mail_body with a <br> if mail_body already has content.
    
    Args:
        mail_body (str): The current mail body content.
        append_text (str): The text to append.

    Returns:
        str: The updated mail body.
    """
    if len(mail_body) > 0:
        mail_body += "<br><br>"
    mail_body += append_text
    return mail_body

def SendEmail(to_address: str | list[str], subject: str, body: str, bcc: str):
    msg = EmailMessage()
    msg['to'] = to_address
    msg['from'] = "VejmanFakturaRobot <noreply@aarhus.dk>"
    msg['subject'] = subject
    msg['bcc'] = bcc

    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(body, subtype='html')

    # Send message
    with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.send_message(msg)


def FetchInvoice(data_frame: pd.DataFrame, token, pricebook_map, equipment_type, fakturalinjer, conn, faktura_db, developer_email, orchestrator_connection: OrchestratorConnection):
    locale.setlocale(locale.LC_NUMERIC, 'da_DK')
    
    with requests.Session() as client:
        for index, row in data_frame.iterrows():
            mail_body = ""
            # Extract necessary variables
            case_id = row['case_id']
            case_number = row['case_number']
            print(f"Checking {case_number} - https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}")            

            start_date = datetime.strptime(row.get('start_date', ''), "%d-%m-%Y")
            end_date = datetime.strptime(row.get('end_date', ''), "%d-%m-%Y")
            completion_date = datetime.strptime(row.get('completion_date', ''), "%d-%m-%Y")
            auto_completed = row.get('auto_completed')
            #cvr_number = row.get('cvr_number')
            applicant = row.get('applicant')
            tilladelse_nr = case_number
            address = row['street_name']

            
            # Fetch detailed case data
            response = client.get(f"https://vejman.vd.dk/permissions/getcase?caseid={case_id}&token={token}", timeout=500)
            response.raise_for_status()
            json_object = response.json().get('data')
            caseworker_email = json_object['authEmail']

            # Check if there's an invoice in the JSON object
            invoice_data = json_object.get('invoice', {})
            
            if invoice_data:         
                # Get invoice role, select 1 (ansøger) if no role selected
                invoice_role_id = invoice_data.get('role', {}).get('id', 1)
                att = "Intet navn angivet"
                # Get name for ATT and update the cvr_number if not found in the main DataFrame
                contacts = json_object.get('contacts', [])
                for contact in contacts:
                    # Check if this contact has a role matching the invoice role id
                    roles = contact.get("roles", [])
                    if any(role.get("role", {}).get("id") == invoice_role_id for role in roles):
                        # Combine name components
                        name_parts = [
                            contact.get("given_name", ""),
                            contact.get("middle_name", ""),
                            contact.get("surname", ""),
                        ]
                        combined_name = " ".join(part for part in name_parts if part)
                        if combined_name:
                            att = f"Att: {combined_name}"  # Set att with the combined name

                        cvr_number = contact.get("cvr_number")
                        break  # Exit loop once the matching contact is processed

                if not cvr_number:
                    print("Intet CVR nummer")
                    mail_body = append_to_mail_body(mail_body, f'Der er intet CVR nummer angivet for faktura modtager på tilladelse <a href="https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}">{tilladelse_nr}</a>.')
                    cvr_number = '00000000'
                else:
                    if re.fullmatch(r'\d{8}', cvr_number) is None:
                        print("Forkert angivet CVR nummer")
                        mail_body = append_to_mail_body(mail_body, f'CVR nummer er angivet som {cvr_number} for faktura modtager på tilladelse <a href="https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}">{tilladelse_nr}</a>, men det burde være udelukkende 8 cifre. Check venligst om det er angivet korrekt og om der er skjulte tegn eller mellemrum.')
                        cvr_number = '00000000'
                invoice_details = invoice_data.get('details', [])
                # Iterate through each invoice detail and extract relevant information
                Matches = False
                AlreadyCreated = False
                for detail in invoice_details:
                    
                    detail_text = detail.get('text')
                    # Check if detail_text exists in the Fakturalinje column
                    VejmanFakturaID = detail.get('id')
                    
                    # Access columns by name
                    matching_row = next(
                        (row for row in faktura_db if row.VejmanFakturaID == VejmanFakturaID),
                        None
                    )        
                    if matching_row and (matching_row.Faktureret == 1 or matching_row.SendTilFakturering == 1 or matching_row.FakturerIkke == 1):
                        print("Row already sent to invoice, deleted or has been invoiced, skipping.")
                        continue           
                    
                    matched_fakturalinje = None
                    
                    for f in fakturalinjer.split(','):
                        if f.strip().lower() in detail_text.strip().lower():
                            print(f"Match found for Fakturalinje: {detail_text} with MaterielIDVejman = {equipment_type}: {f}")
                            Matches = True
                            matched_fakturalinje = f  # Assign the matching Fakturalinje
                            break  # Break inner loop when match is found

                    if matched_fakturalinje:
                        # Do something with the matched Fakturalinje
                        Fakturalinje = matched_fakturalinje
                    else:
                        print(f"No match found for Fakturalinje: {detail_text} with MaterielIDVejman = {equipment_type}")
                        continue  # Continue to next detail if no match
                    pricebook_entry = pricebook_map.get(detail_text, {})
                    
                    if matching_row:
                        AlreadyCreated = True
                        print("Row already exists, updating without replacing dates")
                        # Fetch Startdato and Slutdato from the SQL row if it exists
                        start_date = datetime.strptime(matching_row.Startdato, "%Y-%m-%d")
                        end_date = datetime.strptime(matching_row.Slutdato, "%Y-%m-%d")
                        chosen_end_date = end_date
                        completion_date = end_date
                    else:
                        print("No matching row")
                        # Check if autocompleted, if so then end_date, if not check if completion_date is lesser than end_date, else use end_date
                        chosen_end_date = end_date if auto_completed == "AF" else min(completion_date, end_date) if completion_date and end_date else end_date
                    
                    
                    if start_date and chosen_end_date:
                        # Convert both to date objects, ignoring time part
                        days_difference = (chosen_end_date.date() - start_date.date()).days
                        
                        # Add 1 to count both start and end dates (as in the example provided)
                        days_period = days_difference + 1
                    else:
                        days_period = None


                    # Handle cases where unit_price can be a string or a float
                    raw_detail_unit_price = detail.get('unit_price', 0)
                    if isinstance(raw_detail_unit_price, str):
                        raw_detail_unit_price = float(raw_detail_unit_price.replace(",", "."))
                    detail_unit_price = raw_detail_unit_price

                    unit_price = pricebook_entry.get('unit_price', 0)
                    unit_price = float(unit_price)

                    try:
                        match = re.search(r'\d+(\.\d+)?', str(json_object['connected_case']).replace(",","."))
                        length = float(match.group()) if match else 0
                    except:
                        length = 0
                    
                    total_calculated_price = round(days_period * (unit_price * length),2) if days_period is not None else None
                    
                    # Compare the calculated price with the actual price in the detail
                    price_match_status = "MATCH" if total_calculated_price and abs(total_calculated_price - detail.get('price', 0)) <= 0.01 else "MISMATCH"
                    
                    if price_match_status == 'MISMATCH':
                        days_written = detail.get('units')
                        calculated_length = round(detail_unit_price / unit_price,2) if unit_price else 0
                        mail_body = append_to_mail_body(mail_body, f'Der er uoverensstemmelse mellem de angivne værdier på tilladelse <a href="https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}">{tilladelse_nr}</a> for fakturalinjnen med teksten {Fakturalinje}. Robotten har opdaget følgende:')
                        print(str(total_calculated_price)+" - "+str(detail.get('price', 0))+f" {length} <> {calculated_length}, {days_period} <> {days_written}")
                        if length != calculated_length:
                            mail_body = append_to_mail_body(mail_body, f'Længden/m2 er opgivet til {length}, men ud fra fakturalinjen udregnes længden/m2 til at være {calculated_length} hvis enhedsprisen er på {unit_price}. Du skal derfor rette fakturalinjen eller sørge for at længden/m2 er angivet korrekt i "Relateret sag" feltet. Sørg for kun at have længden eller m2 værdien stående i relateret sag for at robotten kan læse det korrekt, og f.eks. ikke udregningen af kvadratmeter. Hvis der er flere fakturalinjer på tilladelsen med forskellige længder må du rette dem til i Vejmankassen når du sender dem til fakturering.')
                        if days_period != days_written:
                            if chosen_end_date != end_date:
                                mail_body = append_to_mail_body(mail_body, f'Antal af dage er angivet til {days_written} i fakturalinjen, men ud fra startdato og færdigmeldingsdato udregnes antallet af dage fra {start_date.strftime("%d-%m-%Y")} til og med {chosen_end_date.strftime("%d-%m-%Y")} til at være {days_period} dage. Færdigmeldingsdatoen {chosen_end_date.strftime("%d-%m-%Y")} benyttes da den er angivet til at være færdig før slutdatoen som er sat til {end_date.strftime("%d-%m-%Y")}.')
                            else:
                                mail_body = append_to_mail_body(mail_body, f'Antal af dage er angivet til {days_written} i fakturalinjen, men ud fra startdato og slutdato udregnes antallet af dage fra {start_date.strftime("%d-%m-%Y")} til og med {end_date.strftime("%d-%m-%Y")} til at være {days_period} dage')
                        mail_body = append_to_mail_body(mail_body, f'Du har fået tilsendt denne mail da du står som sagsbehandler på sagen inde i Vejman. Tilladelsen er angivet som værende type {equipment_type} under Materiel - med følgende fakturatekst: {Fakturalinje}.')

                    
                    short_start_date = start_date.strftime('%Y-%m-%d')
                    short_end_date = chosen_end_date.strftime('%Y-%m-%d')
                    
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
                            Enhedspris, Meter, Startdato, Slutdato, VejmanFakturaID, ATT
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                    """

                    # Execute the query
                    with conn.cursor() as cursor:
                        cursor.execute(
                            merge_query,
                            (
                                # For source (to check if exists)
                                VejmanFakturaID,
                                
                                # For update (if exists)
                                applicant, address, tilladelse_nr, cvr_number, Fakturalinje, 
                                unit_price, length, short_start_date, short_end_date, att,
                                
                                # For insert (if not exists)
                                case_id, applicant, address, tilladelse_nr, cvr_number, Fakturalinje, 
                                unit_price, length, short_start_date, short_end_date, VejmanFakturaID, att
                            )
                        )
                        conn.commit()


                    # Call the update_case function to send the data and verify the response
                    # update_case(filtered_data, token)
                if Matches == False:
                    print(f"No invoice line matches for {tilladelse_nr}")
                    continue
                    #mail_body = append_to_mail_body(mail_body, f'Der var intet match for følgende fakturalinje tekst i vejman: {detail_text} hvis materieltypen er {Fakturalinje}. Dette kan være fordi der er flere materieltyper på vejman tilladelsen <a href="https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}">{tilladelse_nr}</a>, men tjek venligst efter om alt ser korrekt ud.')
                if len(mail_body) > 0 and AlreadyCreated == False:
                    mail_body = f'''Der er fundet uoverensstemmelser på fakturalinje(r) for tilladelse <a href="https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}">{tilladelse_nr}</a>. Ret dem til inde i Vejman, så bliver de automatisk opdateret i Vejmankassen næste dag medmindre de er slettet eller sendt til fakturering. Hvis datoerne er forkerte eller der er flere fakturalinjer pr. tilladelse skal de opdateres i <a href="https://vejmankassen.adm.aarhuskommune.dk/">Vejmankassen</a>. For at undgå spam får du kun denne mail en gang pr. fakturalinje, så du skal selv tjekke op på om alt er korrekt før du sender den til fakturering.<br><br>'''+mail_body
                    SendEmail(caseworker_email,f"Uoverensstemmelser for fakturering på tilladelse {tilladelse_nr}", mail_body, developer_email)
            else:
                print(f"No invoices found for case ID: {case_id}")
        
orchestrator_connection = OrchestratorConnection("VejmanOpusFakturering", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)


token = orchestrator_connection.get_credential("VejmanToken").password
pricebook_map = FetchPricebookData(token)
developer_email = orchestrator_connection.get_constant("JADT").value

sql_server = orchestrator_connection.get_constant("SqlServer")
conn_string = "DRIVER={SQL Server};"+f"SERVER={sql_server.value};DATABASE=PYORCHESTRATOR;Trusted_Connection=yes;"
conn = pyodbc.connect(conn_string)
cursor = conn.cursor()

# Make this as a single process dispatcher so multiple processes can run
# Fetch all rows from the table
query = """SELECT 
    MaterielIDVejman, 
    STRING_AGG(Fakturalinje, ',') AS Fakturalinjer, 
    MIN(FraStartdato) AS EarliestStartDate,
    MIN(FraSlutdato) AS EarliestSlutDate
FROM [dbo].[VejmanFakturaTekster]
GROUP BY MaterielIDVejman
"""
cursor.execute(query)
rows = cursor.fetchall()

query = """SELECT * FROM [PyOrchestrator].[dbo].[VejmanFakturering]"""
cursor.execute(query)
faktura_db = cursor.fetchall()

for row in rows:
    fakturalinjer = row.Fakturalinjer
    eq_type = row.MaterielIDVejman
    start_date = row.EarliestStartDate
    from_end_date = row.EarliestSlutDate
    
    equipment_types = [eq_type]  # Use only the current equipment_type
    
    # Check if equipment_type is 5
    if eq_type == 1:
        equipment_types = [1, 9]  # List of equipment types to iterate over

    if eq_type == 2:
        equipment_types = [2, 7]  # List of equipment types to iterate over
        
    # Iterate through the relevant equipment types
    for equipment_type in equipment_types:
        # Fetch permissions data
        data_frame = FetchVejmanPermissions(token, equipment_type, start_date, from_end_date, orchestrator_connection)

        if data_frame.empty:
            print(f'Ingen rækker for {equipment_type} fra startdato {start_date} og fra slutdato {from_end_date}')
            continue
        # Clean authority_reference_number column
        data_frame['cleaned_authority_reference_number'] = data_frame['authority_reference_number'].apply(
            lambda x: re.sub(r'[^\x20-\x7E]', '', str(x).strip().lower()) if pd.notnull(x) else ''
        )

        # Filter rows based on substring checks for 'faktura sendt' and 'faktureres ikke', and exact match for 'fak'
        filtered_rows = data_frame[
            ~(
                data_frame['cleaned_authority_reference_number'].str.contains('faktura sendt') | 
                data_frame['cleaned_authority_reference_number'].str.contains('faktureres ikke') | 
                data_frame['cleaned_authority_reference_number'].str.contains('annulleret') | 
                (data_frame['cleaned_authority_reference_number'] == 'fak')
            ) & 
            (data_frame['initials'] != 'JADT')
        ]


        # Fetch invoices for filtered rows
        FetchInvoice(filtered_rows, token, pricebook_map, equipment_type, fakturalinjer, conn, faktura_db, developer_email, orchestrator_connection)

orchestrator_connection.update_constant("VejmanKassenSynkroniseret", datetime.now().strftime("%d-%m-%Y %H:%M"))