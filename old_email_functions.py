import smtplib
from email.message import EmailMessage
from robot_framework import config


def build_case_email_body(
    case_id: str,
    case_number: str,
    mismatch_issues: list[dict],
) -> str:
    """
    Bygger en samlet HTML-mail pr. tilladelse med:
    - tabel over uoverensstemmelser (MATCH men fejl i længde/dage/pris)
    - tabel over fakturalinjer uden matchende materiel
    """
    case_url = f"https://vejman.vd.dk/permissions/update.jsp?caseid={case_id}"

    intro = (
        f'Denne mail vedrører tilladelse <a href="{case_url}">{case_number}</a>.<br>'
        f"Robotten har fundet nogle forhold, som du skal tjekke og evt. rette:"
    )

    html_parts = [intro, "<br><br>"]

    # Mismatch table
    if mismatch_issues:
        html_parts.append("<h3>Uoverensstemmelser</h3>")
        html_parts.append(
            "<p>Robotten har fundet uoverensstemmelser. Ret fejlen enten i Vejman eller i Vejmankassen, som beskrevet.</p>"
        )
        html_parts.append(
            "<table border='1' cellpadding='5' cellspacing='0'>"
            "<tr>"
            "<th>Fakturalinje</th>"
            "<th>Fejl</th>"
            "<th>Sådan retter du</th>"
            "</tr>"
        )

        for issue in mismatch_issues:
            fakturalinje = issue["fakturalinje"]
            detail_text = issue["detail_text"]
            for i in issue["issues"]:
                html_parts.append(
                    "<tr>"
                    f"<td>{detail_text or fakturalinje}</td>"
                    f"<td>{i['type']}<br>{i['description']}</td>"
                    f"<td>{i['fix']}</td>"
                    "</tr>"
                )

        html_parts.append("</table><br>")

    html_parts.append(
        "<p>For at undgå spam sendes denne mail kun første gang robotten opdager disse linjer "
        "for den pågældende tilladelse. Tjek derfor, at alt er korrekt, før du sender sagen til fakturering.</p>"
    )

    return "".join(html_parts)

def append_to_mail_body(mail_body, append_text):
    if len(mail_body) > 0:
        mail_body += "<br><br>"
    mail_body += append_text
    return mail_body


def SendEmail(to_address: str | list[str], subject: str, body: str, bcc: str):
    msg = EmailMessage()
    msg["to"] = to_address
    msg["from"] = "VejmanFakturaRobot <noreply@aarhus.dk>"
    msg["subject"] = subject
    msg["bcc"] = bcc

    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(body, subtype="html")

    with smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.send_message(msg)
