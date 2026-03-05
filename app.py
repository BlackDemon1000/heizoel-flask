from flask import Flask, render_template, request, jsonify
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

EXCEL_FILE = "orders.xlsx"

# -------------------------------------------------
# Excel initialisieren
# -------------------------------------------------
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append([
            "Zeit",
            "Bestellnummer",
            "Anrede", "Vorname", "Nachname",
            "E-Mail", "Telefon",
            "Straße", "Hausnr", "PLZ", "Ort",
            "Menge (Liter)", "Tankart", "Einfüllstutzen",
            "Bemerkung",
            "Status"
        ])
        wb.save(EXCEL_FILE)

def save_to_excel(order, order_id):
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        order_id,
        order.get("anrede"),
        order.get("vorname"),
        order.get("nachname"),
        order.get("email"),
        order.get("telefon"),
        order.get("strasse"),
        order.get("hausnr"),
        order.get("plz"),
        order.get("ort"),
        order.get("menge"),
        order.get("tankart"),
        order.get("einfuellstutzen"),
        order.get("bemerkung", ""),
        "WARTET AUF BESTÄTIGUNG"
    ])
    wb.save(EXCEL_FILE)

# -------------------------------------------------
# Bestätigungs-E-Mail senden
# -------------------------------------------------
""")def send_confirmation_mail(order, order_id):
    msg = EmailMessage()
    msg["From"] = os.getenv("MAIL_USER")
    msg["To"] = order["email"]
    msg["Subject"] = f"Heizölbestellung – Bitte bestätigen (Nr. {order_id})"

    msg.set_content(f
Sehr geehrte Damen und Herren,

vielen Dank für Ihre Heizölbestellung.
Bitte prüfen Sie die Angaben und bestätigen Sie per Antwort mit:

BESTÄTIGEN

────────────────────────────
🧾 BESTELLDETAILS
────────────────────────────
Bestellnummer: {order_id}

Menge:            {order["menge"]} Liter
Tankart:          {order["tankart"]}
Einfüllstutzen:   {order["einfuellstutzen"]}

────────────────────────────
📍 LIEFERADRESSE
────────────────────────────
{order["anrede"]} {order["vorname"]} {order["nachname"]}
{order["strasse"]} {order["hausnr"]}
{order["plz"]} {order["ort"]}

────────────────────────────
📞 KONTAKT
────────────────────────────
E-Mail:   {order["email"]}
Telefon: {order["telefon"]}

────────────────────────────
📝 BEMERKUNG
────────────────────────────
{order.get("bemerkung", "—")}

Ohne Bestätigung wird die Bestellung nicht ausgeführt.

Mit freundlichen Grüßen
Ihr Heizöl-Service
)

    with smtplib.SMTP_SSL(os.getenv("SMTP_SERVER"), 465) as server:
        server.login(os.getenv("MAIL_USER"), os.getenv("MAIL_PASS"))
        server.send_message(msg)
)"""
# -------------------------------------------------
# Routes
# -------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/submit", methods=["POST"])
def submit():
    order = request.json
    order_id = f"BO-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    save_to_excel(order, order_id)
    #send_confirmation_mail(order, order_id)

    return jsonify({
        "status": "ok",
        "order_id": order_id
    })

# -------------------------------------------------
if __name__ == "__main__":
    init_excel()
    app.run(debug=True)