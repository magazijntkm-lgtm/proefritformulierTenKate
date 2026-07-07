import streamlit as st
import pandas as pd
import os
import requests
import smtplib
from email.message import EmailMessage
from datetime import datetime, date
from streamlit_drawable_canvas import st_canvas
from PIL import Image
from fpdf import FPDF

# Configuratie van de webpagina
st.set_page_config(page_title="Proefrit Aanvraag", page_icon="🏍️", layout="centered")

# --- CUSTOM STYLING & LOGO ---
if os.path.exists("logo.png"):
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("logo.png", use_container_width=True)

# Verberg het standaard Streamlit menu en footer voor een strakke 'App' look
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
# -----------------------------

# Bestandslocaties & Mappen
OPSLAG_FILE = "ingevulde_formulieren.xlsx"
HANDTEKENING_DIR = "handtekeningen"
PDF_DIR = "pdf_formulieren"

# Zorg dat de mappen bestaan
os.makedirs(HANDTEKENING_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)

# Helper: Postcode + Huisnummer ophalen via PDOK (Overheid API)
def haal_adres_op(postcode, huisnummer):
    pc = postcode.replace(" ", "").upper()
    hn = str(huisnummer).strip()
    url = f"https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?q={pc}%20{hn}&fl=straatnaam,woonplaatsnaam"
    try:
        response = requests.get(url, timeout=3)
        data = response.json()
        if data['response']['numFound'] > 0:
            doc = data['response']['docs'][0]
            return doc.get('straatnaam', ''), doc.get('woonplaatsnaam', '')
    except Exception:
        pass
    return "", ""

# Helper: PDF Genereren (A4 Formaat)
def genereer_pdf(data_dict, sig_filename, pdf_filename):
    pdf = FPDF()
    pdf.add_page()
    
    # Gebruik standaard helvetica font
    pdf.set_font("helvetica", 'B', 16)
    pdf.cell(0, 10, "Aanvraagformulier Proefrit - Ten Kate Motoren", ln=True, align="C")
    pdf.ln(5)
    
    # Persoonsgegevens
    pdf.set_font("helvetica", 'B', 12)
    pdf.cell(0, 8, "1. Persoonsgegevens", ln=True)
    pdf.set_font("helvetica", '', 10)
    
    # Vaste breedte van 40 voor de labels
    pdf.cell(40, 6, "Datum & Tijd:"); pdf.cell(0, 6, data_dict["Datum_Tijd"], ln=True)
    pdf.cell(40, 6, "Naam:"); pdf.cell(0, 6, data_dict["Naam"], ln=True)
    pdf.cell(40, 6, "Geboortedatum:"); pdf.cell(0, 6, data_dict["Geboortedatum"], ln=True)
    pdf.cell(40, 6, "Adres:"); pdf.cell(0, 6, f"{data_dict['Straat']} {data_dict['Huisnummer']}", ln=True)
    pdf.cell(40, 6, "Postcode & Plaats:"); pdf.cell(0, 6, f"{data_dict['Postcode']} {data_dict['Woonplaats']}", ln=True)
    pdf.cell(40, 6, "E-mail:"); pdf.cell(0, 6, data_dict["Email"], ln=True)
    pdf.cell(40, 6, "Telefoon:"); pdf.cell(0, 6, data_dict["Telefoon"], ln=True)
    pdf.cell(40, 6, "Rijbewijsnummer:"); pdf.cell(0, 6, data_dict["Rijbewijsnummer"], ln=True)
    
    pdf.ln(5)
    
    # Motorgegevens
    pdf.set_font("helvetica", 'B', 12)
    pdf.cell(0, 8, "2. Motorgegevens", ln=True)
    pdf.set_font("helvetica", '', 10)
    pdf.cell(40, 6, "Merk:"); pdf.cell(0, 6, data_dict["Merk"], ln=True)
    pdf.cell(40, 6, "Type / Model:"); pdf.cell(0, 6, data_dict["Type"], ln=True)
    
    # Invul-lijntje voor de verkoper
    pdf.set_font("helvetica", 'B', 10)
    pdf.cell(40, 6, "Kenteken:"); pdf.cell(0, 6, "________________________________________", ln=True)
    
    pdf.ln(5)
    
    # Voorwaarden
    pdf.set_font("helvetica", 'B', 12)
    pdf.cell(0, 8, "3. Voorwaarden & Akkoord", ln=True)
    pdf.set_font("helvetica", '', 8)
    
    voorwaarden = """De gebruiker krijgt de motorfiets in bruikleen voor de genoemde periode. De motorfiets blijft eigendom van Ten Kate Motoren. De gebruiker dient de motorfiets uiterlijk op hierboven genoemde datum en tijdstip einde proefrit bij Ten Kate Motoren ingeleverd te hebben.

Het is niet toegestaan:
- Om welke reden dan ook, deze termijn te verlengen zonder schriftelijke toestemming van Ten Kate Motoren.
- Om buiten een straal van 35km, vanaf Nieuwleusen te rijden.
- Om te rijden op een afgesloten circuit.
- De proefrit buiten de grenzen van Nederland te laten plaatsvinden.

Per heel uur dat de motor te laat is ingeleverd, kan Ten Kate Motoren EUR 50,- in rekening brengen.

FH Kenteken: Bij het rijden met een FH kenteken dient de gebruiker zich, in verband met de wettelijke regeling daarvoor, te beperken tot het testen van de motor. Dit betekent o.a. dat het verboden is de motor op de openbare weg te parkeren of te gebruiken voor vervoer van personen/goederen. De FH platen zijn niet geldig in het buitenland.

Kentekendocumenten: De gebruiker dient alle bij aanvang van de proefrit overhandigde documenten bij einde proefrit te retourneren aan Ten Kate Motoren.

Vrijwaring: De gebruiker vrijwaart het bedrijf voor alle schade ontstaan ten gevolge van of tijdens het gebruik van de motor, zoals onder meer ten gevolge van boetes, overtredingen en/of inbeslagname en/of verbeurdverklaring van de motor, evenals voor aanspraken van derden verband houdend met het gebruik van de motor.

Verzekering: De motor is door Ten Kate Motoren verzekerd met: Een WA + Casco verzekering met een eigen risico van EUR 750,-."""
    
    pdf.multi_cell(0, 4, voorwaarden)
    
    pdf.ln(5)
    
    # Handtekening
    pdf.set_font("helvetica", 'B', 10)
    pdf.cell(0, 6, "Handtekening Klant:", ln=True)
    # Voeg de opgeslagen PNG in
    pdf.image(sig_filename, x=10, y=pdf.get_y(), w=60)
    
    pdf.output(pdf_filename)

# Helper: PDF e-mailen naar de verkoop
def stuur_email_met_pdf(pdf_filename, klant_naam):
    msg = EmailMessage()
    msg['Subject'] = f"Nieuwe Proefrit Aanvraag: {klant_naam}"
    msg['From'] = st.secrets["email"]["username"]
    msg['To'] = "verkoop@tenkatemotoren.nl"
    msg.set_content(f"Beste Verkoop-team,\n\nEr is zojuist een nieuw proefritformulier ingevuld door {klant_naam}.\nIn de bijlage vinden jullie de door de klant ondertekende PDF.\n\nDe data is tevens toegevoegd aan de Google Sheet.\n\nGroet,\nHet Proefrit Systeem")

    # Voeg de PDF toe als bijlage
    with open(pdf_filename, 'rb') as f:
        pdf_data = f.read()
    
    pdf_naam = os.path.basename(pdf_filename)
    msg.add_attachment(pdf_data, maintype='application', subtype='pdf', filename=pdf_naam)

    try:
        # Inloggen en versturen via SMTP
        with smtplib.SMTP_SSL(st.secrets["email"]["smtp_server"], st.secrets["email"]["port"]) as server:
            server.login(st.secrets["email"]["username"], st.secrets["email"]["password"])
            server.send_message(msg)
        return True, ""
    except Exception as e:
        return False, f"E-mail versturen mislukt: {e}"

# Data compleet verwerken en opslaan
def save_form_data(data_dict, signature_img, klant_naam):
    # 1. Stuur data naar Google Sheets Webhook
    WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbzyLzZeBcycLqBSSmTy5uQ1R73jhAcfPOVssqmLwvcb6sJiSKq-bkBp-S9lXbZn-pPc/exec"
    
    try:
        response = requests.post(WEBHOOK_URL, json=data_dict)
        if response.status_code != 200:
            return False, "Fout bij opslaan in de database (Google Sheets)."
    except Exception as e:
        return False, f"Verbindingsfout met database: {e}"

    # 2. Sla handtekening en PDF op in de cloud-server
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = "".join([c for c in klant_naam if c.isalpha() or c.isdigit()]).rstrip()
    sig_filename = f"{HANDTEKENING_DIR}/handtekening_{safe_name}_{timestamp}.png"
    
    img = Image.fromarray(signature_img.astype('uint8'), 'RGBA')
    img.save(sig_filename)

    pdf_filename = f"{PDF_DIR}/Proefritformulier_{safe_name}_{timestamp}.pdf"
    genereer_pdf(data_dict, sig_filename, pdf_filename)
    
    # Verstuur direct de email!
    email_succes, email_fout = stuur_email_met_pdf(pdf_filename, klant_naam)
    if not email_succes:
        return False, f"Data opgeslagen, maar {email_fout}"
    
    return True, ""


# Variabelen voor de automatische adres-invuller
if "straat" not in st.session_state: st.session_state.straat = ""
if "woonplaats" not in st.session_state: st.session_state.woonplaats = ""
if "last_lookup" not in st.session_state: st.session_state.last_lookup = ""

# --- START UI ---
st.title("🏍️ Aanmeldformulier Proefrit")
st.write("Vul onderstaande gegevens in om de proefrit te starten.")

# 1. Klantgegevens
st.header("1. Persoonsgegevens")
# De "new-password" hack om Chrome's agressieve autofill te slim af te zijn
naam = st.text_input("Volledige Naam", autocomplete="new-password")
geboortedatum = st.date_input("Geboortedatum", value=None, min_value=date(1900, 1, 1), max_value=date.today(), format="DD-MM-YYYY")

st.markdown("##### Adresgegevens")
col1, col2 = st.columns([1, 1])
with col1: postcode = st.text_input("Postcode (bijv. 1234AB)", autocomplete="new-password")
with col2: huisnummer = st.text_input("Huisnummer", autocomplete="new-password")

# Activeer de PDOK adreszoeker op de achtergrond
huidige_combinatie = f"{postcode}_{huisnummer}"
if postcode and huisnummer and huidige_combinatie != st.session_state.last_lookup:
    gevonden_straat, gevonden_woonplaats = haal_adres_op(postcode, huisnummer)
    if gevonden_straat and gevonden_woonplaats:
        st.session_state.straat = gevonden_straat
        st.session_state.woonplaats = gevonden_woonplaats
    st.session_state.last_lookup = huidige_combinatie
    st.rerun()

straat = st.text_input("Straat", key="straat", autocomplete="new-password")
woonplaats = st.text_input("Woonplaats", key="woonplaats", autocomplete="new-password")

st.markdown("##### Contact")
email = st.text_input("E-mailadres", autocomplete="new-password")
telefoon = st.text_input("Telefoonnummer", autocomplete="new-password")

# 1.B Rijbewijsgegevens
st.header("🪪 Rijbewijs")
st.write("Om te mogen proefrijden hebben wij je rijbewijsnummer nodig. Dit is het nummer van **10 cijfers** dat op de voorkant van je rijbewijs bij **item 5** staat.")

if os.path.exists("voorbeeld_rijbewijs.jpg"):
    st.image("voorbeeld_rijbewijs.jpg", caption="Kijk bij nummer 5 voor het rijbewijsnummer.", width="stretch" if st.__version__ >= '1.30' else None)
else:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/15/NLD_driving_license_2014_front.jpg/600px-NLD_driving_license_2014_front.jpg", caption="Voorbeeld: Het rijbewijsnummer staat bij nummer 5.", width="stretch" if st.__version__ >= '1.30' else None)

rijbewijsnummer = st.text_input("Rijbewijsnummer (Item 5)", autocomplete="new-password")
st.markdown("---")

# 2. Motorgegevens
st.header("2. Motorgegevens")
st.write("Vul in op welke motor je wilt proefrijden.")
gekozen_merk = st.text_input("Merk (bijv. Yamaha, Honda)", autocomplete="new-password")
gekozen_motor = st.text_input("Type / Model (bijv. MT-07, Fireblade)", autocomplete="new-password")
st.markdown("---")

# 3. Voorwaarden & Handtekening
st.header("3. Voorwaarden & Akkoord")

with st.expander("📄 Klik hier om de Algemene Voorwaarden te lezen"):
    st.markdown("""
    **Proefrit:** De gebruiker krijgt de motorfiets in bruikleen voor de genoemde periode. De motorfiets blijft eigendom van Ten Kate Motoren. De gebruiker dient de motorfiets uiterlijk op hierboven genoemde datum en tijdstip einde proefrit bij Ten Kate Motoren ingeleverd te hebben.  
    Het is niet toegestaan:
    - Om welke reden dan ook, deze termijn te verlengen zonder schriftelijke toestemming van Ten Kate Motoren.
    - Om buiten een straal van 35km, vanaf Nieuwleusen te rijden.
    - Om te rijden op een afgesloten circuit.
    - De proefrit buiten de grenzen van Nederland te laten plaatsvinden.  
    
    *Per heel uur dat de motor te laat is ingeleverd, kan Ten Kate Motoren € 50,- in rekening brengen.*

    **FH Kenteken:** Bij het rijden met een FH kenteken dient de gebruiker zich, in verband met de wettelijke regeling daarvoor, te beperken tot het testen van de motor. Dit betekent o.a. dat het verboden is de motor op de openbare weg te parkeren of te gebruiken voor vervoer van personen/goederen. De FH platen zijn niet geldig in het buitenland.

    **Kentekendocumenten:** De gebruiker dient alle bij aanvang van de proefrit overhandigde documenten bij einde proefrit te retourneren aan Ten Kate Motoren.

    **Vrijwaring:** De gebruiker vrijwaart het bedrijf voor alle schade ontstaan ten gevolge van of tijdens het gebruik van de motor, zoals onder meer ten gevolge van boetes, overtredingen en/of inbeslagname en/of verbeurdverklaring van de motor, evenals voor aanspraken van derden verband houdend met het gebruik van de motor.

    **Verzekering:** De motor is door Ten Kate Motoren verzekerd met: **Een WA + Casco verzekering met een eigen risico van € 750,-**
    """)

akkoord = st.checkbox("Ik ga akkoord met de algemene voorwaarden.")

st.write("**Plaats hieronder je handtekening:**")
canvas_result = st_canvas(
    fill_color="rgba(255, 165, 0, 0.3)",
    stroke_width=3,
    stroke_color="#000000",
    background_color="#f0f2f6",
    height=200,
    width=350, 
    drawing_mode="freedraw",
    key="canvas",
)

# 4. Verzenden
if st.button("Formulier Verzenden", type="primary", width="stretch" if st.__version__ >= '1.30' else None):
    
    # Validatie is actief
    if not naam:
        st.error("Vul a.u.b. je volledige naam in.")
    elif geboortedatum is None:
        st.error("Vul a.u.b. je geboortedatum in.")
    elif not postcode:
        st.error("Vul a.u.b. je postcode in.")
    elif not huisnummer:
        st.error("Vul a.u.b. je huisnummer in.")
    elif not straat:
        st.error("Straatnaam ontbreekt. Controleer of de postcode en het huisnummer kloppen.")
    elif not woonplaats:
        st.error("Woonplaats ontbreekt. Controleer of de postcode en het huisnummer kloppen.")
    elif not email:
        st.error("Vul a.u.b. je e-mailadres in.")
    elif not telefoon:
        st.error("Vul a.u.b. je telefoonnummer in.")
    elif not rijbewijsnummer or not rijbewijsnummer.isdigit() or len(rijbewijsnummer) != 10:
        st.error("Het rijbewijsnummer is ongeldig. Dit moet bestaan uit exact 10 cijfers (zie item 5 op je rijbewijs).")
    elif not gekozen_merk or not gekozen_motor:
        st.error("Vul a.u.b. het merk en type van de motor in.")
    elif not akkoord:
        st.error("Je moet akkoord gaan met de voorwaarden om de proefrit te starten.")
    elif canvas_result.image_data is None or len(canvas_result.json_data["objects"]) == 0:
        st.error("Vergeet niet je handtekening te plaatsen.")
    else:
        # Alles gereed!
        form_data = {
            "Datum_Tijd": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Naam": naam,
            "Geboortedatum": geboortedatum.strftime("%d-%m-%Y"),
            "Straat": straat,
            "Huisnummer": huisnummer,
            "Postcode": postcode,
            "Woonplaats": woonplaats,
            "Email": email,
            "Telefoon": telefoon,
            "Rijbewijsnummer": rijbewijsnummer,
            "Merk": gekozen_merk,
            "Type": gekozen_motor,
            "Kenteken": "" 
        }
        
        # Poging tot opslaan
        succes, foutmelding = save_form_data(form_data, canvas_result.image_data, naam)
        
        if succes:
            st.balloons()
            st.success("Formulier succesvol verzonden! De PDF is opgeslagen en gemaild.")
        else:
            st.error(f"⚠️ {foutmelding}")
