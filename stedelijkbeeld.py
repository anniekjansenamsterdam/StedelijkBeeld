import streamlit as st
import os
import json
from datetime import datetime, timedelta
from pathlib import Path
from weasyprint import HTML
import pdfkit

WKHTMLTOPDF_PATH = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

WEEK = (datetime.today() - timedelta(days=7)).isocalendar()[1]
DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")

# -------------------------------
# Inloggen met users in secrets
# -------------------------------

def login():
    st.sidebar.title("🔐 Login")
    username = st.sidebar.text_input("Gebruikersnaam")
    password = st.sidebar.text_input("Wachtwoord", type="password")
    login_knop = st.sidebar.button("Inloggen")

    if login_knop:
        if username in st.secrets["users"] and st.secrets["users"][username] == password:
            st.session_state["logged_in"] = True
            st.session_state["user"] = username
        else:
            st.sidebar.error("Ongeldige gebruikersnaam of wachtwoord")

if "logged_in" not in st.session_state or not st.session_state["logged_in"]:
    login()
    st.stop()

st.set_page_config(page_title="THOR Stedelijk Informatiebeeld", layout="wide")

onderdelen = ["Overlast Personen", "Overlast Jeugd", "Afval"]
stadsdelen = ["Centrum", "Noord", "Oost", "Zuid", "Zuidoost", "Weesp", "West", "Nieuw-West", "VOV"]

st.set_page_config(page_title="Team IM Rapportage", layout="wide")
st.title(f"Invoer Week {WEEK}")

with st.form("invoer_form"):
    onderdeel = st.selectbox("Onderdeel", onderdelen)
    stadsdeel = st.selectbox("Stadsdeel", stadsdelen)
    tekst = st.text_area("Invoer tekst", height=200)
    submitted = st.form_submit_button("Opslaan")

    if submitted:
        if tekst.strip() == "":
            st.error("Tekst mag niet leeg zijn.")
        else:
            DATA_DIR.mkdir(exist_ok=True)
            filename = f"{WEEK}_{onderdeel}_{stadsdeel}.json".replace(" ", "_")
            with open(DATA_DIR / filename, "w", encoding="utf-8") as f:
                json.dump({"week": WEEK, "onderdeel": onderdeel, "stadsdeel": stadsdeel, "tekst": tekst}, f)
            st.success(f"Invoer opgeslagen voor {onderdeel} - {stadsdeel}.")

st.header("Samenvatten en rapport genereren")

samenvattingen = {}
all_entries = {}

if DATA_DIR.exists():
    for file in DATA_DIR.glob(f"{WEEK}_*.json"):
        with open(file, encoding="utf-8") as f:
            entry = json.load(f)
            key = (entry["onderdeel"], entry["stadsdeel"])
            all_entries[key] = entry

for key in sorted(all_entries.keys()):
    onderdeel, stadsdeel = key
    samenv = st.text_area(f"Samenvatting voor {onderdeel} - {stadsdeel}", key=f"sum_{onderdeel}_{stadsdeel}")
    samenvattingen[key] = samenv

if st.button("Genereer PDF"):
    html = f"<h1>Week {WEEK} Rapportage</h1>"

    for onderdeel in onderdelen:
        html += f"<h2>{onderdeel}</h2>"
        for stadsdeel in stadsdelen:
            key = (onderdeel, stadsdeel)
            if key in all_entries:
                sam = samenvattingen.get(key, "")
                invoer_tekst = all_entries[key]['tekst'].replace('\n', '<br>')
                sam_html = sam.replace('\n', '<br>') if sam else "<em>Geen samenvatting</em>"
                html += f"""
                <h3>{stadsdeel}</h3>
                <strong>Samenvatting:</strong><p>{sam_html}</p>
                <strong>Invoer:</strong><p>{invoer_tekst}</p>
                """

    OUTPUT_DIR.mkdir(exist_ok=True)
    pdf_path = OUTPUT_DIR / f"Week_{WEEK}_Rapport.pdf"

    try:
        pdfkit.from_string(html, str(pdf_path), configuration=config)
        st.success(f"PDF succesvol gegenereerd: {pdf_path}")
        with open(pdf_path, "rb") as f:
            st.download_button("Download PDF", data=f, file_name=pdf_path.name, mime="application/pdf")
    except Exception as e:
        st.error(f"Fout bij genereren PDF: {e}")