import streamlit as st
import os
import json
from datetime import datetime, timedelta
from pathlib import Path
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import RGBColor

# Data en output directories
DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")
WEEK = (datetime.today() - timedelta(days=7)).isocalendar()[1]

st.set_page_config(page_title="THOR Stedelijk Informatiebeeld", layout="wide")

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

onderdelen = ["Overlast Personen", "Overlast Jeugd", "Afval"]
stadsdelen = ["Centrum", "Noord", "Oost", "Zuid", "Zuidoost", "Weesp", "West", "Nieuw-West", "VOV"]

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

if st.button("Genereer Word rapport"):
    if not DATA_DIR.exists() or not any(DATA_DIR.glob(f"{WEEK}_*.json")):
        st.warning("Geen invoer gevonden voor deze week.")
    else:
        doc = Document()
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        new_width, new_height = section.page_height, section.page_width
        section.page_width = new_width
        section.page_height = new_height

        # Titel toevoegen
        doc.add_heading(f'Rapportage week {WEEK}', 0)

        # Gegevens laden en groeperen
        data = {}
        for file in DATA_DIR.glob(f"{WEEK}_*.json"):
            with open(file, encoding="utf-8") as f:
                entry = json.load(f)
                data.setdefault(entry["onderdeel"], {})[entry["stadsdeel"]] = entry["tekst"]

        # Onderdelen in rood bold
        for onderdeel in onderdelen:
            heading = doc.add_heading(level=1)
            run = heading.add_run(onderdeel)
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # rood

            stadsdeel_data = data.get(onderdeel, {})
            for stadsdeel in stadsdelen:
                heading2 = doc.add_heading(level=2)
                run2 = heading2.add_run(stadsdeel)
                run2.bold = True
                run2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # zwart

                tekst = stadsdeel_data.get(stadsdeel, "")
                if tekst:
                    for para in tekst.split('\n'):
                        doc.add_paragraph(para)
                else:
                    doc.add_paragraph("Geen input.")

        OUTPUT_DIR.mkdir(exist_ok=True)
        output_path = OUTPUT_DIR / f"Week_{WEEK}_Rapport.docx"
        doc.save(output_path)

        st.success(f"Word rapport gegenereerd: {output_path}")
        with open(output_path, "rb") as f:
            st.download_button("Download Word rapport", f, file_name=output_path.name)
