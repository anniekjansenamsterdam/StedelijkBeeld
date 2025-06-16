import streamlit as st
import json
import re
from datetime import datetime, timedelta
from pathlib import Path
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import RGBColor, Pt

from Login import require_login

DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")
WEEK = (datetime.today() - timedelta(days=7)).isocalendar()[1]

st.set_page_config(page_title="THOR Stedelijk Informatiebeeld", layout="wide")

# require_login()

stadsdelen = ["Centrum", "Noord", "Oost", "Zuid", "Zuidoost", "Weesp", "West", "Nieuw-West", "VOV", "Nautisch Toezicht"]
onderdelen = ["Overlast personen", "Overlast jeugd", "Afval", "Parkeeroverlast/verkeersoverlast", "Overige reguliere taken"]
nautisch = ["Incidenten", "Regulier Werk", "CityControl", "SIG-meldingen"]

st.title(f"Invoer Week {WEEK}")

stadsdeel = st.selectbox("Stadsdeel", stadsdelen)

with st.form("invoer_form"):
    teksten = {}

    if stadsdeel == "Nautisch Toezicht":
        st.write("Invoer Nautisch Toezicht")
        for onderdeel in nautisch:
            teksten[onderdeel] = st.text_area(f"{onderdeel}", height=100)
    else:
        st.write(f"Invoer {stadsdeel}")
        for onderdeel in onderdelen:
            teksten[onderdeel] = st.text_area(f"{onderdeel}", height=100)

    submitted = st.form_submit_button("Opslaan")

    if submitted:
        DATA_DIR.mkdir(exist_ok=True)
        for onderdeel, tekst in teksten.items():
            safe_onderdeel = re.sub(r"[\\/]", "_", onderdeel)  # slash vervangen voor bestandsnaam
            filename = f"{WEEK}_{safe_onderdeel}_{stadsdeel}.json".replace(" ", "_")
            with open(DATA_DIR / filename, "w", encoding="utf-8") as f:
                json.dump({
                    "week": WEEK,
                    "onderdeel": onderdeel,
                    "stadsdeel": stadsdeel,
                    "tekst": tekst
                }, f)
        st.success(f"Invoer opgeslagen voor {stadsdeel}")

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

        # doc.add_heading(f'Rapportage week {WEEK}', 0)
        
        para = doc.add_paragraph()
        run = para.add_run(f"Rapportage week {WEEK}")
        run.bold = True
        run.font.size = Pt(24)
        run.font.color.rgb = RGBColor(0, 0, 0)

        # Datum eronder in kleiner lettertype
        datum_para = doc.add_paragraph()
        datum_run = datum_para.add_run(datetime.now().strftime('%d-%m-%Y'))
        datum_run.font.size = Pt(9)
        datum_run.font.color.rgb = RGBColor(0, 0, 0)

        # Dan een echte heading voor de inhoudsopgave (wel in TOC)
        inhoud = doc.add_heading(level=1)
        inhoud_run = inhoud.add_run("Inhoudsopgave")
        inhoud_run.font.color.rgb = RGBColor(0, 0, 0)
        inhoudsopgave_lijst = onderdelen + ["Nautisch Toezicht"]
        for item in inhoudsopgave_lijst:
            para = doc.add_paragraph(item, style='List Bullet')

        doc.add_page_break()

        # Gegevens laden en groeperen
        data = {}
        for file in DATA_DIR.glob(f"{WEEK}_*.json"):
            with open(file, encoding="utf-8") as f:
                entry = json.load(f)
                data.setdefault(entry["onderdeel"], {})[entry["stadsdeel"]] = entry["tekst"]

        # Eerst alle gewone onderdelen (zonder Nautisch Toezicht)
        for onderdeel in onderdelen:
            heading = doc.add_heading(level=1)
            run = heading.add_run(onderdeel)
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # rood

            stadsdeel_data = data.get(onderdeel, {})
            for stadsdeel in stadsdelen:
                if stadsdeel != "Nautisch Toezicht":
                    heading2 = doc.add_heading(level=2)
                    run2 = heading2.add_run(stadsdeel)
                    run2.bold = True
                    run2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # zwart

                    tekst = stadsdeel_data.get(stadsdeel, "")
                    if tekst:
                        for para in tekst.split('\n'):
                            doc.add_paragraph(para)

        # Nautisch Toezicht als laatste blok
        heading = doc.add_heading(level=1)
        run = heading.add_run("Nautisch Toezicht")
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # rood

        for nautisch_onderdeel in nautisch:
            heading2 = doc.add_heading(level=2)
            run2 = heading2.add_run(nautisch_onderdeel)
            run2.bold = True
            run2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # zwart

            nautisch_data = data.get(nautisch_onderdeel, {})
            tekst = nautisch_data.get("Nautisch Toezicht", "")
            if tekst:
                for para in tekst.split('\n'):
                    doc.add_paragraph(para)

        OUTPUT_DIR.mkdir(exist_ok=True)
        output_path = OUTPUT_DIR / f"Week_{WEEK}_Rapport.docx"
        doc.save(output_path)

        st.success(f"Word rapport gegenereerd: {output_path}")
        with open(output_path, "rb") as f:
            st.download_button("Download Word rapport", f, file_name=output_path.name)
