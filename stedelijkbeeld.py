import streamlit as st
import json
import re
from datetime import datetime, timedelta
from pathlib import Path
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from Login import require_login

require_login()

DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")
WEEK = (datetime.today() - timedelta(days=7)).isocalendar()[1]

st.set_page_config(page_title="THOR Stedelijk Informatiebeeld", layout="wide")

stadsdelen = ["Algemeen beeld", "Centrum", "Noord", "Oost", "Zuid", "Zuidoost", "Weesp", "West", "Nieuw-West", "VOV", "Nautisch Toezicht"]
onderdelen = ["Overlast personen", "Overlast jeugd", "Afval", "Parkeeroverlast/verkeersoverlast", "Overige reguliere taken"]
nautisch = ["Incidenten", "Regulier Werk", "CityControl", "SIG-meldingen"]

st.title(f"Invoer Stedelijk Beeld Week {WEEK}")

def reset_text_fields():
    """Reset de tekstvelden in de session state."""
    for onderdeel in onderdelen + nautisch:
        st.session_state[onderdeel] = ""

def load_data_for_stadsdeel(stadsdeel):
    """Laad eerder opgeslagen data in st.session_state voor het gekozen stadsdeel."""
    # Eerst leegmaken
    reset_text_fields()
    
    # Zoek bestanden voor de week en stadsdeel
    for onderdeel in onderdelen + nautisch:
        safe_onderdeel = re.sub(r"[\\/]", "_", onderdeel)
        filename = f"{WEEK}_{safe_onderdeel}_{stadsdeel}.json".replace(" ", "_")
        filepath = DATA_DIR / filename
        if filepath.exists():
            with open(filepath, encoding="utf-8") as f:
                try:
                    data = json.load(f)
                    # Zet de tekst in session_state zodat het in de text_area verschijnt
                    st.session_state[onderdeel] = data.get("tekst", "")
                except json.JSONDecodeError:
                    # Bestand corrupte of onjuist formaat, overslaan
                    st.session_state[onderdeel] = ""

# Initialiseer session_state voor stadsdeel
if "stadsdeel" not in st.session_state:
    st.session_state.stadsdeel = stadsdelen[0]

# Laad data telkens als stadsdeel verandert
def on_stadsdeel_change():
    load_data_for_stadsdeel(st.session_state.stadsdeel)

stadsdeel = st.selectbox(
    "Stadsdeel / Specialisme / Algemeen Beeld",
    stadsdelen,
    key="stadsdeel",
    on_change=on_stadsdeel_change
)

# Laad data bij eerste keer openen
if not any(st.session_state.get(onderdeel, None) for onderdeel in onderdelen + nautisch):
    load_data_for_stadsdeel(stadsdeel)

with st.form("invoer_form"):
    teksten = {}

    if stadsdeel == "Nautisch Toezicht":
        st.write("Nautisch Toezicht")
        for onderdeel in nautisch:
            teksten[onderdeel] = st.text_area(f"{onderdeel}", height=100, key=onderdeel)
    else:
        st.write(f"{stadsdeel}")
        for onderdeel in onderdelen:
            teksten[onderdeel] = st.text_area(f"{onderdeel}", height=100, key=onderdeel)

    submitted = st.form_submit_button("Opslaan")

    if submitted:
        DATA_DIR.mkdir(exist_ok=True)
        for onderdeel, tekst in teksten.items():
            safe_onderdeel = re.sub(r"[\\/]", "_", onderdeel)
            filename = f"{WEEK}_{safe_onderdeel}_{stadsdeel}.json".replace(" ", "_")
            with open(DATA_DIR / filename, "w", encoding="utf-8") as f:
                json.dump({
                    "week": WEEK,
                    "onderdeel": onderdeel,
                    "stadsdeel": stadsdeel,
                    "tekst": tekst
                }, f)
        st.success(f"Invoer opgeslagen voor {stadsdeel}")

# (rest van de code rapport generatie zoals jij die had)
# ...


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

        # Titel gecentreerd
        titel = doc.add_paragraph()
        titel.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = titel.add_run(f"Stedelijk Beeld Week {WEEK}")
        run.bold = True
        run.font.size = Pt(32)
        run.font.color.rgb = RGBColor(0, 0, 0)

        ondertitel = doc.add_paragraph()
        ondertitel.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ondertitel_run = ondertitel.add_run("THOR Informatiemanagement")
        ondertitel_run.font.size = Pt(20)
        ondertitel_run.font.color.rgb = RGBColor(0, 0, 0)

        # Datum eronder gecentreerd en kleiner
        datum = doc.add_paragraph()
        datum.alignment = WD_ALIGN_PARAGRAPH.CENTER
        datum_run = datum.add_run(datetime.now().strftime('%d-%m-%Y'))
        datum_run.font.size = Pt(20)
        datum_run.font.color.rgb = RGBColor(0, 0, 0)
        datum.paragraph_format.space_after = Pt(60)

        # Inhoudsopgave
        inhoud = doc.add_heading(level=1)
        inhoud_run = inhoud.add_run("Inhoudsopgave")
        inhoud.paragraph_format.space_after = Pt(12)
        inhoud_run.font.color.rgb = RGBColor(0, 0, 0)
        inhoudsopgave_lijst = onderdelen + ["Nautisch Toezicht"]
        for i, item in enumerate(inhoudsopgave_lijst, start=1):
            para = doc.add_paragraph(f"{i}. {item}")
            para.style = 'Normal'
            para.paragraph_format.left_indent = Pt(20)

        doc.add_page_break()

        # Data laden
        data = {}
        for file in DATA_DIR.glob(f"{WEEK}_*.json"):
            with open(file, encoding="utf-8") as f:
                entry = json.load(f)
                data.setdefault(entry["onderdeel"], {})[entry["stadsdeel"]] = entry["tekst"]

        # Nummer teller starten
        counter = 1

        # Hoofdonderdelen met oplopend nummer
        for onderdeel in onderdelen:
            heading = doc.add_heading(level=1)
            run = heading.add_run(f"{counter}. {onderdeel}")
            run.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  #rood
            
            counter += 1

            stadsdeel_data = data.get(onderdeel, {})
            for stadsdeel in stadsdelen:
                if stadsdeel != "Nautisch Toezicht":
                    heading2 = doc.add_heading(level=2)
                    run2 = heading2.add_run(stadsdeel)
                    run2.bold = True
                    
                    if stadsdeel == "Algemeen beeld":
                        run2.font.color.rgb = RGBColor(0x00, 0x70, 0xC0) #blauw
                    else:
                        run2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  #zwart
                    tekst = stadsdeel_data.get(stadsdeel, "")
                    if tekst:
                        for para in tekst.split('\n'):
                            doc.add_paragraph(para)
            doc.add_page_break()

        # Nautisch Toezicht als laatste met nummer
        heading = doc.add_heading(level=1)
        run = heading.add_run(f"{counter}. Nautisch Toezicht")
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # rood
        counter += 1

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
