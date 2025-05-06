import streamlit as st
import pandas as pd
from io import BytesIO
import json
from pathlib import Path
import openpyxl # Importeer openpyxl expliciet, hoewel het door pandas gebruikt wordt

# --- Configuratie ---
BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
CONFIG_DIR = BASE_DIR / "config"
STATIC_DIR = BASE_DIR / "static" # Definieer static map pad

# Laad field mapping
with open(CONFIG_DIR / "field_mapping.json", "r") as f:
    FIELD_MAPPING = json.load(f)

def main():
    st.set_page_config(
        page_title="GHX Template Generator",
        page_icon="📊", # Je kunt hier ook het pad naar een klein logo/favicon opgeven
        layout="centered",
        initial_sidebar_state="collapsed", # Optioneel: sidebar standaard inklappen
        menu_items={ # Optioneel: menu items aanpassen
            'Get Help': None,
            'Report a bug': None,
            'About': """# GHX Template Generator
            Dit is een app om GHX prijslijst templates te genereren."""
        }
     )
 
    # --- Logo toevoegen ---
    logo_path = STATIC_DIR / "ghx_logo.png" # Pad naar je logo
    if logo_path.exists():
        st.image(str(logo_path), width=200) # Toon logo, pas breedte aan naar wens
    else:
        st.warning(f"Logo niet gevonden op {logo_path}")
    # --- Einde Logo toevoegen ---

    st.title("GHX Template Generator")

    # Initialize session state
    if 'step' not in st.session_state:
        st.session_state.step = 'welcome'
    if 'answers' not in st.session_state:
        st.session_state.answers = {
            'all_orderable': None,
            'organizations': [],
            'product_type': None
        }
    
    # Progress indicator
    progress_steps = ['Welcome', 'Bestelbaarheid', 'Zorginstellingen', 
                      'Producttype', 'Overzicht']
    current_step_idx = get_current_step_index(st.session_state.step)
    
    st.progress(current_step_idx / (len(progress_steps) - 1))
    if current_step_idx == 0:
        st.markdown(f"**Stap {current_step_idx + 1}/{len(progress_steps)}**")
    else:
        st.markdown(f"**Stap {current_step_idx + 1}/{len(progress_steps)}:** {progress_steps[current_step_idx]}")
    
    # Render current step
    render_step(st.session_state.step)

def get_current_step_index(step):
    step_map = {
        'welcome': 0,
        'question1': 1,
        'question2': 2,
        'question3': 3,
        'summary': 4
    }
    return step_map.get(step, 0)

def render_step(step):
    if step == 'welcome':
        show_welcome()
    elif step == 'question1':
        show_question1()
    elif step == 'question2':
        show_question2()
    elif step == 'question3':
        show_question3()
    elif step == 'summary':
        show_summary()

def show_welcome():
    st.markdown("""
    #### Op maat gemaakte prijslijsten voor GHX-leveranciers
    
    Welkom bij de GHX Template Generator. Hier kunt u de GHX Prijstemplate downloaden.
    
    Onze standaard prijslijst-template bevat 103 kolommen om aan alle eisen van verschillende zorginstellingen en GS1 te voldoen. Dit kan nogal overweldigend zijn. Wij begrijpen dat niet elke leverancier al deze velden nodig heeft. Vandaar dat wij middels deze generator u de optie bieden een voor u op maat gemaakte template te genereren.
    
    Met deze tool kunt u:
    - Een template krijgen die voor u op maat gemaakt is met enkel de velden die voor u en uw klant relevant zijn
    - Tijd besparen door onnodige kolommen te verbergen
    
    Kies hieronder de optie die het beste bij uw behoeften past:
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Standaard template")
        st.markdown("De complete template met alle 103 kolommen. Ideaal voor leveranciers met een breed productaanbod voor verschillende zorginstellingen.")
        if st.button("📥 Volledige template downloaden", use_container_width=True):
            download_full_template()
    
    with col2:
        st.markdown("### Aangepaste template")
        st.markdown("Een gestroomlijnde template, specifiek afgestemd op uw leveranciersprofiel, producttypes en de zorginstellingen die u bedient.")
        if st.button("⚙️ Op maat gemaakte template", use_container_width=True):
            st.session_state.step = 'question1'
            st.rerun()
            
    # Extra witruimte toevoegen
    st.markdown("<br>" * 3, unsafe_allow_html=True)
    
    # Info blok onderaan plaatsen
    st.info("💡 **Goed om te weten**: De op maat gemaakte template bevat nog steeds alle originele velden, maar verbergt enkel de niet-relevante kolommen. U kunt deze altijd weer zichtbaar maken wanneer nodig.")

def show_question1():
    st.header("Bestelbaarheid")
    
    st.markdown("""
    #### Zijn alle producten in uw prijslijst bestelbaar?
    
    Bij GHX bieden wij de mogelijkheid om uw productdata via ons aan te leveren aan de GDSN van GS1. Deze data kunnen zowel bestelbare producten betreffen als verpakkingslagen die niet direct bestelbaar zijn.
    
    Als u ook niet-bestelbare verpakkingslagen (zoals dozen of pallets) wilt opnemen in uw prijslijst, dan is het belangrijk dat u per artikel kunt aangeven of deze bestelbaar is. In dat geval passen we uw template aan met extra velden om dit te specificeren.
    
    Als al uw artikelen in de prijslijst direct bestelbaar zijn, kunnen we deze velden verbergen om uw template overzichtelijker te maken.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("✅ Ja, alle producten zijn bestelbaar", use_container_width=True):
            st.session_state.answers['all_orderable'] = True
            st.session_state.step = 'question2'
            st.rerun()
    
    with col2:
        if st.button("❌ Nee, niet alle producten zijn bestelbaar", use_container_width=True):
            st.session_state.answers['all_orderable'] = False
            st.session_state.step = 'question2'
            st.rerun()
    
    if st.button("← Terug", use_container_width=True):
        st.session_state.step = 'welcome'
        st.rerun()

def show_question2():
    st.header("Zorginstellingen")
    
    st.markdown("""
    #### Voor welke zorginstellingen levert u de prijslijst aan?
    
    De GHX prijstemplate kan gebruikt worden voor alle verschillende zorginstellingen. Elke zorginstelling heeft eigen specifieke vereisten voor de verplichte velden in de prijslijst. De volledige GHX Prijstemplate houdt rekening met al die vereisten. Met deze stap kunt u dit beperken tot enkel die zorginstellingen waarvoor uw prijslijst bedoeld is.
    
    **Optie 1: Selecteer alle zorginstellingen waarmee u werkt**
    - Voordeel: U hoeft slechts één prijslijst in te vullen die geldig is voor al uw klanten
    - Nadeel: Dit zal resulteren in meer verplichte velden, omdat alle eisen van alle zorginstellingen worden gecombineerd
    
    **Optie 2: Selecteer slechts één of enkele zorginstellingen**
    - Voordeel: Minder verplichte velden omdat alleen de eisen van die specifieke instellingen worden meegenomen
    - Nadeel: U moet voor elke (groep) zorginstelling(en) een aparte prijslijst invullen
    
    Selecteer hieronder de zorginstellingen waarvoor deze specifieke prijslijst bedoeld is:
    """)
    
    # Lijst van alle zorginstellingen (alfabetisch gesorteerd)
    organizations = [
        "Academisch Ziekenhuis Maastricht",
        "Academisch Medisch Centrum",
        "AMCU",
        "Bergman BZ Rijswijk B.V",
        "CareCtrl",
        "GHX",
        "Hogeschool InHolland",
        "Hogeschool van Amsterdam",
        "Hospital Logistics",
        "Jeroen Bosch Ziekenhuis",
        "LUMC",
        "Market4Care Nederland",
        "Maxima Medisch Centrum",
        "NKI-AVL",
        "Noordwest Ziekenhuisgroep",
        "Parnassia Groep",
        "Prinses Máxima Centrum voor kinderoncologie",
        "Prothya Biosolutions Netherlands B.V.",
        "RIVM",
        "Sanquin Bloedvoorziening",
        "Stena Line",
        "Technische Universteit Delft",
        "UMC Groningen",
        "UMC Utrecht",
        "Universiteit Leiden",
        "Universiteit Twente",
        "Universiteit Utrecht",
        "Universiteit van Amsterdam",
        "Vincent van Gogh",
        "Vincent van Gogh (Vigo Groep)",
        "VU Medisch Centrum",
        "Zorgservice XL",
        "GDSN van GS1"
    ]
    
    # Een lege regel voor betere spacing
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Snelle selectie met een duidelijker kop
    st.markdown("#### Snelle selectie:")
    
    # Twee knoppen naast elkaar
    quick_col1, quick_col2 = st.columns(2)
    with quick_col1:
        select_all = st.button("✅ Selecteer alle zorginstellingen", use_container_width=True)
    with quick_col2:
        unselect_all = st.button("❌ Deselecteer alle zorginstellingen", use_container_width=True)
    
    # Haal de huidige selectie op uit session state
    current_selection = st.session_state.answers.get('organizations', [])
    
    # Verwerk selecteer/deselecteer alle
    if select_all:
        current_selection = organizations.copy()
        st.session_state.answers['organizations'] = current_selection
    elif unselect_all:
        current_selection = []
        st.session_state.answers['organizations'] = current_selection
    
    # Een lege regel voor betere spacing
    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown("#### Selecteer zorginstellingen:")
    
    # Bereken het middelpunt om de lijst in twee te splitsen
    mid_point = len(organizations) // 2
    
    # Maak twee kolommen
    col1, col2 = st.columns(2)
    
    new_selection = []
    
    # Eerste helft van de lijst in kolom 1
    with col1:
        for org in organizations[:mid_point]:
            if st.checkbox(org, value=(org in current_selection), key=f"org_{org}"):
                new_selection.append(org)
    
    # Tweede helft van de lijst in kolom 2
    with col2:
        for org in organizations[mid_point:]:
            if st.checkbox(org, value=(org in current_selection), key=f"org_{org}"):
                new_selection.append(org)
    
    # Update session state
    st.session_state.answers['organizations'] = new_selection
    
    # Toon het aantal geselecteerde zorginstellingen
    if new_selection:
        st.markdown(f"**U heeft {len(new_selection)} zorginstelling(en) geselecteerd.**")
    
    # Navigatieknoppen
    nav_col1, nav_col2 = st.columns(2)
    
    with nav_col1:
        if st.button("← Terug", use_container_width=True):
            st.session_state.step = 'question1'
            st.rerun()
    
    with nav_col2:
        # Disable 'Volgende' knop als er geen selectie is
        if st.button("Volgende →", use_container_width=True, disabled=len(new_selection) == 0):
            st.session_state.step = 'question3'
            st.rerun()

def show_question3():
    st.header("Productcategorisatie")
    
    st.markdown("""
    #### Wat is het hoofdtype van uw producten?
    
    De GHX prijstemplate bevat velden die specifiek zijn voor bepaalde producttypen. Bijvoorbeeld, medische producten vereisen andere informatie dan facilitaire producten.
    
    Door aan te geven welk hoofdtype uw producten hebben, kunnen wij de template aanpassen door:
    - Velden die niet relevant zijn voor uw producttype te verbergen
    - De template overzichtelijker en efficiënter te maken voor uw specifieke assortiment
    
    Selecteer hieronder het hoofdtype van uw producten:
    """)

    product_types = [
        'Allemaal facilitair',
        'Allemaal medisch',
        'Allemaal laboratorium',
        'Gemixte producten'
    ]
    
    for ptype in product_types:
        if st.button(f"📦 {ptype}", use_container_width=True):
            st.session_state.answers['product_type'] = ptype
            st.session_state.step = 'summary'
            st.rerun()

    # Extra witruimte
    st.markdown("<br>" * 1, unsafe_allow_html=True)
    
    st.info("💡 Als u verschillende typen producten heeft (bijvoorbeeld zowel medisch als facilitair), kies dan voor 'Gemixte producten'.")
    
    # Extra witruimte voor de terugknop
    st.markdown("<br>" * 2, unsafe_allow_html=True)
    
    if st.button("← Terug", use_container_width=True):
        st.session_state.step = 'question2'
        st.rerun()

def show_summary():
    st.header("Overzicht en Download")
    st.markdown("Controleer uw selecties:")
    
    answers = st.session_state.answers
    
    # Samenvatting
    with st.container():
        st.subheader("Uw selectie:")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"**Bestelbaarheid:** {'✅ Alle producten bestelbaar' if answers['all_orderable'] else '❌ Niet alle producten bestelbaar'}")
            st.markdown(f"**Product type:** {answers['product_type']}")
        
        with col2:
            st.markdown(f"**Aantal zorginstellingen:** {len(answers['organizations'])}")
        
        if answers['organizations']:
            st.markdown("**Geselecteerde zorginstellingen:**")
            for org in answers['organizations']:
                st.markdown(f"- {org}")
    
    # Template info
    hidden_fields = calculate_hidden_fields(answers)
    total_fields = 103
    visible_fields = total_fields - hidden_fields
    
    st.subheader("Template specificaties:")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Totaal velden", total_fields)
    with col2:
        st.metric("Zichtbare velden", visible_fields)
    with col3:
        st.metric("Verborgen velden", hidden_fields)
    
    st.progress(visible_fields / total_fields)
    
    # Download buttons
    st.subheader("Template downloaden")
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📥 Download aangepaste template", use_container_width=True):
            download_custom_template(answers)
    
    with col2:
        if st.button("🔄 Opnieuw beginnen", use_container_width=True):
            st.session_state.step = 'welcome'
            st.session_state.answers = {
                'all_orderable': None,
                'organizations': [],
                'product_type': None
            }
            st.rerun()

def calculate_hidden_fields(answers):
    """Bereken aantal te verbergen velden"""
    hidden = 0
    
    # Bestelbaarheid
    if answers['all_orderable']:
        hidden += len(FIELD_MAPPING['orderable_related'])
    
    # Product type
    if answers['product_type'] == 'Allemaal facilitair':
        hidden += len(FIELD_MAPPING['medical_fields'])
        hidden += len(FIELD_MAPPING['lab_fields'])
    elif answers['product_type'] == 'Allemaal medisch':
        hidden += len(FIELD_MAPPING['facility_fields'])
        hidden += len(FIELD_MAPPING['lab_fields'])
    elif answers['product_type'] == 'Allemaal laboratorium':
        hidden += len(FIELD_MAPPING['facility_fields'])
        hidden += len(FIELD_MAPPING['medical_fields'])
    
    return hidden

def download_custom_template(answers):
    """Genereer en download aangepaste template"""
    try:
        # Laad basis template
        base_template = pd.ExcelFile(TEMPLATES_DIR / "template_full.xlsx")
        
        # Verwerk alle sheets
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name in base_template.sheet_names:
                df = base_template.parse(sheet_name)
                
                # Pas template aan
                if sheet_name == 'PrijsTemplateSheet':
                    df = customize_main_sheet(df, answers)
                
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Download
        st.download_button(
            label="💾 Download template",
            data=output.getvalue(),
            file_name="grx_template_custom.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.success("✅ Template succesvol gegenereerd!")
        
    except Exception as e:
        st.error(f"❌ Fout bij genereren template: {str(e)}")

def customize_main_sheet(df, answers):
    """Pas hoofdsheet aan gebaseerd op antwoorden"""
    columns_to_hide = []
    
    # Bepaal te verbergen kolommen
    if answers['all_orderable']:
        columns_to_hide.extend(FIELD_MAPPING['orderable_related'])
    
    if answers['product_type'] == 'Allemaal facilitair':
        columns_to_hide.extend(FIELD_MAPPING['medical_fields'])
        columns_to_hide.extend(FIELD_MAPPING['lab_fields'])
    elif answers['product_type'] == 'Allemaal medisch':
        columns_to_hide.extend(FIELD_MAPPING['facility_fields'])
        columns_to_hide.extend(FIELD_MAPPING['lab_fields'])
    elif answers['product_type'] == 'Allemaal laboratorium':
        columns_to_hide.extend(FIELD_MAPPING['facility_fields'])
        columns_to_hide.extend(FIELD_MAPPING['medical_fields'])
    
    # Verberg kolommen (zet breedte op 0)
    for col in columns_to_hide:
        if col in df.columns:
            # Markeer deze kolommen om later te verbergen in Excel
            df[col] = df[col].astype(str)
            if col not in ['DataTest', 'ID_Column']:  # Behoud essentiële kolommen
                df[col] = ''
    
    return df

def download_full_template():
    """Download volledige template"""
    try:
        with open(TEMPLATES_DIR / "template_full.xlsx", "rb") as f:
            st.download_button(
                label="💾 Download volledige template",
                data=f,
                file_name="grx_template_full.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except FileNotFoundError:
        st.error("❌ Template bestand niet gevonden!")

if __name__ == "__main__":
    main()