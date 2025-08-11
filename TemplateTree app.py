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
            'template_choice': None,
            'all_orderable': None,
            'product_type': None,
            'chemicals_present': None,
            'volume_pricing': None,
            'gs1_sync': None,
            'organizations': []
        }
    
    # Progress indicator - Uitgebreid naar 8 stappen
    progress_steps = ['Template Keuze', 'Bestelbaarheid', 'Producttype', 'Chemicaliën', 
                      'Staffelprijzen', 'GS1 Sync', 'Zorginstellingen', 'Overzicht']
    current_step_idx = get_current_step_index(st.session_state.step)
    
    # Zorg ervoor dat de progress waarde tussen 0.0 en 1.0 blijft
    progress_value = min(current_step_idx / (len(progress_steps) - 1), 1.0)
    st.progress(progress_value)
    if current_step_idx == 0:
        st.markdown(f"**Stap {current_step_idx + 1}/{len(progress_steps)}**")
    else:
        st.markdown(f"**Stap {current_step_idx + 1}/{len(progress_steps)}:** {progress_steps[current_step_idx]}")
    
    # Render current step
    render_step(st.session_state.step)

def get_current_step_index(step):
    step_map = {
        'welcome': 0,
        'template_choice': 0,
        'question1': 1,
        'question2': 2,
        'question3': 3,
        'question4': 4,
        'question5': 5,
        'question6': 6,
        'question7': 7,
        'summary': 8
    }
    return step_map.get(step, 0)

def render_step(step):
    if step == 'welcome':
        show_welcome()
    elif step == 'template_choice':
        show_template_choice()
    elif step == 'question1':
        show_question1()
    elif step == 'question2':
        show_question2()
    elif step == 'question3':
        show_question3()
    elif step == 'question4':
        show_question4()
    elif step == 'question5':
        show_question5()
    elif step == 'question6':
        show_question6()
    elif step == 'question7':
        show_question7()
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
            st.session_state.step = 'template_choice'
            st.rerun()
            
    # Extra witruimte toevoegen
    st.markdown("<br>" * 3, unsafe_allow_html=True)
    
    # Info blok onderaan plaatsen
    st.info("💡 **Goed om te weten**: De op maat gemaakte template bevat nog steeds alle originele velden, maar verbergt enkel de niet-relevante kolommen. U kunt deze altijd weer zichtbaar maken wanneer nodig.")

def show_template_choice():
    st.header("Template Keuze")
    
    st.markdown("""
    #### Welke template heeft u nodig?
    
    **Standaard template**
    De complete template met alle 103 kolommen. Ideaal voor leveranciers met een breed productaanbod voor verschillende zorginstellingen.
    
    **Aangepaste template**
    Een gestroomlijnde template, specifiek afgestemd op uw leveranciersprofiel, producttypes en de zorginstellingen die u bedient.
    
    Als u "Standaard template" kiest, krijgt u direct de volledige template en zijn er geen verdere vragen.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📋 Standaard template", use_container_width=True):
            st.session_state.answers['template_choice'] = 'standard'
            download_full_template()
    
    with col2:
        if st.button("⚙️ Aangepaste template", use_container_width=True):
            st.session_state.answers['template_choice'] = 'custom'
            st.session_state.step = 'question1'
            st.rerun()
    
    if st.button("← Terug", use_container_width=True):
        st.session_state.step = 'welcome'
        st.rerun()

def show_question1():
    st.header("Bestelbaarheid")
    
    st.markdown("""
    #### Zijn alle producten in uw prijslijst bestelbaar?
    
    Bij GHX bieden wij de mogelijkheid om uw productdata via ons aan te leveren aan de GDSN van GS1. Deze data kunnen zowel bestelbare producten betreffen als verpakkingslagen die niet direct bestelbaar zijn.
    
    **✅ Ja, alle producten zijn bestelbaar**
    Als al uw artikelen in de prijslijst direct bestelbaar zijn, kunnen we deze velden verbergen en passen we de terminologie aan naar "Trade Unit" om uw template overzichtelijker te maken.
    
    **❌ Nee, niet alle producten zijn bestelbaar**
    Als u ook niet-bestelbare verpakkingslagen (zoals dozen of pallets) wilt opnemen, houdt de template de originele terminologie aan en toont alle benodigde velden.
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
        st.session_state.step = 'template_choice'
        st.rerun()

def show_question2():
    st.header("Producttype")
    
    st.markdown("""
    #### Wat is het hoofdtype van uw producten?
    
    De GHX prijstemplate bevat velden die specifiek zijn voor bepaalde producttypen. Door aan te geven welk hoofdtype uw producten hebben, kunnen wij de template aanpassen door irrelevante velden te verbergen.
    
    **📦 Allemaal facilitair**
    Voor leveranciers van facilitaire artikelen zoals kantoorbenodigdheden, meubilair, schoonmaakmiddelen en algemene ziekenhuisbenodigdheden.
    
    **🏥 Allemaal medisch**
    Voor leveranciers van medische hulpmiddelen, implantaten, diagnostische apparatuur en andere medische artikelen.
    
    **🔬 Allemaal laboratorium**
    Voor leveranciers van laboratoriumartikelen, reagentia, chemicaliën en labapparatuur.
    
    **🔄 Gemixte producten**
    Voor leveranciers met verschillende typen producten. In dit geval blijven alle velden zichtbaar.
    """)

    product_types = [
        'Allemaal facilitair',
        'Allemaal medisch',
        'Allemaal laboratorium',
        'Gemixte producten'
    ]
    
    col1, col2 = st.columns(2)
    
    for i, product_type in enumerate(product_types):
        col = col1 if i < 2 else col2
        with col:
            if st.button(product_type, use_container_width=True, key=f"product_type_{i}"):
                st.session_state.answers['product_type'] = product_type
                st.session_state.step = 'question3'
                st.rerun()
    
    if st.button("← Terug", use_container_width=True):
        st.session_state.step = 'question1'
        st.rerun()

def show_question3():
    st.header("Chemicaliën en Gevaarlijke Stoffen")
    
    st.markdown("""
    #### Bevat uw prijslijst producten met chemicaliën die speciale veiligheidsinformatie vereisen?
    
    Sommige producten vereisen extra veiligheidsinformatie zoals veiligheidsbladen (SDS), UN-nummers voor transport, of temperatuurspecificaties.
    
    **⚠️ Ja, chemicaliën aanwezig**
    Uw template bevat velden voor CAS-nummers, stofnamen, veiligheidsbladen, transport- en opslagvoorschriften.
    
    **✅ Nee, geen chemicaliën**
    Velden voor chemische eigenschappen en veiligheidsinformatie worden verborgen om uw template eenvoudiger te maken.
    
    Dit geldt bijvoorbeeld voor schoonmaakmiddelen, desinfectantia, laboratoriumreagentia of andere producten met gevaarlijke stoffen.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("⚠️ Ja, chemicaliën aanwezig", use_container_width=True):
            st.session_state.answers['chemicals_present'] = True
            st.session_state.step = 'question4'
            st.rerun()
    
    with col2:
        if st.button("✅ Nee, geen chemicaliën", use_container_width=True):
            st.session_state.answers['chemicals_present'] = False
            st.session_state.step = 'question4'
            st.rerun()
    
    if st.button("← Terug", use_container_width=True):
        st.session_state.step = 'question2'
        st.rerun()

def show_question4():
    st.header("Staffelprijzen")
    
    st.markdown("""
    #### Hanteert u volumestaffels of kortingen in uw prijslijst?
    
    **📊 Ja, staffelprijzen**
    Uw template bevat velden om staffelgrenzen en bijbehorende prijzen in te vullen (bijvoorbeeld: "1-10 stuks €5,00, vanaf 11 stuks €4,50").
    
    **💰 Nee, vaste prijzen**
    Velden voor volumestaffels worden verborgen. Elk artikel heeft één vaste prijs.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Ja, staffelprijzen", use_container_width=True):
            st.session_state.answers['volume_pricing'] = True
            st.session_state.step = 'question5'
            st.rerun()
    
    with col2:
        if st.button("💰 Nee, vaste prijzen", use_container_width=True):
            st.session_state.answers['volume_pricing'] = False
            st.session_state.step = 'question5'
            st.rerun()
    
    if st.button("← Terug", use_container_width=True):
        st.session_state.step = 'question3'
        st.rerun()

def show_question5():
    st.header("GS1 Datasynchronisatie")
    
    st.markdown("""
    #### Plant u uw productdata te delen via het GS1 GDSN netwerk?
    
    Het Global Data Synchronisation Network (GDSN) is een wereldwijd netwerk voor het delen van productinformatie tussen leveranciers en afnemers.
    
    **🌐 Ja, GDSN syndicatie**
    Uw template bevat extra velden die specifiek door GS1 worden gevraagd, zoals hiërarchie-omschrijvingen en GLN-codes.
    
    **📋 Nee, alleen voor GHX**
    GS1-specifieke velden worden verborgen om uw template eenvoudiger te houden.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🌐 Ja, GDSN syndicatie", use_container_width=True):
            st.session_state.answers['gs1_sync'] = True
            st.session_state.step = 'question6'
            st.rerun()
    
    with col2:
        if st.button("📋 Nee, alleen voor GHX", use_container_width=True):
            st.session_state.answers['gs1_sync'] = False
            st.session_state.step = 'question6'
            st.rerun()
    
    if st.button("← Terug", use_container_width=True):
        st.session_state.step = 'question4'
        st.rerun()

def show_question6():
    st.header("Zorginstellingen")
    
    st.markdown("""
    #### Voor welke zorginstellingen levert u de prijslijst aan?
    
    Elke zorginstelling heeft eigen specifieke vereisten voor de verplichte velden in de prijslijst. Met deze stap kunt u dit beperken tot alleen die zorginstellingen waarvoor uw prijslijst bedoeld is.
    
    **Keuze 1: Selecteer alle zorginstellingen**
    - Voordeel: U hoeft slechts één prijslijst in te vullen die geldig is voor al uw klanten
    - Nadeel: Dit resulteert in meer verplichte velden, omdat alle eisen gecombineerd worden
    
    **Keuze 2: Selecteer specifieke zorginstellingen**
    - Voordeel: Minder verplichte velden omdat alleen specifieke eisen worden meegenomen
    - Nadeel: U moet voor elke groep zorginstellingen een aparte prijslijst invullen
    
    Beschikbare zorginstellingen:
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
            st.session_state.step = 'question5'
            st.rerun()
    
    with nav_col2:
        # Disable 'Volgende' knop als er geen selectie is
        if st.button("Volgende →", use_container_width=True, disabled=len(new_selection) == 0):
            st.session_state.step = 'question7'
            st.rerun()

def show_question7():
    st.header("Overzicht en Bevestiging")
    
    st.markdown("""
    #### Controleer uw antwoorden
    
    Hieronder ziet u een overzicht van alle keuzes die u heeft gemaakt. Controleer of alles correct is voordat u uw template genereert.
    """)
    
    # Toon overzicht van alle antwoorden
    answers = st.session_state.answers
    
    st.markdown("**📋 Template keuze:**")
    if answers.get('template_choice') == 'custom':
        st.success("✅ Aangepaste template")
    else:
        st.info("📋 Standaard template")
    
    if answers.get('template_choice') == 'custom':
        st.markdown("**🛒 Bestelbaarheid:**")
        if answers.get('all_orderable'):
            st.success("✅ Ja, alle producten zijn bestelbaar")
        else:
            st.info("❌ Nee, niet alle producten zijn bestelbaar")
        
        st.markdown("**📦 Producttype:**")
        st.info(f"📦 {answers.get('product_type', 'Niet geselecteerd')}")
        
        st.markdown("**🧪 Chemicaliën:**")
        if answers.get('chemicals_present'):
            st.warning("⚠️ Ja, chemicaliën aanwezig")
        else:
            st.success("✅ Nee, geen chemicaliën")
        
        st.markdown("**💰 Staffelprijzen:**")
        if answers.get('volume_pricing'):
            st.info("📊 Ja, staffelprijzen")
        else:
            st.success("💰 Nee, vaste prijzen")
        
        st.markdown("**🌐 GS1 Sync:**")
        if answers.get('gs1_sync'):
            st.info("🌐 Ja, GDSN syndicatie")
        else:
            st.success("📋 Nee, alleen voor GHX")
        
        st.markdown("**🏥 Zorginstellingen:**")
        orgs = answers.get('organizations', [])
        if orgs:
            st.success(f"✅ {len(orgs)} zorginstelling(en) geselecteerd")
            for org in orgs[:5]:  # Toon eerste 5
                st.write(f"   • {org}")
            if len(orgs) > 5:
                st.write(f"   • ... en {len(orgs) - 5} meer")
        else:
            st.error("❌ Geen zorginstellingen geselecteerd")
    
    # Actieknoppen
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("← Terug", use_container_width=True):
            st.session_state.step = 'question6'
            st.rerun()
    
    with col2:
        if st.button("🔄 Opnieuw beginnen", use_container_width=True):
            st.session_state.step = 'welcome'
            st.session_state.answers = {
                'template_choice': None,
                'all_orderable': None,
                'product_type': None,
                'chemicals_present': None,
                'volume_pricing': None,
                'gs1_sync': None,
                'organizations': []
            }
            st.rerun()
    
    with col3:
        if st.button("🚀 Template genereren", use_container_width=True, type="primary"):
            st.session_state.step = 'summary'
            st.rerun()

def show_summary():
    st.header("Template Genereren en Downloaden")
    st.markdown("Controleer uw selecties en download uw op maat gemaakte template:")
    
    # Samenvatting
    answers = st.session_state.answers
    
    with st.container():
        st.subheader("Uw selectie:")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"**Template keuze:** {'📋 Standaard' if answers.get('template_choice') == 'standard' else '⚙️ Aangepast'}")
            st.markdown(f"**Bestelbaarheid:** {'✅ Alle producten bestelbaar' if answers.get('all_orderable') else '❌ Niet alle producten bestelbaar'}")
            st.markdown(f"**Product type:** {answers.get('product_type', 'Niet geselecteerd')}")
            st.markdown(f"**Chemicaliën:** {'⚠️ Aanwezig' if answers.get('chemicals_present') else '✅ Niet aanwezig'}")
        
        with col2:
            st.markdown(f"**Staffelprijzen:** {'📊 Ja' if answers.get('volume_pricing') else '💰 Nee'}")
            st.markdown(f"**GS1 Sync:** {'🌐 Ja' if answers.get('gs1_sync') else '📋 Nee'}")
            st.markdown(f"**Aantal zorginstellingen:** {len(answers.get('organizations', []))}")
        
        if answers.get('organizations'):
            st.markdown("**Geselecteerde zorginstellingen:**")
            for org in answers.get('organizations', [])[:10]:  # Toon eerste 10
                st.markdown(f"- {org}")
            if len(answers.get('organizations', [])) > 10:
                st.markdown(f"- ... en {len(answers.get('organizations', [])) - 10} meer")
    
    # Template info
    if answers.get('template_choice') == 'custom':
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
            if st.button("📥 Download aangepaste template", use_container_width=True, type="primary"):
                download_custom_template(answers)
        
        with col2:
            if st.button("🔄 Opnieuw beginnen", use_container_width=True):
                st.session_state.step = 'welcome'
                st.session_state.answers = {
                    'template_choice': None,
                    'all_orderable': None,
                    'product_type': None,
                    'chemicals_present': None,
                    'volume_pricing': None,
                    'gs1_sync': None,
                    'organizations': []
                }
                st.rerun()
    else:
        st.success("✅ U heeft gekozen voor de standaard template. Alle 103 kolommen zijn zichtbaar.")
        if st.button("📥 Download standaard template", use_container_width=True, type="primary"):
            download_full_template()
    
    # Terug naar begin
    if st.button("🏠 Terug naar begin", use_container_width=True):
        st.session_state.step = 'welcome'
        st.rerun()

def calculate_hidden_fields(answers):
    """
    Bereken hoeveel velden verborgen moeten worden op basis van de antwoorden.
    """
    hidden_count = 0
    
    # Bestelbaarheid
    if answers.get('all_orderable'):
        hidden_count += 5  # Verberg bestelbaarheid-gerelateerde velden
    
    # Producttype
    product_type = answers.get('product_type')
    if product_type == 'Allemaal facilitair':
        hidden_count += 15  # Verberg medische en lab velden
    elif product_type == 'Allemaal medisch':
        hidden_count += 12  # Verberg facilitaire en lab velden
    elif product_type == 'Allemaal laboratorium':
        hidden_count += 10  # Verberg facilitaire en medische velden
    # Gemixte producten: geen velden verbergen
    
    # Chemicaliën
    if not answers.get('chemicals_present'):
        hidden_count += 8  # Verberg chemische veiligheidsvelden
    
    # Staffelprijzen
    if not answers.get('volume_pricing'):
        hidden_count += 6  # Verberg staffelprijs velden
    
    # GS1 Sync
    if not answers.get('gs1_sync'):
        hidden_count += 12  # Verberg GS1-specifieke velden
    
    # Zorginstellingen (vereenvoudigde logica)
    orgs = answers.get('organizations', [])
    if orgs and len(orgs) < 10:  # Als minder dan 10 instellingen geselecteerd
        hidden_count += 5  # Verberg enkele algemene velden
    
    return min(hidden_count, 50)  # Maximaal 50 velden verbergen

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