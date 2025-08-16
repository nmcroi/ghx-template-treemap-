#!/usr/bin/env python3
"""
GHX Template Generator - Nieuwe Web Interface
Streamlit app met GHX design die de nieuwe generator backend gebruikt.
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
from datetime import datetime
import tempfile
import os
import base64

# Voeg src toe aan path
import sys
sys.path.append('src')

from src.context import Context
from src.mapping import FieldMapping
from src.engine import TemplateEngine
from src.excel import ExcelProcessor

# Page config met GHX kleuren
st.set_page_config(
    page_title="GHX Template Generator",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

def load_css():
    """Load GHX custom CSS styling."""
    st.markdown("""
    <style>
        /* Import GHX fonts */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        /* GHX Brand Colors */
        :root {
            --ghx-blue: #1e4d8b;
            --ghx-orange: #ff6a00;
            --ghx-light-blue: #5a9bd4;
            --ghx-gray: #f8f9fa;
            --ghx-dark-gray: #6c757d;
        }
        
        /* Global styling */
        .main .block-container {
            padding-top: 1rem;
            font-family: 'Inter', sans-serif;
        }
        
        /* GHX Header Bar */
        .ghx-header {
            background: linear-gradient(90deg, var(--ghx-blue) 0%, var(--ghx-light-blue) 100%);
            padding: 0.8rem 2rem;
            margin: -1rem -1rem 2rem -1rem;
            color: white;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .ghx-logo {
            font-size: 2.5rem;
            font-weight: 900;
            letter-spacing: -2px;
        }
        
        .ghx-nav {
            display: flex;
            gap: 2rem;
            font-size: 0.9rem;
        }
        
        .ghx-nav-item {
            color: rgba(255,255,255,0.8);
            text-decoration: none;
            padding: 0.5rem 1rem;
            border-radius: 4px;
            transition: all 0.3s ease;
        }
        
        .ghx-nav-item:hover, .ghx-nav-item.active {
            background: rgba(255,255,255,0.2);
            color: white;
        }
        
        .ghx-user {
            font-size: 0.9rem;
            color: rgba(255,255,255,0.9);
        }
        
        /* Main content area */
        .main-content {
            background: white;
            border-radius: 10px;
            box-shadow: 0 2px 20px rgba(0,0,0,0.05);
            margin: 1rem 0;
        }
        
        /* Page title */
        .page-title {
            font-size: 2rem;
            font-weight: 600;
            color: var(--ghx-blue);
            margin-bottom: 0.5rem;
            border-bottom: 3px solid var(--ghx-orange);
            padding-bottom: 0.5rem;
            display: inline-block;
        }
        
        /* Section cards */
        .section-card {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
            border-left: 4px solid var(--ghx-orange);
            margin-bottom: 1.5rem;
        }
        
        /* Form sections */
        .form-section {
            background: #f8f9fa;
            padding: 1.5rem;
            border-radius: 8px;
            margin: 1rem 0;
        }
        
        .form-section h3 {
            color: var(--ghx-blue);
            margin-bottom: 1rem;
            font-weight: 600;
        }
        
        /* Buttons */
        .stButton > button {
            background: linear-gradient(90deg, var(--ghx-orange) 0%, #ff8533 100%);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 0.6rem 2rem;
            font-weight: 600;
            font-family: 'Inter', sans-serif;
            transition: all 0.3s ease;
            box-shadow: 0 2px 10px rgba(255, 106, 0, 0.2);
        }
        
        .stButton > button:hover {
            background: linear-gradient(90deg, #e55a00 0%, var(--ghx-orange) 100%);
            transform: translateY(-2px);
            box-shadow: 0 4px 20px rgba(255, 106, 0, 0.3);
        }
        
        /* Upload area */
        .upload-area {
            border: 2px dashed var(--ghx-light-blue);
            border-radius: 10px;
            padding: 2rem;
            text-align: center;
            background: #f8f9fa;
            margin: 1rem 0;
        }
        
        /* Progress steps */
        .progress-container {
            display: flex;
            justify-content: center;
            align-items: center;
            margin: 2rem 0;
        }
        
        .progress-step {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 40px;
            height: 40px;
            border-radius: 50%;
            background: var(--ghx-gray);
            color: var(--ghx-dark-gray);
            font-weight: 600;
            margin: 0 1rem;
            position: relative;
        }
        
        .progress-step.active {
            background: var(--ghx-orange);
            color: white;
        }
        
        .progress-step.completed {
            background: #10b981;
            color: white;
        }
        
        .progress-line {
            width: 60px;
            height: 2px;
            background: var(--ghx-gray);
        }
        
        .progress-line.completed {
            background: #10b981;
        }
        
        /* Alert boxes */
        .alert-success {
            background: linear-gradient(90deg, #10b981, #34d399);
            color: white;
            padding: 1rem 1.5rem;
            border-radius: 8px;
            margin: 1rem 0;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        .alert-error {
            background: linear-gradient(90deg, #ef4444, #f87171);
            color: white;
            padding: 1rem 1.5rem;
            border-radius: 8px;
            margin: 1rem 0;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        .alert-warning {
            background: linear-gradient(90deg, #f59e0b, #fbbf24);
            color: white;
            padding: 1rem 1.5rem;
            border-radius: 8px;
            margin: 1rem 0;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        /* Sidebar styling */
        .css-1d391kg {
            background-color: var(--ghx-gray);
        }
        
        /* Form controls */
        .stSelectbox > div > div {
            border-radius: 8px;
            border: 2px solid #e0e7ff;
            font-family: 'Inter', sans-serif;
        }
        
        .stSelectbox > div > div:focus-within {
            border-color: var(--ghx-orange);
            box-shadow: 0 0 0 3px rgba(255, 106, 0, 0.1);
        }
        
        .stTextInput > div > div > input {
            border-radius: 8px;
            border: 2px solid #e0e7ff;
            font-family: 'Inter', sans-serif;
        }
        
        .stTextInput > div > div > input:focus {
            border-color: var(--ghx-orange);
            box-shadow: 0 0 0 3px rgba(255, 106, 0, 0.1);
        }
        
        /* Hide Streamlit elements */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .stDeployButton {display: none;}
        header {visibility: hidden;}
        
        /* Custom metric cards */
        .metric-card {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
            text-align: center;
            border-top: 4px solid var(--ghx-orange);
        }
        
        .metric-value {
            font-size: 2.5rem;
            font-weight: 700;
            color: var(--ghx-blue);
            margin-bottom: 0.5rem;
        }
        
        .metric-label {
            color: var(--ghx-dark-gray);
            font-size: 0.9rem;
            font-weight: 500;
        }
    </style>
    """, unsafe_allow_html=True)

def render_ghx_header():
    """Render GHX-style header navigation."""
    st.markdown("""
    <div class="ghx-header">
        <div class="ghx-logo">GHX</div>
        <div class="ghx-nav">
            <a href="#" class="ghx-nav-item">Home</a>
            <a href="#" class="ghx-nav-item">Berichtencentrum</a>
            <a href="#" class="ghx-nav-item">Orders</a>
            <a href="#" class="ghx-nav-item">Leveranciers</a>
            <a href="#" class="ghx-nav-item">VTA Analyse</a>
            <a href="#" class="ghx-nav-item">Status</a>
            <a href="#" class="ghx-nav-item">Artikelen</a>
            <a href="#" class="ghx-nav-item">Rapportages</a>
            <a href="#" class="ghx-nav-item active">Prijslijst</a>
            <a href="#" class="ghx-nav-item">Applicatie onderhoud</a>
        </div>
        <div class="ghx-user">mariska</div>
    </div>
    """, unsafe_allow_html=True)

def render_progress_steps(current_step):
    """Render progress indicator."""
    steps = ["Context", "Template", "Download"]
    
    progress_html = '<div class="progress-container">'
    
    for i, step in enumerate(steps):
        # Step circle
        if i < current_step - 1:
            step_class = "completed"
        elif i == current_step - 1:
            step_class = "active"
        else:
            step_class = ""
        
        progress_html += f'<div class="progress-step {step_class}">{i+1}</div>'
        
        # Line between steps
        if i < len(steps) - 1:
            line_class = "completed" if i < current_step - 1 else ""
            progress_html += f'<div class="progress-line {line_class}"></div>'
    
    progress_html += '</div>'
    
    st.markdown(progress_html, unsafe_allow_html=True)

def load_field_mapping():
    """Laad field mapping."""
    try:
        mapping_path = Path("config/field_mapping.json")
        return FieldMapping.from_file(mapping_path)
    except Exception as e:
        st.error(f"Kan field mapping niet laden: {e}")
        return None

def main():
    """Main app functie."""
    
    # Load CSS
    load_css()
    
    # Header
    render_ghx_header()
    
    # Page title
    st.markdown('<h1 class="page-title">Upload Prijslijst</h1>', unsafe_allow_html=True)
    
    # Initialize session state
    if 'step' not in st.session_state:
        st.session_state.step = 1
    if 'context' not in st.session_state:
        st.session_state.context = None
    if 'generated_file' not in st.session_state:
        st.session_state.generated_file = None
    
    # Progress indicator
    render_progress_steps(st.session_state.step)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 🔧 Template Generator")
        
        # Reset button
        if st.button("🔄 Opnieuw beginnen"):
            st.session_state.step = 1
            st.session_state.context = None
            st.session_state.generated_file = None
            st.rerun()
        
        st.markdown("---")
        
        # Quick actions
        st.markdown("### ⚡ Snelle Acties")
        
        sample_contexts = {
            "🏥 GS1 Medisch": "tests/samples/sample_context_gs1.json",
            "🧪 Lab + Chemicaliën": "tests/samples/sample_context_lab_chemicals.json",
            "🏢 Facilitair": "tests/samples/sample_context_facilitair.json",
            "📊 Staffel": "tests/samples/sample_context_staffel.json"
        }
        
        st.markdown("**Snel laden:**")
        for name, path in sample_contexts.items():
            if st.button(name, key=f"quick_{name}"):
                load_sample_context(path)
        
        st.markdown("---")
        
        # Documentation links
        st.markdown("### 📚 Documentatie")
        st.markdown("- [Template Handleiding](#)")
        st.markdown("- [Veld Mapping](#)")
        st.markdown("- [GHX Support](#)")
    
    # Main content based on step
    if st.session_state.step == 1:
        render_context_step()
    elif st.session_state.step == 2:
        render_generation_step()
    elif st.session_state.step == 3:
        render_download_step()

def load_sample_context(file_path):
    """Load sample context from file."""
    try:
        with open(file_path) as f:
            sample_data = json.load(f)
        st.session_state.sample_data = sample_data
        st.success(f"Voorbeeld context geladen!")
        st.rerun()
    except Exception as e:
        st.error(f"Kan voorbeeld niet laden: {e}")

def render_context_step():
    """Render context configuration step."""
    
    st.markdown('<div class="main-content">', unsafe_allow_html=True)
    
    # Section card
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    
    # Main question section (like GHX interface)
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### Is de prijslijst voor 1 of meerdere bestellers?")
        
        template_choice = st.radio(
            "Keuze bestellers",
            ["1 besteller", "Meerdere bestellers"],
            horizontal=True,
            label_visibility="hidden"
        )
        
        # Map to our format
        template_choice_mapped = "custom" if template_choice == "1 besteller" else "standard"
        
        st.markdown("### Besteller")
        
        # Sample data
        sample_data = st.session_state.get('sample_data', {})
        
        # Institution selection (like dropdown in GHX)
        institutions_options = ["UMCU", "LUMC", "AMC", "VUmc", "Erasmus MC", "MUMC", "UMC Groningen", "Radboudumc", "Andere..."]
        selected_institutions = st.multiselect(
            "Selecteer instellingen:",
            institutions_options,
            default=sample_data.get('institutions', [])
        )
        
        # Date field (like in GHX)
        st.markdown("### Startdatum prijslijst")
        start_date = st.date_input("Startdatum:", datetime.now().date())
        
        # Staff file option
        st.markdown("### Staffelbestand")
        staffel_options = ["Ja", "Nee"]
        is_staffel = st.radio("Staffel optie", staffel_options, index=1 if not sample_data.get('is_staffel_file', False) else 0, horizontal=True, label_visibility="hidden")
        is_staffel_mapped = is_staffel == "Ja"
    
    with col2:
        # Right panel like GHX
        st.markdown('<div class="form-section">', unsafe_allow_html=True)
        st.markdown("### Voorwaarden")
        st.info("""
        - De prijslijst moet een Excel-bestand zijn in het formaat .xlsx
        - De sheet waar de prijzen en productinformatie in staat moet de eerste sheet zijn
        - Er mogen geen verwijzingen en/of (Excel) formules in de velden van de productinformatie staan
        """)
        
        # Download template link (like GHX)
        st.markdown("📥 **Download template formaat**")
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Advanced configuration
    with st.expander("🔧 Geavanceerde Configuratie"):
        col1, col2 = st.columns(2)
        
        with col1:
            gs1_mode = st.selectbox(
                "GS1 Modus:",
                ["none", "gs1", "gs1_only"],
                index=["none", "gs1", "gs1_only"].index(sample_data.get('gs1_mode', 'none'))
            )
            
            product_type = st.selectbox(
                "Product type:",
                ["medisch", "lab", "facilitair", "mixed"],
                index=["medisch", "lab", "facilitair", "mixed"].index(sample_data.get('product_type', 'medisch'))
            )
        
        with col2:
            all_orderable = st.checkbox(
                "Besteleenheid terminologie",
                value=sample_data.get('all_orderable', True)
            )
            
            has_chemicals = st.checkbox(
                "Chemicaliën/safety velden",
                value=sample_data.get('has_chemicals', False)
            )
    
    # Context validation and preview
    context_dict = {
        "template_choice": template_choice_mapped,
        "gs1_mode": gs1_mode,
        "all_orderable": all_orderable,
        "product_type": product_type,
        "has_chemicals": has_chemicals,
        "is_staffel_file": is_staffel_mapped,
        "institutions": selected_institutions,
        "version": "v1.0.0"
    }
    
    try:
        context = Context(**context_dict)
        errors = context.validate()
        
        if errors:
            st.markdown('<div class="alert-error">⚠️ Configuratie fouten: ' + '; '.join(errors) + '</div>', unsafe_allow_html=True)
        else:
            st.session_state.context = context
            
            # Show preview
            with st.expander("👀 Context Preview"):
                col1, col2 = st.columns(2)
                with col1:
                    st.json(context_dict)
                with col2:
                    st.write("**Labels:**", ", ".join(sorted(context.labels())))
                    st.write("**Preset:**", context.get_preset_code())
            
            # Next button (GHX style)
            st.markdown('<div style="text-align: center; margin: 2rem 0;">', unsafe_allow_html=True)
            if st.button("➡️ Verstuur", type="primary", key="next_step"):
                st.session_state.step = 2
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
    
    except Exception as e:
        st.markdown(f'<div class="alert-error">❌ Context fout: {e}</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Clear sample data
    if 'sample_data' in st.session_state:
        del st.session_state.sample_data

def render_generation_step():
    """Render template generation step."""
    st.markdown('<div class="main-content">', unsafe_allow_html=True)
    
    if not st.session_state.context:
        st.markdown('<div class="alert-error">❌ Geen geldige context. Ga terug naar stap 1.</div>', unsafe_allow_html=True)
        if st.button("⬅️ Terug"):
            st.session_state.step = 1
            st.rerun()
        return
    
    context = st.session_state.context
    
    # Context summary (like GHX confirmation)
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("### 📋 Configuratie Bevestiging")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{context.product_type.upper()}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">Product Type</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{context.gs1_mode.upper()}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">GS1 Modus</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{len(context.institutions)}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">Instellingen</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Generation section
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("### 🚀 Template Genereren")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if st.button("📄 Genereer Aangepast Template", type="primary", key="generate"):
            generate_template(context)
    
    with col2:
        if st.button("⬅️ Terug naar Configuratie"):
            st.session_state.step = 1
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

def generate_template(context):
    """Generate template with progress indicator."""
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Load mapping
        status_text.text("🔄 Field mapping laden...")
        progress_bar.progress(20)
        
        mapping = load_field_mapping()
        if not mapping:
            st.error("Kan field mapping niet laden.")
            return
        
        # Step 2: Create engine
        status_text.text("⚙️ Template engine initialiseren...")
        progress_bar.progress(40)
        
        engine = TemplateEngine(context, mapping)
        decisions = engine.process_all_fields()
        
        # Step 3: Statistics
        status_text.text("📊 Veld beslissingen berekenen...")
        progress_bar.progress(60)
        
        visible_count = sum(1 for d in decisions if d.visible)
        mandatory_count = sum(1 for d in decisions if d.visible and d.mandatory)
        
        # Step 4: Find template
        status_text.text("📁 Template bestand zoeken...")
        progress_bar.progress(70)
        
        templates_dir = Path("templates")
        template_name = f"{context.get_template_basename()}.xlsx"
        template_path = templates_dir / template_name
        
        if not template_path.exists():
            st.error(f"Template bestand niet gevonden: {template_path}")
            return
        
        # Step 5: Generate
        status_text.text("✨ Excel template genereren...")
        progress_bar.progress(90)
        
        excel_processor = ExcelProcessor()
        
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_path = Path(temp_file.name)
        temp_file.close()
        
        context_dict = context.to_dict()
        context_dict["_labels"] = list(context.labels())
        
        excel_processor.process_template(
            template_path,
            temp_path,
            decisions,
            context_dict,
            "Template NL"
        )
        
        # Complete
        progress_bar.progress(100)
        status_text.text("✅ Template succesvol gegenereerd!")
        
        st.session_state.generated_file = temp_path
        
        # Success message
        st.markdown(f'<div class="alert-success">🎉 Template gegenereerd! {visible_count} zichtbare velden, {mandatory_count} verplicht</div>', unsafe_allow_html=True)
        
        # Show statistics
        with st.expander("📊 Generatie Statistieken"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Totaal Velden", len(decisions))
            with col2:
                st.metric("Zichtbaar", visible_count)
            with col3:
                st.metric("Verplicht", mandatory_count)
        
        # Next step
        if st.button("➡️ Ga naar Download", type="primary"):
            st.session_state.step = 3
            st.rerun()
        
    except Exception as e:
        progress_bar.progress(0)
        status_text.text("")
        st.markdown(f'<div class="alert-error">❌ Fout bij genereren: {e}</div>', unsafe_allow_html=True)

def render_download_step():
    """Render download step."""
    st.markdown('<div class="main-content">', unsafe_allow_html=True)
    
    if not st.session_state.generated_file or not Path(st.session_state.generated_file).exists():
        st.markdown('<div class="alert-error">❌ Geen template gevonden. Ga terug naar generatie.</div>', unsafe_allow_html=True)
        if st.button("⬅️ Terug"):
            st.session_state.step = 2
            st.rerun()
        return
    
    context = st.session_state.context
    file_path = Path(st.session_state.generated_file)
    
    # Success section
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("### 🎉 Template Klaar!")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("Je aangepaste GHX template is klaar voor download:")
        
        file_size = file_path.stat().st_size / 1024
        filename = f"GHX_Template_{context.get_preset_code()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        st.markdown(f"**📄 Bestand:** {filename}")
        st.markdown(f"**📊 Grootte:** {file_size:.1f} KB")
        st.markdown(f"**⏰ Gegenereerd:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        st.markdown(f"**🏷️ Preset:** {context.get_preset_code()}")
        
        # Download button
        with open(file_path, "rb") as file:
            file_data = file.read()
        
        st.download_button(
            label="📥 Download Template",
            data=file_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    
    with col2:
        st.markdown("**🏷️ Context Labels:**")
        labels = sorted(context.labels())
        for label in labels:
            st.markdown(f"- `{label}`")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Actions
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🔄 Nieuw Template Maken"):
            cleanup_temp_file()
            st.session_state.step = 1
            st.session_state.context = None
            st.session_state.generated_file = None
            st.rerun()
    
    with col2:
        if st.button("⬅️ Terug naar Generatie"):
            st.session_state.step = 2
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

def cleanup_temp_file():
    """Cleanup temporary files."""
    if st.session_state.generated_file:
        try:
            os.unlink(st.session_state.generated_file)
        except:
            pass

if __name__ == "__main__":
    main()
