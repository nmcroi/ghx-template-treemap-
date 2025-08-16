#!/usr/bin/env python3
"""
GHX Template Generator - Web Interface
Streamlit app met GHX design voor template generatie.
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
from datetime import datetime
import tempfile
import os

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

# GHX Custom CSS
st.markdown("""
<style>
    /* GHX Brand Colors */
    :root {
        --ghx-blue: #1e4d8b;
        --ghx-orange: #ff6a00;
        --ghx-light-blue: #5a9bd4;
        --ghx-gray: #f8f9fa;
    }
    
    /* Header styling */
    .main-header {
        background: linear-gradient(90deg, var(--ghx-blue) 0%, var(--ghx-light-blue) 100%);
        padding: 1rem 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        color: white;
    }
    
    .main-header h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: 600;
    }
    
    .main-header p {
        margin: 0.5rem 0 0 0;
        opacity: 0.9;
        font-size: 1.1rem;
    }
    
    /* GHX Logo styling */
    .ghx-logo {
        font-size: 3rem;
        font-weight: 900;
        background: linear-gradient(45deg, white, #e0e7ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-right: 1rem;
    }
    
    /* Section styling */
    .section-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        border-left: 4px solid var(--ghx-orange);
        margin-bottom: 1.5rem;
    }
    
    /* Buttons */
    .stButton > button {
        background: linear-gradient(90deg, var(--ghx-orange) 0%, #ff8533 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        background: linear-gradient(90deg, #e55a00 0%, var(--ghx-orange) 100%);
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(255, 106, 0, 0.3);
    }
    
    /* Progress indicators */
    .progress-step {
        display: inline-block;
        background: var(--ghx-light-blue);
        color: white;
        padding: 0.3rem 1rem;
        border-radius: 20px;
        margin-right: 0.5rem;
        font-size: 0.9rem;
        font-weight: 600;
    }
    
    .progress-step.active {
        background: var(--ghx-orange);
    }
    
    /* Form styling */
    .stSelectbox > div > div {
        border-radius: 8px;
        border: 2px solid #e0e7ff;
    }
    
    .stSelectbox > div > div:focus-within {
        border-color: var(--ghx-orange);
        box-shadow: 0 0 0 3px rgba(255, 106, 0, 0.1);
    }
    
    /* Success/Error messages */
    .success-message {
        background: linear-gradient(90deg, #10b981, #34d399);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .error-message {
        background: linear-gradient(90deg, #ef4444, #f87171);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background-color: var(--ghx-gray);
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display:none;}
</style>
""", unsafe_allow_html=True)

def render_header():
    """Render GHX style header."""
    st.markdown("""
    <div class="main-header">
        <div style="display: flex; align-items: center;">
            <span class="ghx-logo">GHX</span>
            <div>
                <h1>Template Generator</h1>
                <p>Genereer aangepaste Excel prijstemplates voor de Nederlandse zorgsector</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

def render_progress_steps(current_step):
    """Render progress steps."""
    steps = [
        "Context configureren",
        "Template genereren", 
        "Download resultaat"
    ]
    
    progress_html = "<div style='margin: 1rem 0;'>"
    for i, step in enumerate(steps, 1):
        active_class = "active" if i == current_step else ""
        progress_html += f'<span class="progress-step {active_class}">{i}. {step}</span>'
    progress_html += "</div>"
    
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
    
    # Header
    render_header()
    
    # Initialize session state
    if 'step' not in st.session_state:
        st.session_state.step = 1
    if 'context' not in st.session_state:
        st.session_state.context = None
    if 'generated_file' not in st.session_state:
        st.session_state.generated_file = None
    
    # Progress steps
    render_progress_steps(st.session_state.step)
    
    # Sidebar navigation
    with st.sidebar:
        st.markdown("### 🔧 Configuratie")
        
        # Reset button
        if st.button("🔄 Opnieuw beginnen"):
            st.session_state.step = 1
            st.session_state.context = None
            st.session_state.generated_file = None
            st.rerun()
        
        st.markdown("---")
        
        # Info
        st.markdown("### ℹ️ Informatie")
        st.info("""
        **Template Types:**
        - Besteleenheid
        - Verpakkingseenheid  
        - Staffel
        
        **Producttypes:**
        - Medisch
        - Lab
        - Facilitair
        - Mixed
        """)
        
        # Sample contexts
        st.markdown("### 📁 Voorbeelden")
        sample_files = {
            "GS1 Medisch": "tests/samples/sample_context_gs1.json",
            "Lab Chemicaliën": "tests/samples/sample_context_lab_chemicals.json", 
            "Facilitair": "tests/samples/sample_context_facilitair.json",
            "Staffel": "tests/samples/sample_context_staffel.json"
        }
        
        selected_sample = st.selectbox("Laad voorbeeld:", ["Selecteer..."] + list(sample_files.keys()))
        
        if selected_sample != "Selecteer..." and st.button("📥 Laad voorbeeld"):
            try:
                sample_path = Path(sample_files[selected_sample])
                with open(sample_path) as f:
                    sample_data = json.load(f)
                st.session_state.sample_data = sample_data
                st.success(f"Voorbeeld '{selected_sample}' geladen!")
                st.rerun()
            except Exception as e:
                st.error(f"Kan voorbeeld niet laden: {e}")
    
    # Main content based on step
    if st.session_state.step == 1:
        render_context_configuration()
    elif st.session_state.step == 2:
        render_template_generation()
    elif st.session_state.step == 3:
        render_download_results()

def render_context_configuration():
    """Render context configuration step."""
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("## 🎯 Stap 1: Context Configuratie")
    
    col1, col2 = st.columns(2)
    
    # Check if sample data was loaded
    sample_data = st.session_state.get('sample_data', {})
    
    with col1:
        st.markdown("### Template Instellingen")
        
        template_choice = st.selectbox(
            "Template keuze:",
            ["custom", "standard"],
            index=0 if sample_data.get('template_choice') == 'custom' else 1
        )
        
        gs1_mode = st.selectbox(
            "GS1 Modus:",
            ["none", "gs1", "gs1_only"],
            index=["none", "gs1", "gs1_only"].index(sample_data.get('gs1_mode', 'none'))
        )
        
        all_orderable = st.checkbox(
            "Besteleenheid terminologie (anders verpakking)",
            value=sample_data.get('all_orderable', True)
        )
        
        product_type = st.selectbox(
            "Product type:",
            ["medisch", "lab", "facilitair", "mixed"],
            index=["medisch", "lab", "facilitair", "mixed"].index(sample_data.get('product_type', 'medisch'))
        )
    
    with col2:
        st.markdown("### Geavanceerde Opties")
        
        has_chemicals = st.checkbox(
            "Chemicaliën/safety velden",
            value=sample_data.get('has_chemicals', False)
        )
        
        is_staffel_file = st.checkbox(
            "Staffel template",
            value=sample_data.get('is_staffel_file', False)
        )
        
        st.markdown("**Instellingen:** (optioneel)")
        institutions_text = st.text_area(
            "Instellingen (één per regel):",
            value="\n".join(sample_data.get('institutions', [])),
            height=100,
            placeholder="UMCU\nLUMC\nAMC"
        )
        
        institutions = [inst.strip() for inst in institutions_text.split('\n') if inst.strip()]
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Preview context
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("### 👀 Context Preview")
    
    context_dict = {
        "template_choice": template_choice,
        "gs1_mode": gs1_mode,
        "all_orderable": all_orderable,
        "product_type": product_type,
        "has_chemicals": has_chemicals,
        "is_staffel_file": is_staffel_file,
        "institutions": institutions,
        "version": "v1.0.0"
    }
    
    try:
        context = Context(**context_dict)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.json(context_dict)
        
        with col2:
            st.markdown("**Gegenereerde labels:**")
            labels = sorted(context.labels())
            for label in labels:
                st.markdown(f"- `{label}`")
            
            st.markdown(f"**Preset code:** `{context.get_preset_code()}`")
            st.markdown(f"**Template:** `{context.get_template_basename()}.xlsx`")
        
        # Validation
        errors = context.validate()
        if errors:
            st.markdown('<div class="error-message">', unsafe_allow_html=True)
            st.markdown("**⚠️ Validatie fouten:**")
            for error in errors:
                st.markdown(f"- {error}")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="success-message">', unsafe_allow_html=True)
            st.markdown("✅ **Context is geldig!**")
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Store context
            st.session_state.context = context
            
            # Next step button
            if st.button("➡️ Ga naar Template Generatie", type="primary"):
                st.session_state.step = 2
                st.rerun()
    
    except Exception as e:
        st.error(f"Context validatie fout: {e}")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Clear sample data after use
    if 'sample_data' in st.session_state:
        del st.session_state.sample_data

def render_template_generation():
    """Render template generation step."""
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("## ⚙️ Stap 2: Template Generatie")
    
    if not st.session_state.context:
        st.error("Geen geldige context gevonden. Ga terug naar stap 1.")
        if st.button("⬅️ Terug naar configuratie"):
            st.session_state.step = 1
            st.rerun()
        return
    
    context = st.session_state.context
    
    # Show context summary
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📋 Configuratie Samenvatting")
        st.markdown(f"**Template:** {context.get_template_basename()}")
        st.markdown(f"**Preset Code:** {context.get_preset_code()}")
        st.markdown(f"**Product Type:** {context.product_type}")
        st.markdown(f"**GS1 Modus:** {context.gs1_mode}")
        st.markdown(f"**Instellingen:** {', '.join(context.institutions) if context.institutions else 'Geen'}")
    
    with col2:
        st.markdown("### 🎨 Styling Opties")
        mandatory_color = st.color_picker("Verplichte velden kleur:", "#FFF2CC")
        hidden_color = st.color_picker("Verborgen velden kleur:", "#EEEEEE")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Generate template
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("### 🔧 Template Genereren")
    
    if st.button("🚀 Genereer Template", type="primary"):
        try:
            with st.spinner("Template genereren..."):
                # Load mapping
                mapping = load_field_mapping()
                if not mapping:
                    st.error("Kan field mapping niet laden.")
                    return
                
                # Create engine
                engine = TemplateEngine(context, mapping)
                decisions = engine.process_all_fields()
                
                # Show decisions summary
                visible_count = sum(1 for d in decisions if d.visible)
                mandatory_count = sum(1 for d in decisions if d.visible and d.mandatory)
                
                st.success(f"✅ Beslissingen berekend: {len(decisions)} velden, {visible_count} zichtbaar, {mandatory_count} verplicht")
                
                # Find template file
                templates_dir = Path("templates")
                template_name = f"{context.get_template_basename()}.xlsx"
                template_path = templates_dir / template_name
                
                if not template_path.exists():
                    st.error(f"Template bestand niet gevonden: {template_path}")
                    return
                
                # Generate template
                excel_processor = ExcelProcessor(mandatory_color, hidden_color)
                
                # Use temporary file
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
                
                # Store generated file
                st.session_state.generated_file = temp_path
                
                st.markdown('<div class="success-message">', unsafe_allow_html=True)
                st.markdown("🎉 **Template succesvol gegenereerd!**")
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Show field decisions
                with st.expander("📊 Veld Beslissingen Details"):
                    decisions_df = pd.DataFrame([
                        {
                            "Veld": d.field_name,
                            "Kolom": d.column,
                            "Zichtbaar": "✅" if d.visible else "❌",
                            "Verplicht": "⚠️" if d.mandatory else "",
                            "Dependencies": "✅" if d.dependency_satisfied else "❌" if d.depends_on else "",
                            "Notes": d.notes[:50] + "..." if len(d.notes) > 50 else d.notes
                        }
                        for d in decisions
                    ])
                    
                    st.dataframe(
                        decisions_df,
                        use_container_width=True,
                        height=400
                    )
                
                # Next step
                if st.button("➡️ Ga naar Download", type="primary"):
                    st.session_state.step = 3
                    st.rerun()
                
        except Exception as e:
            st.markdown('<div class="error-message">', unsafe_allow_html=True)
            st.markdown(f"**❌ Fout bij genereren:** {e}")
            st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Back button
    if st.button("⬅️ Terug naar configuratie"):
        st.session_state.step = 1
        st.rerun()

def render_download_results():
    """Render download results step."""
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown("## 📥 Stap 3: Download Resultaat")
    
    if not st.session_state.generated_file or not Path(st.session_state.generated_file).exists():
        st.error("Geen gegenereerd template gevonden. Ga terug naar stap 2.")
        if st.button("⬅️ Terug naar generatie"):
            st.session_state.step = 2
            st.rerun()
        return
    
    context = st.session_state.context
    file_path = Path(st.session_state.generated_file)
    
    # File info
    file_size = file_path.stat().st_size / 1024  # KB
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Template Informatie")
        st.markdown(f"**Preset Code:** {context.get_preset_code()}")
        st.markdown(f"**Template Type:** {context.get_template_basename()}")
        st.markdown(f"**Bestandsgrootte:** {file_size:.1f} KB")
        st.markdown(f"**Gegenereerd:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    with col2:
        st.markdown("### 🎯 Context Labels")
        labels = sorted(context.labels())
        for label in labels:
            st.markdown(f"- `{label}`")
    
    # Download button
    with open(file_path, "rb") as file:
        file_data = file.read()
    
    filename = f"GHX_Template_{context.get_preset_code()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    st.download_button(
        label="📥 Download Template",
        data=file_data,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    st.markdown('<div class="success-message">', unsafe_allow_html=True)
    st.markdown("✅ **Template klaar voor download!**")
    st.markdown("Het template bevat embedded metadata voor traceability.")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Actions
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🔄 Nieuwe Template Maken"):
            # Cleanup temp file
            try:
                os.unlink(file_path)
            except:
                pass
            
            st.session_state.step = 1
            st.session_state.context = None
            st.session_state.generated_file = None
            st.rerun()
    
    with col2:
        if st.button("⬅️ Terug naar Generatie"):
            st.session_state.step = 2
            st.rerun()

if __name__ == "__main__":
    main()
