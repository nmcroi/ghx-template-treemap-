#!/usr/bin/env python3
"""
Flask Backend API voor GHX Template Generator.
Verbindt het HTML frontend met de Python template generator modules.
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import tempfile
import os
from pathlib import Path
import json

# Import onze generator modules
import sys
sys.path.append('src')

from src.context import Context
from src.mapping import FieldMapping
from src.engine import TemplateEngine
from src.excel import ExcelProcessor

app = Flask(__name__)
CORS(app)  # Allow cross-origin requests from HTML

# Global variabelen
field_mapping = None
temp_files = {}  # Track temporary files

def load_field_mapping():
    """Laad field mapping eenmalig bij startup."""
    global field_mapping
    try:
        mapping_path = Path("config/field_mapping.json")
        field_mapping = FieldMapping.from_file(mapping_path)
        print(f"✅ Field mapping geladen: {len(field_mapping.get_all_fields())} velden")
    except Exception as e:
        print(f"❌ Kan field mapping niet laden: {e}")
        field_mapping = None

@app.route('/api/validate-context', methods=['POST'])
def validate_context():
    """Valideer context JSON en return labels."""
    try:
        context_data = request.json
        
        # Maak context object
        context = Context(**context_data)
        
        # Valideer
        errors = context.validate()
        
        # Genereer labels
        labels = list(context.labels())
        preset_code = context.get_preset_code()
        template_basename = context.get_template_basename()
        
        return jsonify({
            'valid': len(errors) == 0,
            'errors': errors,
            'labels': labels,
            'preset_code': preset_code,
            'template_basename': template_basename,
            'context': context.to_dict()
        })
        
    except Exception as e:
        return jsonify({
            'valid': False,
            'errors': [str(e)],
            'labels': [],
            'preset_code': '',
            'template_basename': '',
            'context': {}
        }), 400

@app.route('/api/generate-template', methods=['POST'])
def generate_template():
    """Genereer Excel template en return download info."""
    try:
        context_data = request.json
        
        if not field_mapping:
            return jsonify({'error': 'Field mapping niet geladen'}), 500
        
        # Maak context
        context = Context(**context_data)
        errors = context.validate()
        
        if errors:
            return jsonify({'error': f'Context validatie gefaald: {"; ".join(errors)}'}), 400
        
        # Maak engine en bereken beslissingen
        engine = TemplateEngine(context, field_mapping)
        decisions = engine.process_all_fields()
        
        # Statistieken
        visible_count = sum(1 for d in decisions if d.visible)
        mandatory_count = sum(1 for d in decisions if d.visible and d.mandatory)
        
        # Vind template bestand
        templates_dir = Path("templates")
        template_name = f"{context.get_template_basename()}.xlsx"
        template_path = templates_dir / template_name
        
        if not template_path.exists():
            return jsonify({'error': f'Template bestand niet gevonden: {template_path}'}), 404
        
        # Maak temporary output bestand
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_path = Path(temp_file.name)
        temp_file.close()
        
        # Genereer template
        excel_processor = ExcelProcessor()
        context_dict = context.to_dict()
        context_dict["_labels"] = list(context.labels())
        
        excel_processor.process_template(
            template_path,
            temp_path,
            decisions,
            context_dict,
            "Template NL"  # GHX templates gebruiken "Template NL"
        )
        
        # Bewaar bestand info voor download
        file_id = os.path.basename(temp_path)
        temp_files[file_id] = {
            'path': temp_path,
            'filename': f"GHX_Template_{context.get_preset_code()}_{context_data.get('timestamp', 'generated')}.xlsx",
            'context': context_dict
        }
        
        return jsonify({
            'success': True,
            'file_id': file_id,
            'filename': temp_files[file_id]['filename'],
            'stats': {
                'total_fields': len(decisions),
                'visible_fields': visible_count,
                'mandatory_fields': mandatory_count
            },
            'preset_code': context.get_preset_code(),
            'file_size_kb': round(temp_path.stat().st_size / 1024, 1)
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/download/<file_id>')
def download_template(file_id):
    """Download gegenereerd template bestand."""
    try:
        if file_id not in temp_files:
            return jsonify({'error': 'Bestand niet gevonden'}), 404
        
        file_info = temp_files[file_id]
        file_path = file_info['path']
        filename = file_info['filename']
        
        if not file_path.exists():
            return jsonify({'error': 'Bestand bestaat niet meer'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/info')
def api_info():
    """API informatie."""
    return jsonify({
        'name': 'GHX Template Generator API',
        'version': '1.0.0',
        'field_mapping_loaded': field_mapping is not None,
        'field_count': len(field_mapping.get_all_fields()) if field_mapping else 0,
        'temp_files': len(temp_files)
    })

@app.route('/api/cleanup', methods=['POST'])
def cleanup_temp_files():
    """Ruim temporary bestanden op."""
    try:
        cleaned = 0
        for file_id, file_info in list(temp_files.items()):
            try:
                file_path = file_info['path']
                if file_path.exists():
                    os.unlink(file_path)
                del temp_files[file_id]
                cleaned += 1
            except:
                pass
        
        return jsonify({
            'success': True,
            'cleaned_files': cleaned,
            'remaining_files': len(temp_files)
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Error handlers
@app.errorhandler(404)
def not_found(error):
    return jsonify({'error': 'Endpoint niet gevonden'}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': 'Interne server fout'}), 500

if __name__ == '__main__':
    print("🚀 GHX Template Generator API wordt gestart...")
    
    # Laad field mapping bij startup
    load_field_mapping()
    
    if not field_mapping:
        print("❌ WAARSCHUWING: Field mapping niet geladen! API werkt niet correct.")
    
    print("📡 API endpoints beschikbaar:")
    print("   POST /api/validate-context - Valideer context")
    print("   POST /api/generate-template - Genereer template")
    print("   GET  /api/download/<file_id> - Download template")
    print("   GET  /api/info - API informatie")
    print("   POST /api/cleanup - Ruim temp bestanden op")
    
    # Start Flask server
    app.run(
        host='127.0.0.1',
        port=5000,
        debug=True
    )
