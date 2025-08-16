#!/usr/bin/env python3
"""
Tests voor GHX Template Generator engine.
"""

import unittest
import sys
from pathlib import Path

# Voeg src toe aan path voor imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from src.context import Context
from src.mapping import FieldMapping
from src.engine import TemplateEngine


class TestTemplateEngine(unittest.TestCase):
    """Test cases voor TemplateEngine."""
    
    def setUp(self):
        """Setup voor elke test."""
        # Simpele test mapping
        self.test_mapping_data = {
            "Artikelnummer": {
                "col": "B",
                "visible": "always",
                "mandatory": "always",
                "notes": "Altijd zichtbaar en verplicht."
            },
            "GTIN Besteleenheid": {
                "col": "T",
                "visible": "always",
                "notes": "Barcode identificatie."
            },
            "Artikelomschrijving Taal Code": {
                "col": "E",
                "visible_only": ["gs1"],
                "depends_on": [
                    {
                        "field": "Artikelomschrijving",
                        "not_empty": True
                    }
                ],
                "notes": "GS1 afhankelijk veld."
            },
            "CE Certificaat nummer": {
                "col": "AW",
                "visible_only": ["medisch"],
                "notes": "Medisch CE-certificaat nummer."
            },
            "CAS nummer": {
                "col": "BC",
                "visible_only": ["chemicals"],
                "notes": "Chemie-informatie."
            },
            "Staffel Vanaf": {
                "col": "AA",
                "visible_only": ["staffel"],
                "mandatory_only": ["staffel"],
                "notes": "Alleen bij staffelbestanden."
            }
        }
        
        self.mapping = FieldMapping(self.test_mapping_data)
    
    def test_gs1_context(self):
        """Test GS1 context beslissingen."""
        context = Context(
            template_choice="custom",
            gs1_mode="gs1",
            all_orderable=True,
            product_type="medisch",
            has_chemicals=False,
            is_staffel_file=False,
            institutions=["UMCU"],
            version="v1.0.0"
        )
        
        engine = TemplateEngine(context, self.mapping)
        decisions = engine.process_all_fields()
        
        # Converteer naar dictionary voor makkelijke testing
        decisions_dict = {d.field_name: d for d in decisions}
        
        # Test basis velden
        self.assertTrue(decisions_dict["Artikelnummer"].visible)
        self.assertTrue(decisions_dict["Artikelnummer"].mandatory)
        
        # Test GS1 veld
        self.assertTrue(decisions_dict["Artikelomschrijving Taal Code"].visible)
        
        # Test medisch veld
        self.assertTrue(decisions_dict["CE Certificaat nummer"].visible)
        
        # Test chemie veld (niet zichtbaar)
        self.assertFalse(decisions_dict["CAS nummer"].visible)
        
        # Test staffel veld (niet zichtbaar)
        self.assertFalse(decisions_dict["Staffel Vanaf"].visible)
    
    def test_facilitair_context(self):
        """Test facilitair context beslissingen."""
        context = Context(
            template_choice="custom",
            gs1_mode="none",
            all_orderable=False,
            product_type="facilitair",
            has_chemicals=False,
            is_staffel_file=False,
            institutions=[],
            version="v1.0.0"
        )
        
        engine = TemplateEngine(context, self.mapping)
        decisions = engine.process_all_fields()
        decisions_dict = {d.field_name: d for d in decisions}
        
        # Test basis velden
        self.assertTrue(decisions_dict["Artikelnummer"].visible)
        
        # Test GS1 veld (niet zichtbaar)
        self.assertFalse(decisions_dict["Artikelomschrijving Taal Code"].visible)
        
        # Test medisch veld (niet zichtbaar)
        self.assertFalse(decisions_dict["CE Certificaat nummer"].visible)
    
    def test_chemicals_context(self):
        """Test chemicaliën context beslissingen."""
        context = Context(
            template_choice="custom",
            gs1_mode="none",
            all_orderable=True,
            product_type="lab",
            has_chemicals=True,
            is_staffel_file=False,
            institutions=[],
            version="v1.0.0"
        )
        
        engine = TemplateEngine(context, self.mapping)
        decisions = engine.process_all_fields()
        decisions_dict = {d.field_name: d for d in decisions}
        
        # Test chemie veld (zichtbaar)
        self.assertTrue(decisions_dict["CAS nummer"].visible)
    
    def test_staffel_context(self):
        """Test staffel context beslissingen."""
        context = Context(
            template_choice="custom",
            gs1_mode="none",
            all_orderable=True,
            product_type="medisch",
            has_chemicals=False,
            is_staffel_file=True,
            institutions=[],
            version="v1.0.0"
        )
        
        engine = TemplateEngine(context, self.mapping)
        decisions = engine.process_all_fields()
        decisions_dict = {d.field_name: d for d in decisions}
        
        # Test staffel veld (zichtbaar en verplicht)
        self.assertTrue(decisions_dict["Staffel Vanaf"].visible)
        self.assertTrue(decisions_dict["Staffel Vanaf"].mandatory)
    
    def test_dependency_logic(self):
        """Test dependency evaluatie."""
        context = Context(
            template_choice="custom",
            gs1_mode="gs1",
            all_orderable=True,
            product_type="medisch",
            has_chemicals=False,
            is_staffel_file=False,
            institutions=[],
            version="v1.0.0"
        )
        
        engine = TemplateEngine(context, self.mapping)
        
        # Test met lege row data (dependency niet voldaan)
        decisions = engine.process_all_fields({})
        decisions_dict = {d.field_name: d for d in decisions}
        
        taal_code_decision = decisions_dict["Artikelomschrijving Taal Code"]
        self.assertTrue(taal_code_decision.visible)  # Zichtbaar door GS1
        self.assertFalse(taal_code_decision.dependency_satisfied)  # Dependency niet voldaan
        
        # Test met gevulde row data (dependency voldaan)
        row_data = {"Artikelomschrijving": "Test artikel"}
        decisions = engine.process_all_fields(row_data)
        decisions_dict = {d.field_name: d for d in decisions}
        
        taal_code_decision = decisions_dict["Artikelomschrijving Taal Code"]
        self.assertTrue(taal_code_decision.dependency_satisfied)  # Dependency voldaan
    
    def test_context_labels(self):
        """Test context label generatie."""
        context = Context(
            template_choice="custom",
            gs1_mode="gs1_only",
            all_orderable=True,
            product_type="lab",
            has_chemicals=True,
            is_staffel_file=False,
            institutions=["UMCU", "LUMC"],
            version="v1.0.0"
        )
        
        labels = context.labels()
        
        expected_labels = {
            "lab", "gs1", "gs1_only", "orderable_true", 
            "chemicals", "UMCU", "LUMC"
        }
        
        self.assertEqual(labels, expected_labels)


if __name__ == "__main__":
    unittest.main()
