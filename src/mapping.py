"""
Field mapping loader en validator voor GHX Template Generator.

Bevat functies voor het laden en valideren van de field_mapping.json configuratie.
"""

import json
from pathlib import Path
from typing import Dict, Any, List, Set


class FieldMapping:
    """
    Field mapping configuratie loader en validator.
    """
    
    # Verwachte sleutel volgorde volgens specificatie
    EXPECTED_KEY_ORDER = [
        "col",
        "visible", "visible_only", "visible_except",
        "mandatory", "mandatory_only", "mandatory_except", 
        "depends_on",
        "depends_trigger_for",  # Optioneel
        "notes"
    ]
    
    # Geldige visibility waarden
    VALID_VISIBILITY = {"always", "never"}
    
    # Geldige mandatory waarden
    VALID_MANDATORY = {"always", "never"}
    
    def __init__(self, mapping_data: Dict[str, Any]):
        """
        Initialiseer field mapping.
        
        Args:
            mapping_data: Dictionary met veld configuraties
        """
        self.fields = mapping_data
        self._validate_structure()
    
    @classmethod
    def from_file(cls, file_path: Path) -> "FieldMapping":
        """
        Laad field mapping van JSON bestand.
        
        Args:
            file_path: Pad naar field_mapping.json
            
        Returns:
            FieldMapping object
            
        Raises:
            FileNotFoundError: Als bestand niet bestaat
            ValueError: Als JSON ongeldig of validatie faalt
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            return cls(data)
            
        except FileNotFoundError:
            raise FileNotFoundError(f"Field mapping bestand niet gevonden: {file_path}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Ongeldig JSON in {file_path}: {e}")
    
    def _validate_structure(self) -> None:
        """
        Valideer de mapping structuur.
        
        Raises:
            ValueError: Als structuur ongeldig is
        """
        errors = []
        
        for field_name, field_config in self.fields.items():
            field_errors = self._validate_field(field_name, field_config)
            errors.extend(field_errors)
        
        if errors:
            raise ValueError(f"Field mapping validatie gefaald:\n" + "\n".join(errors))
    
    def _validate_field(self, field_name: str, config: Dict[str, Any]) -> List[str]:
        """
        Valideer een enkel veld configuratie.
        
        Args:
            field_name: Naam van het veld
            config: Veld configuratie dictionary
            
        Returns:
            Lijst van foutmeldingen
        """
        errors = []
        
        # Verplichte 'col' sleutel
        if "col" not in config:
            errors.append(f"Veld '{field_name}': ontbrekende 'col' waarde")
        elif not isinstance(config["col"], str) or len(config["col"]) == 0:
            errors.append(f"Veld '{field_name}': 'col' moet non-empty string zijn")
        
        # Valideer visibility configuratie
        visibility_keys = {"visible", "visible_only", "visible_except"}
        present_visibility = visibility_keys.intersection(config.keys())
        
        if len(present_visibility) > 1 and "visible" in present_visibility:
            # visible samen met visible_only/except is toegestaan voor leesbaarheid
            pass
        
        if "visible" in config and config["visible"] not in self.VALID_VISIBILITY:
            errors.append(f"Veld '{field_name}': 'visible' moet 'always' of 'never' zijn")
        
        if "visible_only" in config and not isinstance(config["visible_only"], list):
            errors.append(f"Veld '{field_name}': 'visible_only' moet een lijst zijn")
        
        if "visible_except" in config and not isinstance(config["visible_except"], list):
            errors.append(f"Veld '{field_name}': 'visible_except' moet een lijst zijn")
        
        # Valideer mandatory configuratie
        mandatory_keys = {"mandatory", "mandatory_only", "mandatory_except"}
        present_mandatory = mandatory_keys.intersection(config.keys())
        
        if len(present_mandatory) > 1 and "mandatory" in present_mandatory:
            # mandatory samen met mandatory_only/except is toegestaan voor leesbaarheid
            pass
        
        if "mandatory" in config and config["mandatory"] not in self.VALID_MANDATORY:
            errors.append(f"Veld '{field_name}': 'mandatory' moet 'always' of 'never' zijn")
        
        if "mandatory_only" in config:
            if isinstance(config["mandatory_only"], str):
                # Accepteer enkele string en converteer naar lijst
                config["mandatory_only"] = [config["mandatory_only"]]
            elif not isinstance(config["mandatory_only"], list):
                errors.append(f"Veld '{field_name}': 'mandatory_only' moet een lijst of string zijn")
        
        if "mandatory_except" in config:
            if isinstance(config["mandatory_except"], str):
                # Accepteer enkele string en converteer naar lijst
                config["mandatory_except"] = [config["mandatory_except"]]
            elif not isinstance(config["mandatory_except"], list):
                errors.append(f"Veld '{field_name}': 'mandatory_except' moet een lijst of string zijn")
        
        # Valideer depends_on
        if "depends_on" in config:
            if not isinstance(config["depends_on"], list):
                errors.append(f"Veld '{field_name}': 'depends_on' moet een lijst zijn")
            else:
                for i, dep in enumerate(config["depends_on"]):
                    if not isinstance(dep, dict):
                        errors.append(f"Veld '{field_name}': depends_on[{i}] moet een dictionary zijn")
                    elif "field" not in dep:
                        errors.append(f"Veld '{field_name}': depends_on[{i}] moet 'field' sleutel hebben")
        
        # Valideer depends_trigger_for (optioneel veld)
        if "depends_trigger_for" in config and not isinstance(config["depends_trigger_for"], list):
            errors.append(f"Veld '{field_name}': 'depends_trigger_for' moet een lijst zijn")
        
        # Valideer notes
        if "notes" in config and not isinstance(config["notes"], str):
            errors.append(f"Veld '{field_name}': 'notes' moet een string zijn")
        
        return errors
    
    def get_field(self, field_name: str) -> Dict[str, Any]:
        """
        Krijg configuratie van een specifiek veld.
        
        Args:
            field_name: Naam van het veld
            
        Returns:
            Veld configuratie dictionary
            
        Raises:
            KeyError: Als veld niet bestaat
        """
        if field_name not in self.fields:
            raise KeyError(f"Veld '{field_name}' niet gevonden in mapping")
        
        return self.fields[field_name]
    
    def get_all_fields(self) -> Dict[str, Dict[str, Any]]:
        """
        Krijg alle veld configuraties.
        
        Returns:
            Dictionary met alle veld configuraties
        """
        return self.fields.copy()
    
    def get_columns(self) -> Set[str]:
        """
        Krijg alle Excel kolom letters.
        
        Returns:
            Set van kolom letters (A, B, C, etc.)
        """
        return {config["col"] for config in self.fields.values() if "col" in config}
    
    def get_field_by_column(self, column: str) -> str:
        """
        Vind veld naam op basis van Excel kolom.
        
        Args:
            column: Excel kolom letter (A, B, C, etc.)
            
        Returns:
            Veld naam
            
        Raises:
            KeyError: Als kolom niet gevonden
        """
        for field_name, config in self.fields.items():
            if config.get("col") == column:
                return field_name
        
        raise KeyError(f"Geen veld gevonden voor kolom '{column}'")
    
    def validate_dependencies(self) -> List[str]:
        """
        Valideer dat alle dependency references bestaan.
        
        Returns:
            Lijst van dependency fouten
        """
        errors = []
        all_field_names = set(self.fields.keys())
        
        for field_name, config in self.fields.items():
            # Check depends_on references
            for dep in config.get("depends_on", []):
                if "field" in dep and dep["field"] not in all_field_names:
                    errors.append(f"Veld '{field_name}' heeft dependency op onbekend veld '{dep['field']}'")
            
            # Check depends_trigger_for references
            for trigger_field in config.get("depends_trigger_for", []):
                if trigger_field not in all_field_names:
                    errors.append(f"Veld '{field_name}' triggert onbekend veld '{trigger_field}'")
        
        return errors
