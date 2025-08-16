"""
Context datamodel voor GHX Template Generator.

Bevat de Context dataclass en validatie logica voor input parameters.
"""

from dataclasses import dataclass
from typing import List, Literal, Set
import json
from pathlib import Path

# Type definitios volgens specificatie
GS1Mode = Literal["none", "gs1", "gs1_only"]
ProductType = Literal["medisch", "lab", "facilitair", "mixed"]
TemplateChoice = Literal["standard", "custom"]

# Bekende instellingen (uitbreidbaar)
KNOWN_INSTITUTIONS = {
    "UMCU", "LUMC", "AMC", "VUmc", "Erasmus MC", "MUMC", "UMC Groningen",
    "Radboudumc", "Isala", "MST", "Catharina", "Elisabeth-TweeSteden",
    "Franciscus", "HagaZiekenhuis", "HMC", "Jeroen Bosch", "Maasstad",
    "Medisch Spectrum Twente", "OLVG", "Reinier de Graaf", "Rijnstate",
    "Sint Antonius", "Sint Franciscus", "Spaarne Gasthuis", "Tergooi",
    "Zuyderland", "ZGT", "Ziekenhuis Gelderse Vallei", "Zorggroep Twente",
    "Admiraal De Ruyter", "Albert Schweitzer", "Alrijne", "Amphia"
}


@dataclass
class Context:
    """
    Context object dat alle input parameters bevat voor template generatie.
    
    Attributes:
        template_choice: Standard of custom template
        gs1_mode: GS1 modus (none/gs1/gs1_only)
        all_orderable: Bestelbaar terminologie (true) vs verpakking (false)
        product_type: Product categorie
        has_chemicals: Chemicaliën/safety velden actief
        is_staffel_file: Gebruik staffel template
        institutions: Lijst van instellingen
        version: Config/generator versie
    """
    template_choice: TemplateChoice
    gs1_mode: GS1Mode
    all_orderable: bool
    product_type: ProductType
    has_chemicals: bool
    is_staffel_file: bool
    institutions: List[str]
    version: str = "v1.0.0"

    def labels(self) -> Set[str]:
        """
        Genereer set van context labels voor gebruik in veld mapping.
        
        Returns:
            Set van strings die gebruikt worden in visible_only/except en mandatory_only/except
        """
        labs = {self.product_type}
        
        # GS1 labels
        if self.gs1_mode in ("gs1", "gs1_only"):
            labs.add("gs1")
        else:
            labs.add("none")
            
        if self.gs1_mode == "gs1_only":
            labs.add("gs1_only")
        
        # Template type
        if self.is_staffel_file:
            labs.add("staffel")
        
        # Terminologie
        labs.add("orderable_true" if self.all_orderable else "orderable_false")
        
        # Chemicaliën  
        if self.has_chemicals:
            labs.add("chemicals")
        
        # Instellingen
        labs.update(self.institutions)
        
        return labs

    def validate(self) -> List[str]:
        """
        Valideer context object en return lijst van fouten.
        
        Returns:
            Lijst van foutmeldingen (leeg als geldig)
        """
        errors = []
        
        # Valideer instellingen
        unknown_institutions = set(self.institutions) - KNOWN_INSTITUTIONS
        if unknown_institutions:
            errors.append(f"Onbekende instellingen: {', '.join(unknown_institutions)}")
        
        # Logische validaties
        if self.gs1_mode == "gs1_only" and not self.all_orderable:
            errors.append("GS1-only modus vereist orderable terminologie")
        
        if self.is_staffel_file and self.gs1_mode == "gs1_only":
            errors.append("Staffel templates zijn niet compatibel met GS1-only modus")
        
        return errors

    @classmethod
    def from_json_file(cls, file_path: Path) -> "Context":
        """
        Laad context van JSON bestand.
        
        Args:
            file_path: Pad naar JSON bestand
            
        Returns:
            Context object
            
        Raises:
            FileNotFoundError: Als bestand niet bestaat
            ValueError: Als JSON ongeldig of validatie faalt
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            context = cls(**data)
            errors = context.validate()
            
            if errors:
                raise ValueError(f"Context validatie gefaald: {'; '.join(errors)}")
            
            return context
            
        except FileNotFoundError:
            raise FileNotFoundError(f"Context bestand niet gevonden: {file_path}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Ongeldig JSON in {file_path}: {e}")
        except TypeError as e:
            raise ValueError(f"Ongeldig context formaat in {file_path}: {e}")

    def to_dict(self) -> dict:
        """Convert naar dictionary voor serialisatie."""
        return {
            "template_choice": self.template_choice,
            "gs1_mode": self.gs1_mode,
            "all_orderable": self.all_orderable,
            "product_type": self.product_type,
            "has_chemicals": self.has_chemicals,
            "is_staffel_file": self.is_staffel_file,
            "institutions": self.institutions,
            "version": self.version
        }

    def get_template_basename(self) -> str:
        """
        Bepaal welke basis template te gebruiken.
        
        Returns:
            Template bestandsnaam zonder extensie
        """
        if self.is_staffel_file:
            return "template_staffel"
        elif self.all_orderable:
            return "template_besteleenheid"
        else:
            return "template_verpakkingseenheid"

    def get_preset_code(self) -> str:
        """
        Genereer compacte preset code voor stempel.
        
        Returns:
            Compacte string zoals "MED-GS1-ORDER"
        """
        parts = []
        
        # Product type
        type_codes = {
            "medisch": "MED",
            "lab": "LAB", 
            "facilitair": "FAC",
            "mixed": "MIX"
        }
        parts.append(type_codes[self.product_type])
        
        # GS1 modus
        if self.gs1_mode == "gs1_only":
            parts.append("GS1ONLY")
        elif self.gs1_mode == "gs1":
            parts.append("GS1")
        else:
            parts.append("COMM")
        
        # Terminologie
        parts.append("ORDER" if self.all_orderable else "PACK")
        
        # Staffel
        if self.is_staffel_file:
            parts.append("STAFF")
        
        # Chemicaliën
        if self.has_chemicals:
            parts.append("CHEM")
        
        return "-".join(parts)
