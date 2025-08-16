"""
Core engine voor GHX Template Generator.

Bevat de beslislogica voor veld zichtbaarheid, verplichte velden en dependencies.
"""

from typing import Dict, Any, Set, List, Optional
from dataclasses import dataclass
from .context import Context
from .mapping import FieldMapping


@dataclass
class FieldDecision:
    """
    Beslissing voor een enkel veld.
    
    Attributes:
        field_name: Naam van het veld
        column: Excel kolom letter
        visible: Of het veld zichtbaar moet zijn
        mandatory: Of het veld verplicht is (alleen relevant als visible=True)
        dependency_satisfied: Of dependencies zijn voldaan
        notes: Opmerkingen voor header
        machine_notes: Machine-leesbare uitleg voor debugging
    """
    field_name: str
    column: str
    visible: bool
    mandatory: bool
    dependency_satisfied: bool
    notes: str
    machine_notes: str


class TemplateEngine:
    """
    Core engine voor template beslissingen.
    """
    
    def __init__(self, context: Context, field_mapping: FieldMapping):
        """
        Initialiseer engine.
        
        Args:
            context: Context object met input parameters
            field_mapping: Field mapping configuratie
        """
        self.context = context
        self.field_mapping = field_mapping
        self.context_labels = context.labels()
    
    def process_all_fields(self, row_data: Optional[Dict[str, Any]] = None) -> List[FieldDecision]:
        """
        Verwerk alle velden en maak beslissingen.
        
        Args:
            row_data: Optionele rij data voor dependency evaluatie (voor toekomstige uitbreiding)
            
        Returns:
            Lijst van FieldDecision objecten
        """
        decisions = []
        
        for field_name, field_config in self.field_mapping.get_all_fields().items():
            decision = self._process_field(field_name, field_config, row_data or {})
            decisions.append(decision)
        
        return decisions
    
    def _process_field(self, field_name: str, field_config: Dict[str, Any], row_data: Dict[str, Any]) -> FieldDecision:
        """
        Verwerk een enkel veld.
        
        Args:
            field_name: Naam van het veld
            field_config: Veld configuratie uit mapping
            row_data: Rij data voor dependency evaluatie
            
        Returns:
            FieldDecision object
        """
        column = field_config["col"]
        
        # Bepaal zichtbaarheid
        visible = self._is_visible(field_config)
        
        # Bepaal dependency status
        dependency_satisfied = self._deps_satisfied(field_config, row_data)
        
        # Bepaal verplichte status (alleen als zichtbaar)
        mandatory = False
        if visible:
            mandatory = self._is_mandatory(field_config)
        
        # Genereer opmerkingen
        notes = self._generate_notes(field_config, visible, mandatory, dependency_satisfied)
        machine_notes = self._generate_machine_notes(field_config, visible, mandatory, dependency_satisfied)
        
        return FieldDecision(
            field_name=field_name,
            column=column,
            visible=visible,
            mandatory=mandatory,
            dependency_satisfied=dependency_satisfied,
            notes=notes,
            machine_notes=machine_notes
        )
    
    def _is_visible(self, field_config: Dict[str, Any]) -> bool:
        """
        Bepaal of veld zichtbaar moet zijn.
        
        Args:
            field_config: Veld configuratie
            
        Returns:
            True als veld zichtbaar moet zijn
        """
        # visible_only heeft voorrang
        if "visible_only" in field_config:
            return any(label in self.context_labels for label in field_config["visible_only"])
        
        # visible_except heeft voorrang
        if "visible_except" in field_config:
            return not any(label in self.context_labels for label in field_config["visible_except"])
        
        # Fallback naar visible
        return field_config.get("visible", "always") == "always"
    
    def _is_mandatory(self, field_config: Dict[str, Any]) -> bool:
        """
        Bepaal of veld verplicht moet zijn.
        
        Args:
            field_config: Veld configuratie
            
        Returns:
            True als veld verplicht moet zijn
        """
        # mandatory_only heeft voorrang
        if "mandatory_only" in field_config:
            return any(label in self.context_labels for label in field_config["mandatory_only"])
        
        # mandatory_except heeft voorrang
        if "mandatory_except" in field_config:
            return not any(label in self.context_labels for label in field_config["mandatory_except"])
        
        # Fallback naar mandatory
        return field_config.get("mandatory", "never") == "always"
    
    def _deps_satisfied(self, field_config: Dict[str, Any], row_data: Dict[str, Any]) -> bool:
        """
        Controleer of dependencies zijn voldaan.
        
        Args:
            field_config: Veld configuratie
            row_data: Rij data voor evaluatie
            
        Returns:
            True als alle dependencies zijn voldaan
        """
        conditions = field_config.get("depends_on", [])
        
        # Geen dependencies = altijd voldaan
        if not conditions:
            return True
        
        # Alle conditions moeten waar zijn (AND logica)
        for condition in conditions:
            if not self._evaluate_condition(condition, row_data):
                return False
        
        return True
    
    def _evaluate_condition(self, condition: Dict[str, Any], row_data: Dict[str, Any]) -> bool:
        """
        Evalueer een enkele dependency condition.
        
        Args:
            condition: Dependency condition dictionary
            row_data: Rij data voor evaluatie
            
        Returns:
            True als condition is voldaan
        """
        field_name = condition.get("field")
        if not field_name:
            return False
        
        value = row_data.get(field_name)
        
        # not_empty check
        if condition.get("not_empty", False):
            if value is None or str(value).strip() == "":
                return False
        
        # equals check
        if "equals" in condition:
            if value != condition["equals"]:
                return False
        
        # is_true check
        if "is_true" in condition:
            if not bool(value):
                return False
        
        # in check
        if "in" in condition:
            if value not in condition["in"]:
                return False
        
        return True
    
    def _generate_notes(self, field_config: Dict[str, Any], visible: bool, mandatory: bool, deps_satisfied: bool) -> str:
        """
        Genereer menselijke opmerkingen voor header.
        
        Args:
            field_config: Veld configuratie
            visible: Of veld zichtbaar is
            mandatory: Of veld verplicht is
            deps_satisfied: Of dependencies zijn voldaan
            
        Returns:
            Opmerking string
        """
        base_notes = field_config.get("notes", "")
        
        additional_notes = []
        
        if not visible:
            additional_notes.append("VERBORGEN in huidige context")
        elif mandatory:
            additional_notes.append("VERPLICHT in huidige context")
        
        if field_config.get("depends_on") and not deps_satisfied:
            additional_notes.append("Dependency NIET voldaan")
        
        if additional_notes:
            separator = " | " if base_notes else ""
            return base_notes + separator + " | ".join(additional_notes)
        
        return base_notes
    
    def _generate_machine_notes(self, field_config: Dict[str, Any], visible: bool, mandatory: bool, deps_satisfied: bool) -> str:
        """
        Genereer machine-leesbare debug informatie.
        
        Args:
            field_config: Veld configuratie
            visible: Of veld zichtbaar is
            mandatory: Of veld verplicht is
            deps_satisfied: Of dependencies zijn voldaan
            
        Returns:
            Machine-leesbare string voor debugging
        """
        parts = []
        
        # Visibility logica
        if "visible_only" in field_config:
            matched = [label for label in field_config["visible_only"] if label in self.context_labels]
            parts.append(f"VISIBLE_ONLY:{','.join(field_config['visible_only'])}->matched:{','.join(matched)}")
        elif "visible_except" in field_config:
            forbidden = [label for label in field_config["visible_except"] if label in self.context_labels]
            parts.append(f"VISIBLE_EXCEPT:{','.join(field_config['visible_except'])}->forbidden:{','.join(forbidden)}")
        else:
            parts.append(f"VISIBLE:{field_config.get('visible', 'always')}")
        
        # Mandatory logica
        if visible:
            if "mandatory_only" in field_config:
                matched = [label for label in field_config["mandatory_only"] if label in self.context_labels]
                parts.append(f"MANDATORY_ONLY:{','.join(field_config['mandatory_only'])}->matched:{','.join(matched)}")
            elif "mandatory_except" in field_config:
                forbidden = [label for label in field_config["mandatory_except"] if label in self.context_labels]
                parts.append(f"MANDATORY_EXCEPT:{','.join(field_config['mandatory_except'])}->forbidden:{','.join(forbidden)}")
            else:
                parts.append(f"MANDATORY:{field_config.get('mandatory', 'never')}")
        
        # Dependencies
        if field_config.get("depends_on"):
            deps_str = "|".join([f"{dep.get('field')}:{','.join(str(v) for k,v in dep.items() if k!='field')}" 
                                for dep in field_config["depends_on"]])
            parts.append(f"DEPS:{deps_str}->satisfied:{deps_satisfied}")
        
        return " | ".join(parts)
    
    def get_decisions_by_column(self, decisions: List[FieldDecision]) -> Dict[str, FieldDecision]:
        """
        Converteer beslissingen naar dictionary geïndexeerd op kolom.
        
        Args:
            decisions: Lijst van FieldDecision objecten
            
        Returns:
            Dictionary met kolom -> FieldDecision mapping
        """
        return {decision.column: decision for decision in decisions}
    
    def get_visible_columns(self, decisions: List[FieldDecision]) -> Set[str]:
        """
        Krijg set van zichtbare kolommen.
        
        Args:
            decisions: Lijst van FieldDecision objecten
            
        Returns:
            Set van kolom letters die zichtbaar moeten zijn
        """
        return {decision.column for decision in decisions if decision.visible}
    
    def get_mandatory_columns(self, decisions: List[FieldDecision]) -> Set[str]:
        """
        Krijg set van verplichte kolommen.
        
        Args:
            decisions: Lijst van FieldDecision objecten
            
        Returns:
            Set van kolom letters die verplicht moeten zijn
        """
        return {decision.column for decision in decisions if decision.visible and decision.mandatory}
