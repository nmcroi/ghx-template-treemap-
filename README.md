# GHX Template Generator

Een tool voor het genereren van aangepaste Excel prijstemplates op basis van context-gevoelige veld configuraties voor de Nederlandse zorgsector.

## ✨ Functionaliteit

- **Context-gevoelig**: Genereert templates op basis van GS1-modus, producttype en instellingen
- **Veld management**: Automatische kolom zichtbaarheid en verplichte velden
- **Dependencies**: Geavanceerde afhankelijkheidslogica tussen velden
- **Template varianten**: Ondersteuning voor bestel-, verpakkings- en staffeltemplates
- **Stempel systeem**: Embedded metadata voor traceability
- **CLI interface**: Commandline tool voor automatisering

## 🚀 Snelstart

### Installatie

```bash
# Clone het project
git clone <repository-url>
cd "Project TemplateTree app v2"

# Installeer dependencies
pip install -r requirements.txt
```

### Basis gebruik

```bash
# Genereer template met GS1 medisch profiel
python -m src.main \
  --context tests/samples/sample_context_gs1.json \
  --out output/medisch_gs1_template.xlsx

# Genereer facilitair template
python -m src.main \
  --context tests/samples/sample_context_facilitair.json \
  --out output/facilitair_template.xlsx

# Toon info over bestaand template
python -m src.main --info existing_template.xlsx
```

## 📁 Projectstructuur

```
├── config/
│   ├── field_mapping.json          # Veld configuraties (A-CY kolommen)
│   └── context_schema.json         # JSON schema voor context validatie
├── src/
│   ├── main.py                     # CLI entrypoint
│   ├── context.py                  # Context datamodel
│   ├── mapping.py                  # Field mapping loader
│   ├── engine.py                   # Core beslislogica
│   ├── excel.py                    # Excel manipulatie
│   └── stamp.py                    # Metadata embedding
├── templates/
│   ├── template_besteleenheid.xlsx
│   ├── template_verpakkingseenheid.xlsx
│   └── template_staffel.xlsx
├── tests/
│   ├── test_engine.py              # Unit tests
│   └── samples/                    # Sample context bestanden
└── out/                            # Generated templates
```

## 🔧 Context Configuratie

Een context JSON bestand definieert de template parameters:

```json
{
  "template_choice": "custom",
  "gs1_mode": "gs1",
  "all_orderable": true,
  "product_type": "medisch",
  "has_chemicals": false,
  "is_staffel_file": false,
  "institutions": ["UMCU", "LUMC"],
  "version": "v1.0.0"
}
```

### Parameters

- **template_choice**: `"standard"` | `"custom"`
- **gs1_mode**: `"none"` | `"gs1"` | `"gs1_only"`
- **all_orderable**: `true` (bestelterminologie) | `false` (verpakkingsterminologie)
- **product_type**: `"medisch"` | `"lab"` | `"facilitair"` | `"mixed"`
- **has_chemicals**: Chemicaliën/safety velden actief
- **is_staffel_file**: Gebruik staffel template
- **institutions**: Array van instelling codes

## 🎯 Veld Mapping

De `config/field_mapping.json` definieert per kolom (A-CY):

```json
{
  "Artikelomschrijving Taal Code": {
    "col": "E",
    "visible_only": ["gs1"],
    "depends_on": [
      { "field": "Artikelomschrijving", "not_empty": true }
    ],
    "notes": "GS1 taalcode veld"
  }
}
```

### Veld Eigenschappen

- **col**: Excel kolom letter (A-CY)
- **visible**: `"always"` | `"never"`
- **visible_only**: Array van context labels (veld alleen zichtbaar bij deze contexten)
- **visible_except**: Array van context labels (veld verborgen bij deze contexten)
- **mandatory**: `"always"` | `"never"`
- **mandatory_only**: Array van context labels (verplicht alleen bij deze contexten)
- **mandatory_except**: Array van context labels (verplicht behalve bij deze contexten)
- **depends_on**: Array van dependencies
- **notes**: Menselijke uitleg

### Context Labels

Automatisch gegenereerd op basis van context:

- **GS1**: `"gs1"`, `"gs1_only"`, `"none"`
- **Product**: `"medisch"`, `"lab"`, `"facilitair"`, `"mixed"`
- **Features**: `"staffel"`, `"chemicals"`
- **Terminologie**: `"orderable_true"`, `"orderable_false"`
- **Instellingen**: `"UMCU"`, `"LUMC"`, `"AMC"`, etc.

## 🔗 Dependencies

Dependencies definiëren wanneer velden zinvol zijn:

```json
{
  "depends_on": [
    { "field": "CE Certificaat nummer", "not_empty": true },
    { "field": "Steriel", "equals": true },
    { "field": "Product Type", "in": ["A", "B"] }
  ]
}
```

Ondersteunde conditions:
- **not_empty**: Veld mag niet leeg zijn
- **equals**: Exacte waarde match
- **in**: Waarde moet in lijst staan
- **is_true**: Boolean true check

## 🏷️ Template Stempel

Gegenereerde templates bevatten embedded metadata:

- **Hidden sheet** `_GHX_META`: Volledige JSON context
- **Named range** `GHX_STAMP`: Compacte preset code (bijv. "MED-GS1-ORDER")

```bash
# Extraheer stempel info
python -m src.main --info generated_template.xlsx
```

## 🧪 Testing

```bash
# Run unit tests
python -m pytest tests/

# Test specifieke module
python tests/test_engine.py

# Valideer configuraties
python -m src.main --validate-context tests/samples/sample_context_gs1.json
python -m src.main --validate-mapping config/field_mapping.json
```

## 🎨 CLI Opties

```bash
python -m src.main [OPTIONS]

# Input/Output
--context, -c          Context JSON bestand
--mapping, -m          Field mapping JSON (default: config/field_mapping.json)
--templates, -t        Templates directory (default: templates/)
--out, -o              Output Excel bestand

# Template keuze
--prefer               Template voorkeur: bestel|verpakking|staffel|auto

# Styling
--mandatory-color      Hex kleur voor verplichte velden (default: #FFF2CC)
--hidden-color         Hex kleur voor verborgen velden (default: #EEEEEE)

# Utilities
--info, -i             Toon template informatie
--validate-context     Valideer context JSON
--validate-mapping     Valideer mapping JSON
--verbose, -v          Verbose output
```

## 📊 Voorbeelden

### GS1 Medisch Template
```bash
python -m src.main \
  --context tests/samples/sample_context_gs1.json \
  --out output/umcu_medisch_gs1.xlsx \
  --verbose
```

### Lab Template met Chemicaliën
```bash
python -m src.main \
  --context tests/samples/sample_context_lab_chemicals.json \
  --mandatory-color "#FFE6CC" \
  --out output/lab_chemicals.xlsx
```

### Staffel Template
```bash
python -m src.main \
  --context tests/samples/sample_context_staffel.json \
  --prefer staffel \
  --out output/staffel_lumc.xlsx
```

## 🔍 Template Validatie

```bash
# Bekijk template details
python -m src.main --info output/generated_template.xlsx

# Output:
# ✅ GHX Template gevonden
# 📄 Preset Code: MED-GS1-ORDER
# ✔️  Geldig: Ja
# 📋 Configuratie:
#    Template Keuze: custom
#    GS1 Modus: gs1
#    Product Type: medisch
#    Instellingen: UMCU, LUMC
```

## 🚀 Uitbreidingen

### Nieuwe Instellingen Toevoegen

Voeg toe aan `KNOWN_INSTITUTIONS` in `src/context.py`:

```python
KNOWN_INSTITUTIONS = {
    "UMCU", "LUMC", "AMC", "VUmc",
    "NIEUWE_INSTELLING"  # Voeg hier toe
}
```

### Nieuwe Context Labels

Voeg logica toe aan `Context.labels()` methode voor custom business rules.

### Custom Dependencies

Extend `engine._evaluate_condition()` voor geavanceerde dependency logica.

## 📝 Licentie

© 2024 GHX Template Generator Team