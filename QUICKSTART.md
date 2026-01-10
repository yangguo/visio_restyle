# Quick Reference Guide

## Installation
```bash
pip install -r requirements.txt
cp .env.example .env
# Edit .env and add your OPENAI_API_KEY
```

## Commands

### Full Conversion (One Command)
```bash
python -m visio_restyle.main convert INPUT.vsdx -t TEMPLATE.vsdx -o OUTPUT.vsdx
```

### Step-by-Step

#### 1. Extract Source Diagram
```bash
python -m visio_restyle.main extract INPUT.vsdx -o diagram.json
```

#### 2. Extract Template Masters
```bash
python -m visio_restyle.main extract-masters TEMPLATE.vsdx -o masters.json
```

#### 3. Generate Mapping
```bash
python -m visio_restyle.main map diagram.json masters.json -o mapping.json
```

#### 4. Rebuild Diagram
```bash
python -m visio_restyle.main rebuild INPUT.vsdx TEMPLATE.vsdx mapping.json -o OUTPUT.vsdx
```

## Options

### Convert Command
- `-t, --template`: Target template file (required)
- `-o, --output`: Output file (required)
- `-m, --model`: LLM model (default: gpt-4)
- `--save-intermediate`: Save intermediate JSON files

### Map Command
- `-m, --model`: LLM model to use
  - `gpt-4` (default, most accurate)
  - `gpt-4-turbo` (faster)
  - `gpt-3.5-turbo` (fastest, less accurate)

## Environment Variables

Create a `.env` file:
```
OPENAI_API_KEY=sk-your-key-here
LLM_MODEL=gpt-4
```

## File Formats

### Diagram JSON
```json
{
  "filename": "diagram.vsdx",
  "pages": [{
    "name": "Page-1",
    "shapes": [{
      "id": "1",
      "text": "Shape text",
      "master_name": "Process",
      "position": {"x": 4.25, "y": 8.5},
      "size": {"width": 1.5, "height": 0.75}
    }],
    "connectors": [{
      "id": "2",
      "from_shape": "1",
      "to_shape": "3"
    }]
  }]
}
```

### Masters JSON
```json
{
  "masters": [
    {"name": "ModernProcess", "id": "1", "description": "..."}
  ]
}
```

### Mapping JSON
```json
{
  "1": "ModernProcess",
  "3": "ModernDecision"
}
```

## Common Workflows

### Convert Single File
```bash
python -m visio_restyle.main convert input.vsdx -t template.vsdx -o output.vsdx
```

### Convert with Custom Mapping
```bash
# Extract and generate initial mapping
python -m visio_restyle.main extract input.vsdx -o diagram.json
python -m visio_restyle.main extract-masters template.vsdx -o masters.json
python -m visio_restyle.main map diagram.json masters.json -o mapping.json

# Edit mapping.json manually

# Rebuild with custom mapping
python -m visio_restyle.main rebuild input.vsdx template.vsdx mapping.json -o output.vsdx
```

### Batch Convert
```bash
for file in *.vsdx; do
  python -m visio_restyle.main convert "$file" -t template.vsdx -o "new_${file}"
done
```

## Testing
```bash
pytest tests/ -v
```

## Troubleshooting

**Problem**: ModuleNotFoundError
- **Solution**: `pip install -r requirements.txt`

**Problem**: API Key Error
- **Solution**: Set `OPENAI_API_KEY` in `.env` or environment

**Problem**: Master not found
- **Solution**: Check available masters: `python -m visio_restyle.main extract-masters template.vsdx -o masters.json`

**Problem**: Incorrect mappings
- **Solution**: Use step-by-step workflow and manually edit `mapping.json`

## Help
```bash
python -m visio_restyle.main --help
python -m visio_restyle.main convert --help
python -m visio_restyle.main extract --help
```
