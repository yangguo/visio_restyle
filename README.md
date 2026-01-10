# Visio Restyle

Convert Visio diagrams to different visual styles using LLM-powered shape mapping.

## Overview

This tool allows you to automatically convert a Visio diagram from one visual style to another while preserving the layout, text, and connections. It uses an LLM (Large Language Model) to intelligently map shapes from your source diagram to masters in a target template.

## Features

- **Extract**: Parse Visio diagrams into structured JSON format
- **Map**: Use LLM to intelligently map source shapes to target masters
- **Rebuild**: Generate new Visio diagram with target style while preserving:
  - Shape positions and sizes
  - Text content
  - Connector relationships
- **Full Workflow**: One-command conversion from source to target style

## How It Works

1. **Extract**: Script reads your source diagram and extracts shape list, text content, and connector information
2. **Map**: Extracted data + target template masters are sent to an LLM, which returns intelligent shape mappings
3. **Rebuild**: Script creates a new diagram using the mappings, preserving position, size, text, and re-gluing connectors

## Installation

### Prerequisites

- Python 3.8 or higher
- Windows OS (for Visio file manipulation)
- OpenAI API key

### Setup

1. Clone the repository:
```bash
git clone https://github.com/yangguo/visio_restyle.git
cd visio_restyle
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

Or install in development mode:
```bash
pip install -e .
```

3. Set up your OpenAI API key:
```bash
cp .env.example .env
# Edit .env and add your OPENAI_API_KEY
```

## Usage

### Quick Start - Full Conversion

Convert a diagram in one command:

```bash
visio-restyle convert input.vsdx -t template.vsdx -o output.vsdx
```

Options:
- `-t, --template`: Target template .vsdx file with desired masters
- `-o, --output`: Output file path
- `-m, --model`: LLM model to use (default: gpt-4)
- `--save-intermediate`: Save intermediate JSON files for debugging

### Step-by-Step Workflow

For more control, use the individual commands:

#### 1. Extract source diagram to JSON

```bash
visio-restyle extract input.vsdx -o diagram.json
```

#### 2. Extract available masters from target template

```bash
visio-restyle extract-masters template.vsdx -o masters.json
```

#### 3. Generate shape mappings using LLM

```bash
visio-restyle map diagram.json masters.json -o mapping.json
```

You can manually edit `mapping.json` to adjust the mappings if needed.

#### 4. Rebuild diagram with new masters

```bash
visio-restyle rebuild input.vsdx template.vsdx mapping.json -o output.vsdx
```

## Examples

### Example 1: Convert flowchart to modern style

```bash
visio-restyle convert old_flowchart.vsdx \
  -t modern_template.vsdx \
  -o modern_flowchart.vsdx \
  --save-intermediate
```

### Example 2: Use step-by-step for fine-tuning

```bash
# Extract and map
visio-restyle extract diagram.vsdx -o diagram.json
visio-restyle extract-masters new_style.vsdx -o masters.json
visio-restyle map diagram.json masters.json -o mapping.json

# Edit mapping.json manually if needed

# Rebuild with custom mapping
visio-restyle rebuild diagram.vsdx new_style.vsdx mapping.json -o result.vsdx
```

## Configuration

### Environment Variables

Create a `.env` file in the project root:

```bash
# Required: OpenAI API Key
OPENAI_API_KEY=your_api_key_here

# Optional: Default LLM Model
LLM_MODEL=gpt-4

# Optional: OpenAI Organization ID
# OPENAI_ORG_ID=your_org_id_here

# Optional: Custom OpenAI API Base URL (for Azure OpenAI or compatible APIs)
# OPENAI_API_BASE=https://api.openai.com/v1

# Optional: OpenAI API Version (for Azure OpenAI)
# OPENAI_API_VERSION=2023-05-15

# Optional: Request timeout in seconds
# OPENAI_TIMEOUT=60

# Optional: Max retries for API calls
# OPENAI_MAX_RETRIES=3
```

### Supported LLM Models

- `gpt-4` (default, recommended)
- `gpt-4-turbo`
- `gpt-3.5-turbo` (faster but less accurate)

### Using Azure OpenAI or Compatible APIs

To use Azure OpenAI or other OpenAI-compatible APIs, configure the base URL:

```bash
OPENAI_API_KEY=your_azure_key
OPENAI_API_BASE=https://your-resource.openai.azure.com/
OPENAI_API_VERSION=2023-05-15
LLM_MODEL=your-deployment-name
```

## Project Structure

```
visio_restyle/
├── visio_restyle/
│   ├── __init__.py
│   ├── main.py              # CLI application
│   ├── visio_extractor.py   # Extract diagram to JSON
│   ├── llm_mapper.py        # LLM-based shape mapping
│   └── visio_rebuilder.py   # Rebuild diagram with new masters
├── requirements.txt
├── setup.py
├── .env.example
└── README.md
```

## JSON Format

### Extracted Diagram Format

```json
{
  "filename": "diagram.vsdx",
  "pages": [
    {
      "name": "Page-1",
      "shapes": [
        {
          "id": "1",
          "text": "Start",
          "master_name": "Process",
          "position": {"x": 4.25, "y": 8.5},
          "size": {"width": 1.5, "height": 0.75}
        }
      ],
      "connectors": [
        {
          "id": "2",
          "from_shape": "1",
          "to_shape": "3"
        }
      ]
    }
  ]
}
```

### Shape Mapping Format

```json
{
  "1": "ModernProcess",
  "3": "ModernDecision",
  "5": "ModernData"
}
```

## Troubleshooting

### Common Issues

**Issue**: `ModuleNotFoundError: No module named 'vsdx'`
- Solution: Run `pip install -r requirements.txt`

**Issue**: `ValueError: OpenAI API key must be provided`
- Solution: Set `OPENAI_API_KEY` in `.env` file or environment variable

**Issue**: Master not found in target template
- Solution: Check target template has the required masters, or manually edit mapping JSON

**Issue**: Connectors not properly glued
- Solution: The vsdx library handles connections automatically. If issues persist, try re-saving the output file in Visio.

## Limitations

- Requires Windows OS for full Visio file compatibility
- Some complex shape properties may not transfer perfectly
- Custom formatting (colors, line styles) are not preserved
- Works best with standard flowchart and diagram shapes

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License

## Credits

Built with:
- [vsdx](https://github.com/dave-howard/vsdx) - Python library for reading/writing .vsdx files
- [OpenAI API](https://openai.com/) - LLM-powered shape mapping