# Development Guide

## Development Setup

1. Clone the repository:
```bash
git clone https://github.com/yangguo/visio_restyle.git
cd visio_restyle
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install in development mode:
```bash
pip install -e .
```

4. Set up environment variables:
```bash
cp .env.example .env
# Edit .env and add your OPENAI_API_KEY
```

## Running the Application

### Using Python Module

```bash
python -m visio_restyle.main --help
python -m visio_restyle.main convert input.vsdx -t template.vsdx -o output.vsdx
```

### Using Wrapper Script

```bash
python visio-restyle.py --help
python visio-restyle.py convert input.vsdx -t template.vsdx -o output.vsdx
```

## Project Structure

```
visio_restyle/
├── visio_restyle/              # Main package
│   ├── __init__.py
│   ├── main.py                 # CLI entry point
│   ├── visio_extractor.py      # Visio extraction logic
│   ├── llm_mapper.py           # LLM integration
│   └── visio_rebuilder.py      # Visio rebuilding logic
├── visio-restyle.py            # Wrapper script
├── setup.py                    # Package setup
├── requirements.txt            # Dependencies
├── .env.example                # Environment template
└── README.md                   # User documentation
```

## Architecture

### 1. Visio Extractor (`visio_extractor.py`)

**Purpose**: Parse Visio diagrams and extract structured information.

**Key Classes**:
- `VisioExtractor`: Main extraction class

**Key Methods**:
- `extract()`: Extract complete diagram structure
- `_extract_shapes()`: Extract shape information
- `_extract_connectors()`: Extract connector information
- `get_masters_from_visio()`: Extract available masters from a template

**Output Format**:
```json
{
  "filename": "diagram.vsdx",
  "pages": [
    {
      "name": "Page-1",
      "shapes": [...],
      "connectors": [...]
    }
  ]
}
```

### 2. LLM Mapper (`llm_mapper.py`)

**Purpose**: Use LLM to intelligently map source shapes to target masters.

**Key Classes**:
- `LLMMapper`: Main mapping class
- `ShapeMapping`: Pydantic model for single mapping
- `MappingResponse`: Pydantic model for complete response

**Key Methods**:
- `create_mapping()`: Generate mappings using LLM
- `_build_prompt()`: Build LLM prompt
- `_call_llm()`: Call OpenAI API
- `_parse_response()`: Parse and validate response

**API Integration**:
- Uses OpenAI Chat Completions API
- Supports structured JSON output
- Uses Pydantic for response validation

### 3. Visio Rebuilder (`visio_rebuilder.py`)

**Purpose**: Rebuild Visio diagram with new masters while preserving layout.

**Key Classes**:
- `VisioRebuilder`: Main rebuilding class

**Key Methods**:
- `rebuild()`: Rebuild complete diagram
- `_replace_shape_master()`: Replace shape master while preserving properties
- `_reglue_connectors()`: Maintain connector relationships

**Preservation**:
- Position (x, y coordinates)
- Size (width, height)
- Text content
- Connector relationships

### 4. Main CLI (`main.py`)

**Purpose**: Command-line interface for the application.

**Commands**:
1. `extract`: Extract diagram to JSON
2. `extract-masters`: Extract masters from template
3. `map`: Create shape mappings using LLM
4. `rebuild`: Rebuild diagram with new masters
5. `convert`: Full workflow (all steps combined)

## Key Dependencies

- **vsdx**: Python library for reading/writing Visio files
- **openai**: OpenAI API client
- **pydantic**: Data validation and parsing
- **python-dotenv**: Environment variable management

## Error Handling

Each module includes error handling for common issues:
- Missing or invalid Visio files
- API errors (network, authentication)
- Invalid JSON formats
- Missing masters in templates

## Extending the Application

### Adding New LLM Providers

To add support for other LLM providers (e.g., Anthropic, local models):

1. Create a new mapper class similar to `LLMMapper`
2. Implement the same interface: `create_mapping(diagram_data, target_masters)`
3. Update `main.py` to support the new provider

### Customizing Shape Mapping Logic

To customize how shapes are mapped:

1. Edit `llm_mapper.py`'s `_build_prompt()` method
2. Adjust the prompt to emphasize specific criteria
3. Update temperature or other API parameters for different behavior

### Supporting Additional Shape Properties

To preserve more shape properties:

1. Update `visio_extractor.py` to extract additional properties
2. Update `visio_rebuilder.py` to preserve those properties during rebuild

## Testing

### Manual Testing

1. Create or obtain test Visio files
2. Run extraction:
   ```bash
   python -m visio_restyle.main extract test.vsdx -o test.json
   ```
3. Verify JSON output structure
4. Test with different diagram types (flowcharts, network diagrams, etc.)

### Testing Without API Key

For testing extraction and rebuilding without LLM:

1. Extract diagram: `python -m visio_restyle.main extract ...`
2. Manually create mapping JSON
3. Rebuild: `python -m visio_restyle.main rebuild ...`

## Troubleshooting

### Import Errors

If you get `ModuleNotFoundError`:
```bash
pip install -r requirements.txt
```

### API Key Issues

The LLM configuration supports URL, model name, and API settings through environment variables.

Ensure your `.env` file has the required configuration:
```bash
# Required
OPENAI_API_KEY=sk-...

# Model Configuration
LLM_MODEL=gpt-4

# Optional API Settings
OPENAI_API_BASE=https://api.openai.com/v1
OPENAI_API_VERSION=2023-05-15
OPENAI_ORG_ID=org-...
OPENAI_TIMEOUT=60
OPENAI_MAX_RETRIES=3
```

Or set environment variables:
```bash
export OPENAI_API_KEY=sk-...
export LLM_MODEL=gpt-4
export OPENAI_API_BASE=https://api.openai.com/v1
```

**Configuration Examples:**

Azure OpenAI:
```bash
OPENAI_API_KEY=your_azure_key
OPENAI_API_BASE=https://your-resource.openai.azure.com/
OPENAI_API_VERSION=2023-05-15
LLM_MODEL=your-deployment-name
```

OpenAI-Compatible APIs:
```bash
OPENAI_API_KEY=not-needed
OPENAI_API_BASE=http://localhost:8080/v1
LLM_MODEL=mistral
```

### Visio File Compatibility

The vsdx library works best with:
- Visio 2013 and newer (.vsdx format)
- Standard shape types
- Avoid highly customized or legacy formats

## Future Enhancements

Possible improvements:
- Support for additional LLM providers
- GUI interface
- Batch processing
- Style transfer beyond just shape masters
- Custom mapping rules/constraints
- Undo/redo functionality
- Preview mode before full conversion
