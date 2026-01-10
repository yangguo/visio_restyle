# Copilot Instructions for visio_restyle

## Project Overview
Visio Restyle converts Visio diagrams (.vsdx) to different visual styles using LLM-powered shape mapping. The core workflow is: **Extract → Map → Rebuild**.

## Architecture & Data Flow

```
input.vsdx → VisioExtractor → JSON → LLMMapper/AutoMapper → mapping.json → VisioRebuilder → output.vsdx
                                           ↓
                              template.vsdx (MasterInjector)
```

### Core Modules ([visio_restyle/](visio_restyle/))
- **main.py** - CLI entry point with subcommands: `extract`, `extract-masters`, `map`, `rebuild`, `convert`
- **visio_extractor.py** - Parses .vsdx files, extracts shapes/connectors/masters to JSON using the `vsdx` library
- **llm_mapper.py** - Calls OpenAI API to map source shapes to target masters; uses Pydantic models for validation
- **auto_mapper.py** - Offline heuristic mapper using name normalization and synonym tables (fallback when no API key)
- **master_injector.py** - Injects template masters into source file via XML manipulation
- **visio_rebuilder.py** - Applies mappings, strips style overrides, copies theme/pagesheet from template

## Key Patterns

### JSON Format Conventions
- All JSON uses UTF-8 encoding, `indent=2`, `ensure_ascii=False`
- Shape mapping keys are **string IDs** (e.g., `{"1": "ModernProcess", "2": "ModernDecision"}`)
- Diagram data structure: `{ filename, pages: [{ name, shapes: [...], connectors: [...] }] }`

### Mapper Selection Logic (main.py `_select_mapper`)
```python
# Model value of "auto", "none", or "offline" → uses AutoMapper
# Missing OPENAI_API_KEY → falls back to AutoMapper
# Otherwise → uses LLMMapper
```

### XML Namespaces for Visio Manipulation
When editing .vsdx XML, always use the namespace prefix:
```python
NS = {"v": "http://schemas.microsoft.com/office/visio/2012/main"}
root.findall(".//v:Shape", NS)
```

### Style Override Stripping
`VisioRebuilder._strip_style_overrides()` removes specific sections (Fill, Line, QuickStyle) and cells to let template styling take effect. Extend `STYLE_SECTIONS_TO_REMOVE` / `STYLE_CELLS_TO_REMOVE` sets if needed.

## Development Commands
```bash
pip install -e .                  # Install in editable mode
python -m visio_restyle.main --help
pytest tests/ -v                  # Run unit tests

# Full conversion workflow
python -m visio_restyle.main convert input.vsdx -t template.vsdx -o output.vsdx --save-intermediate
```

## Environment Configuration
Copy `.env.example` to `.env`. Required: `OPENAI_API_KEY`. Optional: `LLM_MODEL`, `OPENAI_API_BASE`, `OPENAI_TIMEOUT`.

Use `--model auto` or set empty `OPENAI_API_KEY` to test without LLM calls.

## Testing Approach
- Unit tests in [tests/](tests/) focus on JSON serialization and data structure validation
- LLM-dependent behavior requires mocks or manual verification with sample files
- Use `--save-intermediate` flag to debug extraction/mapping JSON outputs

## Code Style
- 4-space indentation, `snake_case` for functions/modules, `PascalCase` for classes
- Docstrings on all public methods
- Error handling: raise descriptive exceptions, catch at CLI level in `main.py`
