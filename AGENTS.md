# Repository Guidelines

## Project Structure & Module Organization
- `visio_restyle/` contains the Python package: `main.py` (CLI), `visio_extractor.py`, `llm_mapper.py`, `visio_rebuilder.py`.
- `visio-restyle.py` is a wrapper script; the installed CLI entry point is `visio-restyle`.
- `tests/` holds pytest tests; `examples/` contains runnable workflows; sample assets live at the repo root (`input.vsdx`, `template.vsdx`, `template.jpg`).
- Primary docs: `README.md`, `QUICKSTART.md`, and `DEVELOPMENT.md`.

## Build, Test, and Development Commands
- `pip install -r requirements.txt` installs runtime dependencies.
- `pip install -e .` installs the package in editable mode for development.
- `python -m visio_restyle.main --help` (or `visio-restyle --help`) shows CLI usage.
- `python -m visio_restyle.main convert input.vsdx -t template.vsdx -o output.vsdx` runs the full extract → map → rebuild workflow.
- `pytest tests/ -v` runs unit tests.

## Coding Style & Naming Conventions
- Python, 4-space indentation; keep formatting consistent with existing files and docstrings (no project-specific formatter is configured).
- Naming: modules and functions in `snake_case`, classes in `PascalCase`, constants in `UPPER_SNAKE_CASE`.
- JSON outputs are UTF-8, `indent=2`; mapping keys are string shape IDs.

## Testing Guidelines
- Use pytest; test files follow `test_*.py` naming in `tests/`.
- Prefer fast unit tests for JSON/mapping logic; note any required manual validation for Visio I/O.
- If behavior depends on the LLM, add test data or mocks and document any manual verification steps.

## Commit & Pull Request Guidelines
- Use Conventional Commit-style messages where possible (e.g., `feat:`, `fix:`, `docs:`); recent history uses `feat: ...`.
- PRs should include: a clear description, testing notes, linked issues, and sample outputs or screenshots for visual diagram changes.
- Update docs when CLI behavior or configuration changes.

## Configuration & Secrets
- Copy `.env.example` to `.env` and set `OPENAI_API_KEY` and related LLM settings; never commit secrets.
- Visio file manipulation is Windows-focused; call out any platform assumptions or limitations in your PR.
