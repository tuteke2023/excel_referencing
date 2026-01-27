# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TB-GL Linker is a Python tool that automatically links Trial Balance (TB) accounts to General Ledger (GL) details by creating hyperlinks in Excel files. It uses fuzzy string matching to correlate accounts and supports two linking modes: standard (links to account headers) and Net Movement (links to calculated net movement figures with dynamic cell references).

## Commands

### Run Web Application
```bash
streamlit run web_app.py
```

### Run CLI Tools
```bash
# Standard linker
python3 tb_gl_linker.py TB.xlsx GL.xlsx -o output.xlsx

# Net Movement version
python3 tb_gl_linker_netmovement.py TB.xlsx GL.xlsx -o output.xlsx

# Quick version
python3 quick_link.py TB.xlsx GL.xlsx
```

### Build Executable
```bash
python3 build_exe.py
```

### Docker
```bash
docker build -t tb-gl-linker .
docker run -p 8501:8501 tb-gl-linker
```

### Install Dependencies
```bash
pip install -r requirements.txt
```

## Architecture

### Core Files
- `web_app.py` - Streamlit web UI with Net Movement linking (primary user interface)
- `claude_analyzer.py` - Claude Code headless integration for AI-powered structure detection
- `excel_converter.py` - Excel to text conversion utilities for Claude analysis
- `prompt_templates.py` - Centralized prompts for Claude analysis
- `tb_gl_linker.py` - CLI tool linking to account headers
- `tb_gl_linker_netmovement.py` - CLI tool linking to Net Movement rows

### Two Linking Modes

| Mode | Links To | Hyperlink Display | Use Case |
|------|----------|-------------------|----------|
| Standard | Account headers in GL | Static "View Details" text | General navigation |
| Net Movement | Net Movement figures | Dynamic cell value | Account analysis |

### Processing Pipeline
1. Load workbooks (AI-powered or keyword-based TB/GL sheet detection)
2. Analyze TB structure (AI-powered or keyword-based column detection)
3. Analyze GL structure (AI-powered or keyword-based account section detection)
4. Match accounts using `difflib.SequenceMatcher` (80% similarity threshold)
5. Copy GL sheet to TB workbook preserving formatting
6. Add hyperlinks in new Reference column
7. Save combined workbook

### Claude Code Headless Integration (v1.2+)
The web app now supports AI-powered structure detection using Claude Code headless mode (`claude -p`). This allows the tool to intelligently detect file structures from various accounting software exports.

**How it works:**
1. Excel data is converted to text format using `excel_converter.py`
2. Structured prompts from `prompt_templates.py` are sent to Claude via subprocess
3. Claude analyzes the data and returns JSON with detected structure
4. Falls back to keyword-based detection if Claude is unavailable or times out

**Supported accounting software:**
- QuickBooks
- Sage
- Xero
- NetSuite
- Other formats with TB/GL export capability

**Running tests:**
```bash
python3 test_claude_integration.py
```

### Key Implementation Details
- **Sheet detection**: AI-powered analysis or keyword search for "TB", "Trial Balance", "GL", "General Ledger"
- **Header detection**: AI-powered analysis or scan first 20 rows for "account", "debit", "credit"
- **Summary row detection**: AI detects software-specific patterns ("Net Movement", "Balance", "Total")
- **Fuzzy matching**: Uses `SequenceMatcher.ratio()` with 80% threshold; exact case-insensitive matches score 1.0
- **Dynamic hyperlinks** (Net Movement mode): Formula-based `=HYPERLINK("#'Sheet'!Cell", 'Sheet'!Cell)` that displays actual values
- **GL workbook loading**: Uses `data_only=True` to read calculated values from formulas
- **Fallback strategy**: If Claude CLI unavailable, automatically falls back to keyword-based detection

### Dependencies
- `streamlit` - Web framework
- `openpyxl` (>=3.0.0) - Excel file manipulation
- `pandas` - Data analysis
- `difflib` - Fuzzy string matching (stdlib)
