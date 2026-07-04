# docx-editor

CLI tool to edit `.docx` files programmatically — useful for LLMs, scripts, and batch automation.

## Features

- **Find and replace** text while preserving original formatting (bold, italic, color, etc.)
- **Smart mode** (`--smart`) matches text across run boundaries — essential for Jinja-style `{{ }}` templates that Word fragments across multiple runs
- **Suffix preservation**: when replacing text at run boundaries (e.g. `True }}` → `False }}`), unchanged trailing runs like `}}` are kept intact
- **Inline backup**: editing in-place creates a `.bak` backup automatically
- **Document inspection**: list paragraphs and tables (with cell contents) to understand document structure before editing
- **Table support**: replace text in tables, set cell values, change font sizes
- **JSON output** (`--json`) for machine consumption

Run `docx-editor --help` (or `docx-editor <command> --help`) to explore all commands and options.
