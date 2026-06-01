# Multiple File Column Mappings

CLI script to update values in an Excel `.xlsx` file using another `.xlsx` file as a lookup matrix.

The script automatically uses the first sheet in each file and does not use `pandas`.

## Requirements

- Python 3.8 or higher.
- `pip` to install dependencies.
- Input files in `.xlsx` format.

To check your Python version:

```bash
python3 --version
```

On Windows:

```powershell
python --version
```

The project's external dependency is `openpyxl`, installed from `requirements.txt`.

## Installation

From this folder:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

On Windows PowerShell:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

If `python3` is not available on your system, use `python` in the commands.

## Usage

```bash
python replace_from_lookup.py \
  --dynamic-file dynamic_matrix.xlsx \
  --lookup-file lookup_matrix.xlsx \
  --dynamic-header-row 1 \
  --lookup-header-row 1 \
  --dynamic-match-column "Código Proceso" \
  --dynamic-replace-column "Estado" \
  --lookup-match-column "Código" \
  --lookup-value-column "Nuevo Estado"
```

All arguments are required:

- `--dynamic-file`: `.xlsx` file that will be updated.
- `--lookup-file`: `.xlsx` file used as the lookup matrix.
- `--dynamic-header-row`: header row in the dynamic file.
- `--lookup-header-row`: header row in the lookup file.
- `--dynamic-match-column`: column used to find matches in the dynamic file.
- `--dynamic-replace-column`: column that will be updated in the dynamic file.
- `--lookup-match-column`: column used to find matches in the lookup file.
- `--lookup-value-column`: column that provides the new value.

Optional arguments:

- `--print-duplicated-lookup-keys`: prints normalized duplicated lookup keys that also exist in the dynamic file.

## Columns

Columns can be provided in three ways:

```bash
--dynamic-match-column "Código Proceso"
--dynamic-match-column A
--dynamic-match-column 1
```

When a header name is used, matching ignores letter case and leading or trailing spaces.

## Key Normalization

Normalization is applied only to the match columns:

- `--dynamic-match-column`
- `--lookup-match-column`

Rules:

- Empty values become an empty string.
- The value is converted to text.
- Double quotes are removed.
- Leading and trailing spaces are removed.
- If the value ends in `.0`, it is converted to text without the decimal part.
- Only the first word is used.
- The value is converted to lowercase.

Examples:

| Original value | Normalized key |
| --- | --- |
| `"ABC123 contrato energia"` | `abc123` |
| `ABC123 contrato energia` | `abc123` |
| `"  ABC123  "` | `abc123` |
| `ABC123` | `abc123` |
| `123.0` | `123` |

The value written to `--dynamic-replace-column` is not normalized: it is copied exactly from `--lookup-value-column`.

Rows in the dynamic file with a non-empty key but no match in the lookup file are highlighted in yellow in the output file.

## Output

The original file is not overwritten. A new file is generated next to the dynamic file:

```text
dynamic_matrix.xlsx -> dynamic_matrix_updated.xlsx
```

If the `_updated.xlsx` file already exists, it is overwritten.

## Duplicates

If the lookup file contains duplicated keys, the last value found is used and a warning is shown:

```text
Warning: 4 duplicated lookup keys found. Last value was used.
```

To see which duplicated lookup keys also have a match in the dynamic file:

```bash
python replace_from_lookup.py \
  --dynamic-file dynamic_matrix.xlsx \
  --lookup-file lookup_matrix.xlsx \
  --dynamic-header-row 1 \
  --lookup-header-row 1 \
  --dynamic-match-column "Código Proceso" \
  --dynamic-replace-column "Estado" \
  --lookup-match-column "Código" \
  --lookup-value-column "Nuevo Estado" \
  --print-duplicated-lookup-keys
```

Additional output:

```text
Duplicated lookup keys matching dynamic file:
- abc123
- xyz789
```

## Complete Example

```bash
python replace_from_lookup.py \
  --dynamic-file matriz_dinamica.xlsx \
  --lookup-file matriz_consulta.xlsx \
  --dynamic-header-row 2 \
  --lookup-header-row 1 \
  --dynamic-match-column "Código Proceso" \
  --dynamic-replace-column "Estado Actual" \
  --lookup-match-column "Código" \
  --lookup-value-column "Nuevo Estado"
```

Result:

```text
matriz_dinamica_updated.xlsx
```

Expected summary:

```text
Processed rows: 300
Updated rows: 248
Rows without match: 45
Skipped rows with empty match key: 7
Duplicated lookup keys: 3
Output file: matriz_dinamica_updated.xlsx
```

## Supported Formats

Only `.xlsx` files are accepted.

`.xls`, `.xlsm`, `.csv`, and `.ods` files are not accepted.

If another format is used, the script stops with an error:

```text
Error: Only .xlsx files are supported: data.csv
```
