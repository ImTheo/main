# Multiple File Column Mappings

Script CLI para actualizar valores en un archivo Excel `.xlsx` usando otro archivo `.xlsx` como matriz de consulta.

El script usa automáticamente la primera hoja de cada archivo y no usa `pandas`.

## Requisitos

- Python 3.8 o superior.
- `pip` para instalar dependencias.
- Archivos de entrada en formato `.xlsx`.

Para verificar la versión de Python:

```bash
python3 --version
```

En Windows:

```powershell
python --version
```

La dependencia externa del proyecto es `openpyxl`, instalada desde `requirements.txt`.

## Instalación

Desde esta carpeta:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

En Windows PowerShell:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Si `python3` no existe en tu sistema, usa `python` en los comandos.

## Uso

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

Todos los argumentos son obligatorios:

- `--dynamic-file`: archivo `.xlsx` que será actualizado.
- `--lookup-file`: archivo `.xlsx` usado como matriz de consulta.
- `--dynamic-header-row`: fila de encabezados en el archivo dynamic.
- `--lookup-header-row`: fila de encabezados en el archivo lookup.
- `--dynamic-match-column`: columna usada para buscar coincidencias en dynamic.
- `--dynamic-replace-column`: columna que se actualizará en dynamic.
- `--lookup-match-column`: columna usada para buscar coincidencias en lookup.
- `--lookup-value-column`: columna desde donde se toma el nuevo valor.

Argumentos opcionales:

- `--print-duplicated-lookup-keys`: imprime las claves duplicadas del lookup que tambien existen en dynamic, ya normalizadas.

## Columnas

Las columnas pueden indicarse de tres formas:

```bash
--dynamic-match-column "Código Proceso"
--dynamic-match-column A
--dynamic-match-column 1
```

Cuando se usa un nombre de encabezado, la comparación ignora mayúsculas/minúsculas y espacios al inicio o al final.

## Normalización de claves

La normalización se aplica solo a las columnas de coincidencia:

- `--dynamic-match-column`
- `--lookup-match-column`

Reglas:

- Valores vacíos se convierten en string vacío.
- El valor se convierte a texto.
- Se eliminan comillas dobles.
- Se eliminan espacios al inicio y al final.
- Si termina en `.0`, se convierte a entero textual.
- Se toma solo la primera palabra.
- Se convierte a minúsculas.

Ejemplos:

| Valor original | Clave normalizada |
| --- | --- |
| `"ABC123 contrato energia"` | `abc123` |
| `ABC123 contrato energia` | `abc123` |
| `"  ABC123  "` | `abc123` |
| `ABC123` | `abc123` |
| `123.0` | `123` |

El valor escrito en `--dynamic-replace-column` no se normaliza: se copia exactamente desde `--lookup-value-column`.

Las filas del archivo dynamic con clave no vacía pero sin coincidencia en el lookup se marcan en amarillo en el archivo de salida.

## Salida

El archivo original no se sobrescribe. Se genera un archivo nuevo junto al archivo dynamic:

```text
dynamic_matrix.xlsx -> dynamic_matrix_updated.xlsx
```

Si el archivo `_updated.xlsx` ya existe, se sobrescribe.

## Duplicados

Si el lookup contiene claves duplicadas, se usa el último valor encontrado y se muestra una advertencia:

```text
Warning: 4 duplicated lookup keys found. Last value was used.
```

Para ver cuales claves duplicadas del lookup tienen coincidencia en dynamic:

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

Salida adicional:

```text
Duplicated lookup keys matching dynamic file:
- abc123
- xyz789
```

## Ejemplo completo

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

Resultado:

```text
matriz_dinamica_updated.xlsx
```

Resumen esperado:

```text
Processed rows: 300
Updated rows: 248
Rows without match: 45
Skipped rows with empty match key: 7
Duplicated lookup keys: 3
Output file: matriz_dinamica_updated.xlsx
```

## Formatos soportados

Solo se aceptan archivos `.xlsx`.

No se aceptan `.xls`, `.xlsm`, `.csv` ni `.ods`.

Si se usa otro formato, el script se detiene con un error:

```text
Error: Only .xlsx files are supported: data.csv
```
