# Local VN to JP Translation Agent

Local Vietnamese -> Japanese Markdown translation with
`translategemma:4b` through Ollama by default.

## Structure

```text
JP-VN-translation-agent/
├─ input/
├─ output/
├─ run_translate_all.py
├─ excel_translate.py
├─ glossary.json
├─ requirements.txt
├─ setup.bat
├─ run.bat
└─ configure_ollama_shared.bat
```

## First setup

1. Double-click `setup.bat`.
2. The script creates `.venv`, installs the Python Ollama client,
   checks the Ollama runtime/server, and downloads
   `translategemma:4b` if necessary.

The model download is large, so the first setup requires an Internet
connection and enough free disk space.

## Shared-machine Ollama configuration

If this machine also runs other local Ollama agents/models, run
`configure_ollama_shared.bat` once.

It sets:

```text
OLLAMA_MAX_LOADED_MODELS=1
OLLAMA_NUM_PARALLEL=1
```

Then fully quit Ollama from the Windows system tray and start it again.

The translator keeps TranslateGemma loaded while the batch is active, then
explicitly unloads it when the batch finishes or exits with an error. This
releases RAM for other local Ollama tools.

## Translate Markdown

1. Copy `.md` / `.markdown` files into `input`.
2. Double-click `run.bat`.
3. Results are written under `output`.

The translator recursively scans subdirectories and preserves their relative
directory structure.

Example:

```text
input/
├─ a.md
└─ design/
   ├─ b.md
   └─ api/
      └─ c.md
```

becomes:

```text
output/
├─ a.md
└─ design/
   ├─ b.md
   └─ api/
      └─ c.md
```

## Markdown protection

The translator protects structural/literal content such as:

- YAML front matter
- fenced code blocks
- inline code
- Markdown link destinations and raw URLs
- HTML/XML tags and comments
- Markdown table separator rows
- common template/variable expressions

Each chunk is validated after translation. Headings, list items, blockquotes,
table pipes, and protected placeholders must remain intact. Minor line-wrap
differences are allowed.

## Glossary

Edit `glossary.json` to enforce consistent Vietnamese -> Japanese terminology.

Example:

```json
{
  "màn hình": "画面",
  "chức năng": "機能"
}
```

If `glossary.json` is missing, translation still runs without a glossary.

## Command-line usage

```powershell
.\.venv\Scripts\python.exe run_translate_all.py
```

Custom directories:

```powershell
.\.venv\Scripts\python.exe run_translate_all.py --input-dir "D:\VI" --output-dir "D:\JA"
```

Custom glossary:

```powershell
.\.venv\Scripts\python.exe run_translate_all.py --glossary "D:\glossary.json"
```


## Performance modes

The default is the 4B Q4 model because it is much lighter on RAM and faster on
machines with limited memory:

```text
translategemma:4b
```

Default requests use at most about 1200 source characters per chunk and retry
at most twice. The terminal prints elapsed time, output token count, model load
time, and generation speed when Ollama reports those metrics.

For the slower 12B quality mode:

```powershell
.\.venv\Scripts\python.exe run_translate_all.py --model translategemma:12b-it-q4_K_M
```

You can also adjust chunk size:

```powershell
.\.venv\Scripts\python.exe run_translate_all.py --max-chars 800
```


## Excel translation

Supported workbook formats:

```text
.xlsx
.xlsm
.xls
.xlsb
```

Excel translation requires Windows with Microsoft Excel installed. The tool
uses Excel COM through `pywin32`: it copies the original workbook to
`output` first and edits the copy in Excel itself instead of reconstructing
the workbook.

The Excel path is designed to preserve:

- cell formatting, number formats, fills, borders, alignment and merged cells
- formulas (formula cells are not translated)
- VBA in `.xlsm` files
- text formatting inside cells, including bold/italic/color/font/size
- text inside Excel shapes and text boxes, including per-run formatting
- grouped shapes recursively

For shape/text-box text, the tool checks the rendered text bounds after
translation. If the translated text overflows:

1. it reduces font sizes proportionally
2. it never shrinks a formatted run below 8 pt
3. if text still does not fit at 8 pt, it expands the object horizontally first
4. if vertical overflow remains after horizontal expansion, it increases height

The tool skips Excel temporary lock files such as `~$book.xlsx`.

Example:

```text
input/
├─ spec.md
├─ test_cases.xlsx
└─ design/
   └─ notes.xlsm

output/
├─ spec.md
├─ test_cases.xlsx
└─ design/
   └─ notes.xlsm
```

### Rich-text behavior

If one cell or shape contains multiple formatting runs, for example:

- normal text
- **bold text**
- colored text

the translator sends protected run boundaries through TranslateGemma and then
reapplies the original formatting to the corresponding translated run. If the
model damages those protected boundaries, the tool falls back to translating
each formatting run separately so formatting is not silently lost.
