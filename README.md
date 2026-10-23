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


## Excel and LibreOffice translation

Supported formats: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`.

The agent **automatically probes Microsoft Excel desktop**. If Excel COM is
available it uses that backend; if unavailable or opening fails it falls back
to the **LibreOffice UNO** backend. LibreOffice is sufficient for `.xlsx`
and `.xls`, and Microsoft Excel is not required. Microsoft Excel is still
recommended for files with strict layout fidelity, advanced drawings or macros.

On Windows, install LibreOffice normally (e.g.
`C:\Program Files\LibreOffice`); the agent detects `soffice.exe`. If
installed elsewhere, set `LIBREOFFICE_HOME` to the installation directory.
The UNO worker loads native Python bindings / DLLs from the installed
LibreOffice `program` directory and uses a dedicated headless process.
**Do not copy arbitrary LibreOffice DLLs into this repository**: they depend
on each other and must match the LibreOffice version and Python ABI. The
preferred dynamic dependency source is the installed LibreOffice runtime.

`setup.bat` installs the Ollama Python client and optionally pywin32 on
Windows; LibreOffice does not need to be installed with pip. For UNO the
agent probes the bundled LibreOffice Python or a compatible local Python
interpreter which can `import uno`.

### Workbook protection and fidelity

- Source workbooks are never modified.
- The output is written only after processing a file.
- Formula cells and numeric cells remain unchanged.
- Text cells, text boxes and standard drawing/group shapes are translated.
- Where UNO exposes formatted text portions, the agent attempts to preserve
  bold, color, italic and font size per portion.
- If translated text exceeds shape size, the agent estimates layout, tries
  shrinking to **at least 8pt**, then expands horizontally and then vertically.
- The UNO sizing check is approximate, not rendered overflow measurement.

**Limitations:** LibreOffice may re-serialize `.xlsx` formatting and drawing
objects, especially charts, threaded comments, SmartArt, embedded OLE,
conditional formatting and complex shapes. Excel native formatting fidelity
cannot be guaranteed by a LibreOffice round trip. For safety, `.xlsm`
(macro-enabled) and `.xlsb` **require Microsoft Excel**; the LibreOffice
fallback refuses them rather than risking macro loss or binary corruption.
Make a backup and visually review the output before delivering a critical
client workbook.

Run normally with `run.bat`; it automatically selects the backend and
preserves the relative input/output directory structure.


## Sparse scanning, context batching and checkpoints (LibreOffice)

For LibreOffice UNO, the agent queries nonempty text cells with Calc's
`queryContentCells(CellFlags.STRING)` instead of iterating a rectangular
`UsedRange`. The previous 150,000-position rectangle cap has been removed.
Formula and number cells are never sent to TranslateGemma.

- Cell and shape text is deduplicated via a translation cache: identical
  strings are translated once per workbook and reused for every occurrence.
- Different text strings are grouped into contextual batches of at most
  **16 strings / about 1100 characters**. Each is tagged with an ID; the
  model must return the matching IDs. Invalid responses fall back to
  separate strings, so text is not assigned to the wrong cell.
- **Every 75 completed cells/objects** (default), UNO saves a checkpoint
  workbook plus a JSON manifest including completed object IDs, cache and
  the original input file's SHA-256.
- On restart using the same input/output file names, checkpoints are
  validated and processing resumes at the next unfinished object.
- A checkpoint never overwrites the final output file; the final file is
  published only when the whole workbook finishes successfully.
- After successful completion, temporary checkpoint artifacts are deleted.

Intermediate files in the `output` folder are named
`<name>.checkpoint.xlsx` and `<name>.checkpoint.json` (or `.xls`).
Do not edit those files while translation is running or paused.

Rich-text portions remain separate so formatting can be applied per
translated run. The UNO shape overflow size calculation remains approximate.

**Note:** Excel COM is still preferred when installed, but the sparse scan,
context batching and resume functionality currently apply to the
LibreOffice UNO backend. For consistent behavior on a LibreOffice-only
machine, set `SPREADSHEET_BACKEND=libreoffice`.
