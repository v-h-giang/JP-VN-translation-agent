# TranslateGemma Markdown Batch Translator

Local Vietnamese -> Japanese Markdown translation with
`translategemma:4b` through Ollama by default.

## Structure

```text
JP-VN-translation-agent/
├─ input/
├─ output/
├─ run_translate_all.py
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
