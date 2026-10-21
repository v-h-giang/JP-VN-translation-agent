# TranslateGemma Markdown Batch Translator

## Cấu trúc

```text
translategemma_batch_ready/
├─ input/
│  └─ test_vi.md
├─ output/
│  └─ test_jp.md
├─ run_translate_all.py
├─ glossary.json
├─ setup.bat
└─ run.bat
```

## Lần đầu

1. Double-click `setup.bat`.
2. Script sẽ tự cài Ollama nếu máy chưa có và tự tải model
   `translategemma:12b-it-q4_K_M` nếu model chưa được tải.

`pip install ollama` chỉ cài Python client. `setup.bat` cũng bảo đảm Ollama
runtime/server và model đều sẵn sàng. Model tải xuống khoảng 8,1 GB, vì vậy máy
cần kết nối Internet và đủ dung lượng trống trong lần setup đầu tiên.

## Mỗi lần dịch

1. Copy toàn bộ file `.md` / `.markdown` cần dịch vào `input`.
2. Double-click `run.bat`.
3. Kết quả xuất vào `output`.

## Thư mục con

Nếu input:

```text
input/
├─ a.md
└─ design/
   ├─ b.md
   └─ api/
      └─ c.md
```

thì output:

```text
output/
├─ a.md
└─ design/
   ├─ b.md
   └─ api/
      └─ c.md
```

## Glossary
Sửa `glossary.json` để ép thuật ngữ Việt -> Nhật thống nhất.
