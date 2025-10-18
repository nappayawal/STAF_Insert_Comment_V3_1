# STAF Insert Comment Tool V3.1 (xlwings)

**Goal:** Insert formatted Excel comments (legacy Notes) into `STAF.xlsm` **without** deleting floor plan shapes/graphics.

## Why xlwings?
- `openpyxl` is great for reading logic, but saving `.xlsm` can strip shapes.
- `xlwings` uses the **Excel COM API** so Excel itself writes comments → shapes stay intact.

## Project Structure
```
STAF_Insert_Comment_V3_1/
├── main.py
├── gui.py
├── excel_tools/
│   ├── __init__.py
│   ├── staf_logic.py
│   └── xlwings_comment.py
├── assets/
├── test_files/
│   ├── STAF_sample.xlsm
│   └── Machine_Details.xls
└── README.md
```

## Setup
```bash
python -m venv .venv
# Windows cmd
.venv\Scripts\activate
pip install openpyxl xlwings tk
```

## Run
```bash
python main.py
```
- Use **Insert TEST comment (xlwings)** to validate shape preservation.
- Use **Run FULL logic** to validate reading + metric detection (no write with openpyxl).

## Packaging (optional)
```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --icon assets/icon.ico main.py
```

> ⚠️ Packaging xlwings apps may require including Excel being installed on the target machine.

## Notes
- We insert **legacy Notes** (`Range.AddComment`). In the UI these appear as *Notes* (modern Excel renamed old comments).
- The tool **avoids duplicates** by comparing existing note text before inserting/updating.
- Comments are **autosized** by default; sizing can be tuned in `excel_tools/xlwings_comment.py`.
- We **never save** `.xlsm` with openpyxl—only read—to avoid losing shapes.