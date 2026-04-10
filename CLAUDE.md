# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Is

**해외경비자동정산** (Overseas Expense Auto-Settlement) — a Windows desktop application that:
1. Uses Tesseract OCR (Simplified Chinese) to read Chinese receipt images
2. Classifies them into expense categories (교통비/개인경비/숙박비/식대/통신비) by Chinese keyword matching
3. Optionally generates a `.pptx` slide deck with the classified images
4. Extracts RMB amounts from receipts and writes them into a target Excel workbook (`해외출장비정산서_RMB.xlsx`)

## Running the Application

```bash
# Run v2 (customtkinter dark-theme UI)
python 해외경비자동정산_v2.0.0.py

# Run v1 (standard tkinter UI with optional background image)
python 해외경비자동정산_v1.2.6.py
```

**Required external dependency:** Tesseract-OCR must be installed at `C:\Program Files\Tesseract-OCR\tesseract.exe` with the `chi_sim` language pack.

## Installing Python Dependencies

```bash
pip install pillow pytesseract python-pptx openpyxl
# v2 only:
pip install customtkinter
```

## File Structure

Both scripts are self-contained single-file applications sharing the same architecture. v2.0.0 is the current main version.

```
해외경비자동정산_v1.2.6.py   # tkinter UI, background blur, PPT-based Excel writing
해외경비자동정산_v2.0.0.py   # customtkinter dark UI, direct image-to-Excel mode added
```

## Architecture

Each file is a flat single-module structure divided into three layers:

### 1. Constants & Configuration (top of file)
- `CATEGORIES` dict: Chinese keyword → expense category mappings
- `EXCEL_HEADER_KEYWORDS` dict: Korean/English header keyword → category mappings
- Regex patterns (`CURRENCY_AMOUNT_RE`, `CONTEXT_AMOUNT_RE`, etc.) for multi-priority amount extraction
- `RUN_MODE_OPTIONS`: maps mode key → `(label, do_classify, do_ppt)` tuples

### 2. Pure Processing Functions (middle)
- **OCR & categorization**: `process_images()` — OCR images with Tesseract `chi_sim`, match against `CATEGORIES`, move files to category subfolders
- **Amount extraction pipeline** (multi-priority, highest wins):
  - Score 4: negative-sign amounts (`-255.00`)
  - Score 3: currency-prefixed amounts (`¥255.00`, `RMB 255`)
  - Score 2: context keyword amounts (`합계 255`, `合计 255`)
  - Score 1: generic numeric fallback
- **PPT generation**: `create_ppt_from_subfolders()` — 16 images/slide, 8 columns × 2 rows grid layout
- **Excel writing**:
  - `write_amounts_to_excel()` — writes category amounts from `process_images()` results into "sum" sheet
  - `write_ppt_amounts_to_excel()` — reads amounts from PPT slide images via OCR, writes to Excel (v1 only)
  - `write_images_to_excel_direct()` — reads amounts directly from subfolder images, writes to Excel (v2 only)
- **Excel header detection**: `detect_sum_header_row_and_columns()` scans up to row 80 / col 80 for header keywords

### 3. `ExpenseAutoApp` GUI Class (bottom)
- Builds UI with `_build_ui()` → `_build_header()`, `_build_function_row()`, `_build_bottom_row()`
- Processing runs on a background `threading.Thread`; results flow back via `queue.Queue`
- `_drain_queue()` is polled every 100ms via `root.after()` to update UI from queue messages
- Queue message kinds: `"log"`, `"progress"`, `"preview"`, `"done"`, `"fatal"`
- `desc.txt` in the app directory is loaded as the help/description text (F1)

## Key Differences Between v1 and v2

| Feature | v1.2.6 | v2.0.0 |
|---|---|---|
| GUI framework | `tkinter` + `ttk` | `customtkinter` (dark theme) |
| Background image | Blurred background image | Not present |
| Excel writing mode | Via PPT slide OCR | Direct image-to-Excel (`write_images_to_excel_direct`) |
| Window size | 1140×700 | 1240×820 |

## Processing Flow (Mode 3 — full pipeline)

1. User selects input folder containing `.jpg`/`.jpeg`/`.png` receipt images
2. `process_images()`: OCR each image → categorize → move into category subfolder, extract amount per image
3. `create_ppt_from_subfolders()`: generate `.pptx` from the newly organized subfolders
4. Excel writing: amounts written to the "sum" sheet of `해외출장비정산서_RMB.xlsx` found in the same folder
5. Summary shown in log panel; unmatched images go to `미분류/` subfolder

## Important Path Assumptions

- Tesseract: `C:\Program Files\Tesseract-OCR\tesseract.exe`
- Default input folder: `C:\VAD_PC\경비\해외_출장경비`
- App icon/background assets looked up in app directory and `pics/` subdirectory
- `desc.txt` (help text) must be in the same directory as the script/executable
