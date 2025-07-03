# MagicWand – Help & FAQ

This document provides practical guidance on how to use MagicWand’s features effectively within Microsoft Word.  
It complements the technical `README.md` by focusing on workflows, combinations, and user operations.

## 🔍 Search & Replace Workflows

### Basic Find and Replace

1. Enter the text to find in `txtFind1`, and the replacement in `txtReplace1`.
2. Optionally:
   - Check `chkCase1` to match case.
   - Check `chkWhole1` to match whole words only.
3. Repeat for `txtFind2–5` if multiple terms should be replaced.
4. Click **Start** to run the operation across selected documents.

### Replace + Export as PDF

- Enable `chkExportPDF` to export each processed document to PDF after replacements.
- Select format via `cmbPDFType`:
  - `Normal` or `PDF/A-1b` (archive-safe format).
- Use `txtAltPDFPath` to define a separate folder for exported PDFs.
  - If left empty, PDFs are saved in the same folder as the originals.

### File Renaming and Suffixes

- Use `txtPrefix` to prepend text to the new filename (e.g., `ClientA_`).
- Use `txtSuffix` to append versioning or date codes (e.g., `_v2_2025-07`).
- Original filenames are preserved unless `chkKeepOriginal` is unchecked.

### Preserve Original Files

- Enable `chkKeepOriginal` to save the original `.docx` in a subfolder.
- The subfolder name is taken from `txtPreserveSubFolder`.
  - If left blank, a fallback folder is created.
- The processed file replaces the original only if this is **unchecked**.

## 📚 Spellcheck & Date Scan

### Important Note on Find Dates & Spellcheck

The **Find Dates** and **Spellcheck** functions are standalone scanners.  
They **cannot be combined directly** with each other or with the Replace operation.

If you want to:
- **Replace found dates or spelling errors**, you must:
  1. First run the desired scanner (Spellcheck or Find Dates)
  2. Manually review and adjust the `txtFindX / txtReplaceX` fields
  3. Then click **Start** to apply replacements across files

### Preflight Spellcheck

- Set the language using `cmbLanguage` (Svenska / Engelska).
- Click **Spellcheck** to scan all documents in the selected folder.
- Results:
  - `lstSpellingResult` shows detected terms with their frequency.
  - `txtFind1–5` auto-filled with top 5 errors.
  - `txtReplace1–5` suggests corrections.
- A log file is created: `MagicWand_Spelling.txt`

### Find Dates (YYYY-MM-DD)

- Click **Find Dates (YYYY-MM-DD)** to scan for valid and partial ISO dates.
- Supports:
  - `2024-12-31`, `2024-11-xx`, `xxxx-xx-xx`, `2025-07-Xx`
- Matches populate `lstSpellingResult` with count (e.g., `2024-05-xx (4)`).
- Top 5 patterns are filled into `txtFind1–5`, with today’s date suggested in `txtReplace1–5`.

## 🛠 Settings & Tips

- Use `chkIncludeSubfolders` to include all nested folders in the scan.
- You can interrupt any operation with **Cancel** (sets `cancelRequested = True`).
- Status is shown in `lblStatus2`, and a progress bar tracks completion live.

## 📦 Log Files

- Two log files are generated in the selected folder:
  - `MagicWand_Log.txt` – lists processed documents.
  - `MagicWand_Errors.txt` – lists any files that failed to process.

## 📊 Statistics & Time Saved

After each operation, MagicWand logs the results to a `.csv` file under `/logs/`.

Two real-time statistics summaries are displayed in the form:

- **Global Statistics:** Cumulative results across all users
- **My Statistics:** Filtered summary based on your username

The stats include:

- Files processed
- Replacements made
- PDFs exported
- Estimated time saved based on a hardcoded manual effort model

---

## 💡 Example: Full Workflow

> Replace 3 terms in all documents, add suffix `_v3`, export to PDF, and preserve originals.

1. Set `txtFind1–3` and `txtReplace1–3`
2. Check `chkExportPDF` and set PDF type
3. Set `txtSuffix = "_v3"`
4. Enable `chkKeepOriginal` and define a subfolder name
5. Click **Start**


