# MagicWand Toolkit  
**Version:** v0.25b [Stable]  
**Module:** Search & Replace  
**Date:** 2025-05-21  

## Overview  
MagicWand is a modular VBA-based automation toolkit designed for advanced document management in Microsoft Word.  
The current release focuses on high-volume search, replace, export, and spelling validation across entire folder structures.  

It is built to support professional workflows in environments with heavy documentation demands — such as building automation, project documentation, commissioning protocols, operation cards, and technical descriptions.  

## Features  

### 🔍 Search & Replace  
- Up to 5 simultaneous Find/Replace pairs  
- Case Sensitive and Whole Word matching per row  
- Optional prefix and suffix for renamed output files  
- Preserve original files in dedicated subfolder  
- Export results to PDF or PDF/A-1b  
- Log replacements and errors to local `.txt` files  
- Automatic exclusion of preserve folders to avoid recursion  
- Supports recursive or non-recursive folder processing  
- Real-time status updates, progress bar, and completion summary  

### ✅ Preflight Spellcheck  
- Scans all Word documents for language-specific spelling errors  
- Live display of detected words, sorted by frequency  
- Top 5 errors auto-filled into Find/Replace fields  
- Suggests replacements via Word's spellcheck engine  
- Evaluates Case/WholeWord options per entry  
- Full error dictionary exported to `MagicWand_Spelling.txt`  

## Usage  
1. Place `MagicWand.dotm` in a shared folder synced via OneDrive (SharePoint recommended)  
2. Load the .dotm file into Word via Add-ins:  
   `File → Options → Add-ins → Word Add-ins → Add...`  
3. The form launches automatically or via the macro: `ShowReplaceTool`  
4. Choose folder, language, and options  
5. Run search/replace or spellcheck across multiple documents  

## Suggested Use Cases  

### 🔁 Bulk Replace  
- Update all project documents with a new client name, site reference, or version code  
- Change all occurrences of a legacy product code across hundreds of files  
- Replace old dates, terms, or contact info regardless of whether they are in headers, footers, or body  

### 📁 Batch Rename & Export  
- Append date/version to filenames (e.g., `_v1.2_2025-05`)  
- Export all `.docx` files as PDF/A-1b for archive or client delivery  
- Rename files using client-specific or system-specific prefixes/suffixes  

### 📝 Standardize Documents  
- Automatically replace terms or references across multiple documents  
- Batch-apply naming conventions and versioning 

### 📚 Preflight Spellcheck  
- Detect and resolve frequent misspellings across a project  
- Catch subtle typing errors not flagged manually  
- Export all detected spelling issues to log file for audit or documentation  

## Upcoming Modules
- Metadata injection and extraction (Author, Title, Keywords, Document date)
- Document structure mapping and content validation  
- Rule-based text filtering per document type (e.g., Technical Description vs Operation Card)  
- Integrated language pack support via `.lng` files  
- Modular GUI: Form switching instead of tab navigation  
- Remove highlights and change tracking elements  
- Formatting cleanup: bold, underlines, strikeout, and styles  

## Version & Change Tracking
Current version and changelog are shown in the form interface.
Development for version 0.3 happens on the `dev` branch while `main` holds the
stable v0.25 release. Older releases are tagged for easy reference.
All release notes are stored in `changelog.txt`.

## Additional Documentation
- [Roadmap](ROADMAP.md)
- [File Reference](REFERENCE.md)
- [Metadata Guide](METADATA_GUIDE.md)

## License  
Internal use only – AFRY Buildings Automation Gothenburg 
Authored and maintained by internal development team  
