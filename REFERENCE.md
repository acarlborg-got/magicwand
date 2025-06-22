# MagicWand File Reference

This reference summarizes the purpose of each exported form and module.

## Forms
- **frmIndexBrowser** – User interface for indexing folders and selecting files to process.
- **frmReplaceTool** – Main form that runs search/replace, PDF export, and spellchecking on the selected files.
- **frmMetadata** – Form for viewing and injecting document metadata.

## Modules
- **modConfig** – Stores the application version and helper functions for form titles.
- **modIndex** – Recursively indexes files and folders into global arrays.
- **modReplace** – Executes batch operations such as Find/Replace and PDF export.
- **modSpellCheck** – Automates preflight spellchecking across chosen documents.
- **modValidate** – Ensures user settings are safe before writing to files.
- **modUtils** – Status bar and progress bar helpers for long operations.
- **modFileUtils** – File path helpers and log file creation.
- **modShared** – Shared data structures like selected files and folders.
- **modTypes** – Type definitions for indexed items and selections.
- **modMetadata** – Helpers for reading and writing Word metadata and extracting document dates.
- **modLauncher** – Subroutine that starts the main form.
- **modMain** – Global flags for cancel operations and overall control flow.

