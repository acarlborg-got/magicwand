# MagicWand Detailed Documentation

This document provides an in-depth overview of the MagicWand automation toolkit. It compiles information from the source code and existing documentation for developers, AI agents and regular users.

## 1. Introduction

MagicWand is a collection of VBA macros packaged as a Word template (`MagicWand.dotm`). The toolkit adds advanced processing features to Microsoft Word, focusing on bulk operations across multiple documents. The macros can be launched via the `ShowReplaceTool` entry point, which in turn displays the Index Browser and Search & Replace forms.

### Key Concepts
- **Language**: VBA (Visual Basic for Applications).
- **Host Application**: Microsoft Word.
- **Distribution**: The template can be shared (e.g., via OneDrive/SharePoint) and loaded as an add-in.
- **User Interaction**: A custom button in the Word ribbon or the macro list triggers `ShowReplaceTool`, opening the main interface.

## 2. Components

### 2.1 Forms
- **frmIndexBrowser** – Browse folders, build a list of documents and save the selection. Displays file and folder counts and lets the user choose which module to run next.
- **frmReplaceTool** – Runs search/replace operations, PDF export and spellchecking on the selected files. Shows progress, status and statistics.

### 2.2 Modules
- **modConfig** – Stores the application version (`v0.3 [Dev]`) and returns form titles.
- **modIndex** – Recursively indexes folders and Word files into `IndexedFolders` and `IndexedFiles` arrays.
- **modReplace** – Processes the selected documents: find/replace, optional PDF export, logging and error handling.
- **modSpellCheck** – Scans files for spelling errors, ranks them by frequency and suggests replacements.
- **modValidate** – Checks that user settings are safe (e.g., warns before overwriting files).
- **modUtils** – Updates the form status fields and progress bar.
- **modFileUtils** – File path helpers, folder creation and log file management.
- **modShared** – Shares the selected files/folders between forms.
- **modTypes** – Defines the `IndexedFile`, `IndexedFolder` and `tFileSelection` types.
- **modLauncher** – Contains `ShowReplaceTool`, the macro called when launching the add-in.
- **modMain** – Holds global flags such as `cancelRequested`.

See `REFERENCE.md` for a concise list of these modules and forms.

## 3. Installing the Add-in
1. Place `MagicWand.dotm` in a location that Word can access (OneDrive/SharePoint recommended).
2. In Word, go to **File → Options → Add-ins**.
3. Choose **Word Add-ins** and press **Add...**. Select the `MagicWand.dotm` file.
4. The next time Word starts, the macro `ShowReplaceTool` can be executed. You may create a custom ribbon button:
   - Go to **File → Options → Customize Ribbon**.
   - Create a new group under a tab of your choice.
   - Add a new macro button and select `ShowReplaceTool` as the command.
   - Optionally assign an icon and label (e.g., “MagicWand”).

5. Place additional language files in the `languages` folder (e.g., `EN.lng`, `SE.lng`) to customize interface text.
## 4. Using MagicWand
### 4.1 Index Browser
1. Launch the tool via the custom ribbon button or by running `ShowReplaceTool`.
2. The Index Browser lets you choose a root folder, index all subfolders and display the documents found.
3. Select which folders/files to process. Save the selection to make it available in other forms.

### 4.2 Search & Replace Form
1. After indexing, switch to the Search & Replace module.
2. Enter up to five find/replace pairs. Each row has **Case Sensitive** and **Whole Word** checkboxes.
3. Choose whether to export PDFs and whether to keep the original files in a dedicated subfolder. Prefixes and suffixes can be added to renamed files.
4. Press **Start** to process all selected files. Progress and status are shown live. Logs are written to `MagicWand_<user>_Log.txt` and `MagicWand_<user>_Errors.txt` in the base folder.
5. A spellcheck button analyzes the selected files, lists common mistakes and fills the replace fields automatically.

### 4.3 Cancel and Safety Checks
- The `Cancel` button sets a global flag that stops processing after the current file.
- `ValidateProcessingSettings` warns if settings might overwrite originals or if the “preserve original” folder is undefined.

## 5. Implementation Highlights
- **Recursive Indexing**: `modIndex` uses a stack-based traversal of folders to populate arrays of files and folders. Depth is recorded for visual indentation.
- **Batch Replacement**: `modReplace` opens each document, runs `ReplaceAll` over all story ranges and shapes, optionally exports to PDF, and logs the outcome.
- **Live Spellcheck**: `modSpellCheck` loops through documents, collects all spelling errors into a dictionary, and ranks them. The top entries populate the replace fields with suggestions via Word’s built-in spellchecking API.
- **Shared Data**: `modShared` exposes helper functions like `GetSelectedFilePaths` so forms can access the same selections.


## 6. Logging and Statistics
MagicWand writes two log files in the base folder of the documents being processed. To prevent
conflicts on shared storage, the file names include a sanitized user identifier obtained from
`GetUserIdentifier`:

```
MagicWand_<user>_Log.txt
MagicWand_<user>_Errors.txt
```

Each entry begins with a timestamp, making it easy to aggregate usage data and estimate time saved
across the organization.

## 7. Language Packs
All interface text can be localized through simple `.lng` files in the `languages` directory. The
default code is specified by `DEFAULT_LANGUAGE` in `modConfig`. Example files `EN.lng` and `SE.lng`
illustrate the `KEY=VALUE` format used by the `LoadLanguage` routine. The helper function `T()`
retrieves translations at runtime so additional languages can be added without recompiling the
template.

## 8. Roadmap and Future Plans
`ROADMAP.md` outlines planned features, including metadata editing, language detection and advanced
rules. Current development is tracked in `changelog.txt`.

## 9. Packaging
Run `./build_package.sh` to zip all modules, forms and documentation into
`package/MagicWand_v0.3-dev.zip` for distribution.

## 10. License
The project is intended for internal use at AFRY Buildings Automation Gothenburg.

---

MagicWand simplifies repetitive Word tasks by combining indexing, spellchecking, batch replacement and PDF export into a single, customizable add-in. By assigning `ShowReplaceTool` to a ribbon button, users can quickly launch the toolkit and process large document sets with minimal manual effort.

