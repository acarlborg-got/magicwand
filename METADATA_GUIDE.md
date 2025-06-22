# Metadata Module Guide

This guide describes how to add the metadata functions to your local Word project.

## Import modules
1. Open the VBA editor with **Alt+F11**.
2. Right‑click the project tree and choose **Import File...**
3. Select `modules/modMetadata.txt` and `forms/frmMetadata.txt` from this repository.
4. Ensure `modTypes.txt` is re‑imported so the new `FileMetadata` type is available.

## Form layout
Create a new UserForm named **frmMetadata** with the following controls:

| Control name | Type    | Purpose                |
|--------------|---------|------------------------|
| `txtTitle`   | TextBox | Document title         |
| `txtSubject` | TextBox | Document subject       |
| `txtAuthor`  | TextBox | Document author        |
| `txtKeywords`| TextBox | Document keywords      |
| `txtDate`    | TextBox | Document date (yyyy-mm-dd) |
| `btnInject`  | CommandButton | Write metadata to selected files |

Import the code from `forms/frmMetadata.txt` into this form.

## Using the tool
1. Index files with **frmIndexBrowser**.
2. Choose *Metadata injection* from the next-form dropdown.
3. Enter values in the metadata form and press **Inject**.
4. The fields are written to all selected documents.

The module also contains helpers to read metadata and extract extra
information such as `Datum`, `Handläggare` and `Konstruktör` from the
content of each Word file. The function `ExtractDocumentDateFromContent`
tries to locate the first `yyyy-mm-dd` string if no date is available in
the document properties.

## Advanced MetaTool

A second form **frmMetaTool** provides a larger interface for rule-based
metadata updates.

### Import
1. Import `modules/modMetaTool.txt` and `forms/frmMetaTool.txt`.
2. Ensure `modTypes.txt` is imported so the `MetaRule` type exists.

### Global section controls

| Control | Type | Purpose |
|---------|------|---------|
| `cmbMetaFieldGlobal` | ComboBox | Target metadata field |
| `cmbSourceGlobal` | ComboBox | Value source such as *Current User* |
| `txtValueGlobal` | TextBox | Used when source is *Static Text* |
| `chkOverwriteEmptyOnly` | CheckBox | Only set if current value is empty |
| `btnApplyGlobal` | CommandButton | Apply to all indexed files |

### Conditional matrix controls
The lower half uses `lstMetaMatrix` for rule rows along with the
`btnAddRule`, `btnRemoveRule` and `btnApplyMatrix` buttons. Each row lets
you pick a field, condition, action and replacement source.
