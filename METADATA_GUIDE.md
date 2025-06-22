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
