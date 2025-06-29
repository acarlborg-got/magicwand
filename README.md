# MagicWand Toolkit

MagicWand is a VBA based automation suite for managing large collections of Microsoft Word documents. It bundles search & replace utilities, spell checking and metadata management into a single add‑in.

## Table of Contents
- [Getting Started](#getting-started)
- [Features](#features)
- [Suggested Use Cases](#suggested-use-cases)
- [Upcoming Modules](#upcoming-modules)
- [Documentation](#documentation)
- [License](#license)

## Getting Started
1. Place `MagicWand.dotm` in a shared location (OneDrive/SharePoint recommended).
2. Load the add-in in Word: `File → Options → Add-ins → Word Add-ins → Add...`.
3. The form launches automatically or by running the macro `ShowReplaceTool`.
4. Choose folder, language and options, then run search/replace or spellcheck across multiple documents.

## Features
- Up to 5 simultaneous Find/Replace pairs with case and whole word control.
- Optional prefix/suffix for renamed output files and dedicated preserve folder.
- Export results to PDF or PDF/A‑b with logging of replacements and errors.
- Recursive or non-recursive folder processing with real-time progress display.
- Automated spelling check that exports an error dictionary.

## Suggested Use Cases
- Bulk replace client names or version codes across many documents.
- Batch rename and export documents to PDF/A for archive or delivery.
- Standardise terminology and naming conventions.
- Detect frequent misspellings before final delivery.

## Upcoming Modules
- Metadata injection and extraction.
- Document structure mapping and validation.
- Rule-based text filtering by document type.
- Integrated language pack support.

## Documentation
Documentation files are located in the [`docs`](docs/) directory:
- [FAQ](docs/FAQ.md)
- [Roadmap](docs/ROADMAP.md)
- [Reference](docs/REFERENCE.md)
- [Metadata Guide](docs/METADATA_GUIDE.md)
- [Changelog](docs/CHANGELOG.md)

Source code for forms and modules can be found under [`src/forms`](src/forms/) and [`src/modules`](src/modules/). The legacy plain text changelog remains in [`docs/changelog.txt`](docs/changelog.txt).

## License
Internal use only – AFRY Buildings Automation Gothenburg.
