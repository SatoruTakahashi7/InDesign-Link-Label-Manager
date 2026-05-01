# InDesign Link Label Manager

AppleScript tool for assigning Finder labels to linked files used in Adobe InDesign documents.

Designed for practical DTP / graphic design workflows on macOS.

---

## Overview

This script scans all linked graphics in the currently opened InDesign document and applies Finder labels to the original linked files.

It is intended for:

- organizing Links folders
- checking unused assets
- identifying Illustrator files
- cleaning up project folders before delivery or archiving

---

## Features

- Apply Finder labels to linked source files
- .ai files are marked in orange
- Other image files are marked in red
- Remove labels from currently used linked files
- Process each original file only once
- Safely skips embedded or missing links
- Targets only files inside the document’s parent folder

---

## Supported Environment

- macOS
- Adobe InDesign 2022–2026

---

## Finder Label Rules

| File Type | Finder Label |
|---|---|
| Illustrator (.ai) | Orange |
| Other image files | Red |

---

## Typical Use Cases

- Links folder cleanup
- Unused asset checking
- Illustrator dependency inspection
- Pre-delivery project cleanup
- File organization before archiving

---

## How to Use

1. Open an InDesign document
2. Run the script
3. Choose:
   - Apply labels
   - Remove labels
4. Check Finder labels in macOS Finder

---

## Important Notes

- Always test on duplicated project data first
- Finder labels are applied per file
- Shared assets used by multiple documents may also be affected
- Links located outside the document’s parent folder are ignored

---

## Japanese File Name

The repository uses English file names for version control and GitHub management.

For local Japanese environments, you may freely rename the script file, for example:

text InDesignの配置に色付け.scpt 

---

## Author

GYAHTEI Design Laboratory  
@gyahtei_satoru

---

## License

MIT License
