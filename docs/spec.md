# Certificate Generator Specification

## Version

- Spec version: `0.5`
- Updated: `2026-03-09`

## Purpose

The workbook generates Word-based certificates for military personnel from Excel source data, stores a generation history, and provides an operator-focused workflow through Excel and a custom ribbon tab.

## Current Scope

- Import source data from an external Excel workbook into `ImportedData`.
- Search personnel records from `ImportedData` through `frmSearch`.
- Copy the selected record into the `data` sheet.
- Generate Word documents from templates stored in a configurable template folder.
- Save generated certificates into a configurable output folder without overwriting existing files.
- Save a snapshot workbook in the workbook directory after generation.
- Append generation records into `IssuedDocumentsLog`.
- Expose primary actions through the `Certificates` ribbon tab.

## Implemented Features

### Data handling

- Search results are bound to source row numbers instead of parsing display strings.
- Unit values preserve text-based designations instead of stripping them down to digits.
- Certificate recipient names are declined through `UDFs_FIO.FIO(..., "D")` with fallback to `DativeCase`.
- Unit replacement values support declined forms such as `Войсковая часть 12345 -> войсковой части 12345`.

### Generation

- Output folder is persisted as a workbook-level text setting: `CERTIFICATE_OUTPUT_FOLDER`.
- Template folder is selected through ribbon and persisted as a workbook-level text setting: `FILE_WORD`.
- Template list can be migrated from a legacy named range into a workbook-level text setting: `FILE_TEMPLATE`.
- Existing output files are preserved; new files receive unique names when needed.
- Missing placeholder warnings are suppressed in the completion message because partial placeholder sets are expected.

### UI

- Ribbon tab `Certificates` is defined in `customUI14.xml`.
- Ribbon callbacks are implemented in `RibbonCallbacks.bas`.
- Main ribbon actions:
  - `Generate`
  - `Search Person`
  - `Import Source Data`
  - `Open History`
  - `Template Folder`
  - `Output Folder`
  - `About`

### Maintainability

- VBA text modules include version and update annotations in the header.
- Exported VBA sources are stored in `CreateCertificateForMilitaryPersonnel.xlsb.modules`.
- The temporary `vba-import-ready` export flow was removed in favor of `VbaModuleManager`.

## Current Configuration Model

### Workbook text settings

- `FILE_WORD`: template folder path
- `FILE_TEMPLATE`: semicolon-delimited template list
- `CERTIFICATE_OUTPUT_FOLDER`: output folder path

### Worksheets

- `data`: main operator worksheet
- `ImportedData`: imported personnel data
- `IssuedDocumentsLog`: generation history
- `ЧтоНового`: informational worksheet

The former `const` / `Settings` worksheet is being retired. Configuration is moving to workbook-level text settings so the workbook no longer depends on a visible settings sheet.

## Known Next Step

The next planned enhancement is a ribbon-driven template selection UI:

- Add a ribbon command for template selection.
- Show available `.docx` files from the configured template folder.
- Save selected templates back into `FILE_TEMPLATE`.
- Remove the need to edit `FILE_TEMPLATE` through Name Manager.
