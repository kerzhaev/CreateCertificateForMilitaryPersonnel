# Certificate Generator Specification

## Version

- Spec version: `0.7.5`
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

- Source import is intentionally fast: the first worksheet from the selected workbook is copied into `ImportedData` without full normalization during import.
- Personnel search now performs smart header-based column detection on top of the copied `ImportedData` sheet.
- The search supports both a single `Full Name` column and split `Surname` / `Name` / `Patronymic` columns.
- The search detects separate text and numeric `Лицо` columns to distinguish full name from table number.
- Search results are bound to source row numbers instead of parsing display strings.
- Unit values preserve text-based designations instead of stripping them down to digits.
- Certificate recipient names are declined through `UDFs_FIO.FIO(..., "D")` with fallback to `DativeCase`.
- Unit replacement values support declined forms such as `Войсковая часть 12345 -> войсковой части 12345`.

### Generation

- Output folder is persisted as a workbook-level text setting: `CERTIFICATE_OUTPUT_FOLDER`.
- Template folder is selected through ribbon and persisted as a workbook-level text setting: `FILE_WORD`.
- Template list can be migrated from a legacy named range into a workbook-level text setting: `FILE_TEMPLATE`.
- Existing output files are preserved; new files receive unique names when needed.
- Snapshot workbooks are saved in the workbook directory using only the current date as the file name base.
- Missing placeholder warnings are suppressed in the completion message because partial placeholder sets are expected.
- `IssuedDocumentsLog` rows are formatted on write with `Times New Roman`, `12 pt`, thin borders, a readable datetime format for `Created On`, bold wrapped headers, auto-filter, and fixed widths for key columns.

### UI

- Ribbon tab `Certificates` is defined in `customUI14.xml`.
- Ribbon callbacks are implemented in `RibbonCallbacks.bas`.
- Main ribbon actions:
  - `Generate`
  - `Search Person`
  - `Import Source Data`
  - `Open History`
  - `Template Folder`
  - `Select Templates`
  - `Output Folder`
  - `About`

### Template configuration

- The template folder path is selected from the ribbon and stored in `FILE_WORD`.
- The template catalog is stored in `FILE_TEMPLATE` as a semicolon-delimited list.
- `UserForm1` now acts as a template manager:
  - it reads available `.docx` files from the configured template folder;
  - it preselects templates already stored in `FILE_TEMPLATE`;
  - it saves the selected catalog back into `FILE_TEMPLATE`.
- Generation no longer depends on `FILE_TEMPLATE` being populated because actual row processing uses column 3 from the `data` sheet.

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
- `ImportedData`: raw imported personnel data from the selected source workbook
- `IssuedDocumentsLog`: generation history
- `ЧтоНового`: informational worksheet

The former `const` / `Settings` worksheet is being retired. Configuration is moving to workbook-level text settings so the workbook no longer depends on a visible settings sheet.

## Current Limitation

- The workbook source files now support removing the visible `Settings` sheet, but the live workbook still needs to be synced through `VbaModuleManager` before the sheet can be safely deleted.
- Smart search depends on recognizable source headers for full name, table number, personal number, birth date, and unit fields.
