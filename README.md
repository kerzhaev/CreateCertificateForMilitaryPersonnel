# CreateCertificateForMilitaryPersonnel

Excel/VBA workbook for generating Word-based certificates for military personnel from imported Excel source data.

## Repository Structure

- `CreateCertificateForMilitaryPersonnel.xlsb`
  Main workbook file.
- `CreateCertificateForMilitaryPersonnel.xlsb.modules/`
  Exported VBA source files managed through `VbaModuleManager`.
- `customUI14.xml`
  Custom Ribbon XML definition.
- `docs/`
  Project documentation.
- `шаблоны/`
  Word templates used by the workbook.
- `ГотовыеСправки/`
  Runtime output folder. Ignored by git.

## Working Rules

- Edit VBA sources in `CreateCertificateForMilitaryPersonnel.xlsb.modules/`.
- Sync them back into the workbook through `VbaModuleManager`.
- Do not store temporary Excel lock files or generated output in git.
- Keep workbook-specific runtime artifacts outside source control.

## Main Documentation

- `docs/spec.md` - technical specification
- `docs/user-manual.ru.md` - end-user manual in Russian
