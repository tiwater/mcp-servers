# tiwater-convert

A generic CLI for loss-aware office format conversion.

## Initial scope

- `.xls` -> `.xlsx`
- Office document/workbook/presentation formats -> `.pdf` through a local LibreOffice/soffice install

## Usage

```bash
tiwater-convert xls-to-xlsx <input.xls> <output.xlsx>
tiwater-convert docx-to-pdf <input.docx> <output.pdf>
tiwater-convert xlsx-to-pdf <input.xlsx> <output.pdf>
tiwater-convert pptx-to-pdf <input.pptx> <output.pdf>
```

PDF conversion requires LibreOffice. If `soffice` is not on `PATH`, set one of:

- `TIWATER_SOFFICE`
- `SOFFICE`
- `LIBREOFFICE_PATH`
