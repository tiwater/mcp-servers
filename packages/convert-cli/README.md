# tiwater-convert

A generic CLI for loss-aware office format conversion.

## Initial scope

- `.xls` -> `.xlsx`
- Writer documents -> `.pdf` through WPS Writer RPC when available
- Other Office document/workbook/presentation formats -> `.pdf` through a local LibreOffice/soffice install

## Usage

```bash
tiwater-convert xls-to-xlsx <input.xls> <output.xlsx>
tiwater-convert docx-to-pdf <input.docx> <output.pdf>
tiwater-convert xlsx-to-pdf <input.xlsx> <output.pdf>
tiwater-convert pptx-to-pdf <input.pptx> <output.pdf>
```

XLS conversion prefers WPS Spreadsheets through `pywpsrpc` when available,
because WPS generally preserves legacy workbook formatting better on Linux. Set
`TIWATER_WPSRPC_PYTHON` or `LUCID_WPSRPC_PYTHON` to the pywpsrpc venv Python if
it is not installed at `~/.local/share/lucid-docs/wpsrpc-venv/bin/python`.
`xvfb-run` and the WPS `et` command must be available. If WPS conversion is not
available or fails, the converter falls back to LibreOffice/soffice and then to
the built-in NPOI converter.

DOC, DOCX, ODT, and RTF PDF conversion prefers WPS Writer through `pywpsrpc`
when `wps`, `xvfb-run`, and the configured WPS RPC Python are available. The
JSON result records `backend: "wps-writer"` or `backend: "libreoffice"` and a
fallback reason. Set `TIWATER_OFFICE_PDF_BACKEND=wps-writer` to require the
native WPS backend and fail closed when it is unavailable; use `libreoffice` to
request the auxiliary backend explicitly.

LibreOffice-backed PDF conversion requires `soffice`. If it is not on `PATH`, set one of:

- `TIWATER_SOFFICE`
- `SOFFICE`
- `LIBREOFFICE_PATH`
