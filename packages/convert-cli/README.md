# tiwater-convert

A generic CLI for loss-aware office format conversion.

## Initial scope

- `.xls` -> `.xlsx`
- Writer documents -> `.pdf` through WPS Writer RPC when available
- XLS/XLSX workbooks -> `.pdf` through WPS Spreadsheets RPC when available
- PPT/PPTX/ODP presentations -> `.pdf` through WPS Presentation RPC when available
- Other Office formats -> `.pdf` through a local LibreOffice/soffice install

## Usage

```bash
tiwater-convert xls-to-xlsx <input.xls> <output.xlsx>
tiwater-convert recalculate-xlsx <input.xlsx> <output.xlsx>
tiwater-convert docx-to-pdf <input.docx> <output.pdf>
tiwater-convert xlsx-to-pdf <input.xlsx> <output.pdf>
tiwater-convert pptx-to-pdf <input.pptx> <output.pdf>
```

Successful native WPS PDF conversions include
`native_render_provenance` conforming to
`tiwater.convert-native-render-provenance/v1`. The provider records the WPS
package build and executable hash, OS/.NET runtime identity, a count/hash-only
fontconfig inventory based on font family, style, and font bytes, input/output
SHA-256 identities, and PDF page count. Native conversion fails closed when
any required provenance cannot be collected. Consumers should retain this
provider-owned object unchanged rather than probing the host independently.

XLS conversion prefers WPS Spreadsheets through `pywpsrpc` when available,
because WPS generally preserves legacy workbook formatting better on Linux. Set
`TIWATER_WPSRPC_PYTHON` or `LUCID_WPSRPC_PYTHON` to the pywpsrpc venv Python if
it is not installed at `~/.local/share/lucid-docs/wpsrpc-venv/bin/python`.
`xvfb-run` and the WPS `et` command must be available. If WPS conversion is not
available or fails, the converter falls back to LibreOffice/soffice and then to
the built-in NPOI converter.

DOC, DOCX, ODT, and RTF PDF conversion prefers WPS Writer through `pywpsrpc`
when `wps`, `xvfb-run`, `dbus-run-session`, and the configured WPS RPC Python
are available. The converter launches WPS with an isolated D-Bus session and
writable per-conversion XDG cache/runtime directories so it remains usable from
restricted automation sandboxes. The
JSON result records `backend: "wps-writer"` or `backend: "libreoffice"` and a
fallback reason. Set `TIWATER_OFFICE_PDF_BACKEND=wps-writer` to require the
native WPS backend and fail closed when it is unavailable; use `libreoffice` to
request the auxiliary backend explicitly.

XLS and XLSX PDF conversion likewise prefers WPS Spreadsheets through
`pywpsrpc` when `et`, `xvfb-run`, `dbus-run-session`, and the configured RPC
Python are available. Set `TIWATER_OFFICE_PDF_BACKEND=wps-spreadsheet` to
require that native backend and fail closed; it never treats LibreOffice output
as WPS Spreadsheet proof.

PPT, PPTX, and ODP PDF conversion prefers WPS Presentation through `pywpsrpc`
when `wpp`, `xvfb-run`, `dbus-run-session`, and the configured RPC Python are
available. Set `TIWATER_OFFICE_PDF_BACKEND=wps-presentation` to require that
native backend and fail closed; it never treats LibreOffice output as WPS
Presentation proof.

On macOS, set `TIWATER_WPS_WRITER_LIMA_INSTANCE` to the name of a configured
Lima Linux instance that exposes the documented `/tmp/lucid-wps-render` shared
directory and has the published WPS Writer, Spreadsheets, and Presentation runtimes installed. The converter
then stages each input and output in a unique shared directory and invokes
`limactl shell <instance>` itself; it reports the requested `wps-writer` or
`wps-spreadsheet`, or `wps-presentation` backend and fails closed if the configured instance cannot render. This is an explicit
runtime configuration, not an arbitrary command hook.

LibreOffice-backed PDF conversion requires `soffice`. If it is not on `PATH`, set one of:

- `TIWATER_SOFFICE`
- `SOFFICE`
- `LIBREOFFICE_PATH`
