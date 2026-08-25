# tiwater-convert

A generic CLI for loss-aware office format conversion.

## Initial scope

- `.xls` -> `.xlsx`
- Writer documents -> `.pdf` through WPS RPC when available
- XLS/XLSX workbooks -> `.pdf` through ET RPC when available
- PPT/PPTX/ODP presentations -> `.pdf` through WPP RPC when available
- Other Office formats -> `.pdf` through a local LibreOffice/soffice install

## Usage

```bash
tiwater-convert xls-to-xlsx <input.xls> <output.xlsx>
tiwater-convert recalculate-xlsx <input.xlsx> <output.xlsx>
tiwater-convert refresh-docx-fields <input.docx> <output.docx>
tiwater-convert docx-to-pdf <input.docx> <output.pdf>
tiwater-convert xlsx-to-pdf <input.xlsx> <output.pdf>
tiwater-convert pptx-to-pdf <input.pptx> <output.pdf>
```

`refresh-docx-fields` opens the current DOCX in WPS Writer, refreshes every
table of contents and table of figures, and repaginates the document. The
distinct output imports only those refreshed index results and their referenced
internal bookmarks into the input package; unrelated body content and package
parts remain input-authoritative. Its JSON receipt conforms to
`tiwater.convert-refresh-docx-fields/v1` and binds the input and output bytes by
SHA-256. It fails closed when native WPS Writer is unavailable; no auxiliary
renderer is accepted for this layout-dependent operation.

Successful native WPS PDF conversions include
`native_render_provenance` conforming to
`tiwater.convert-native-render-provenance/v1`. The provider records the WPS
package build and executable hash, OS/.NET runtime identity, a count/hash-only
fontconfig inventory based on font family, style, and font bytes, input/output
SHA-256 identities, and PDF page count. Native conversion fails closed when
any required provenance cannot be collected. Consumers should retain this
provider-owned object unchanged rather than probing the host independently.

XLS conversion prefers ET through `pywpsrpc` when available,
because WPS generally preserves legacy workbook formatting better on Linux. Set
`TIWATER_WPSRPC_PYTHON` to the pywpsrpc venv Python if
it is not installed at `~/.local/share/tiwater/wpsrpc-venv/bin/python`.
`xvfb-run` and the WPS `et` command must be available. If WPS conversion is not
available or fails, the converter falls back to LibreOffice/soffice and then to
the built-in NPOI converter.

DOC, DOCX, ODT, and RTF PDF conversion prefers WPS through `pywpsrpc`
when `wps`, `xvfb-run`, `dbus-run-session`, and the configured WPS RPC Python
are available. The converter launches WPS with an isolated D-Bus session and
writable per-conversion XDG cache/runtime directories so it remains usable from
restricted automation sandboxes. The
JSON result records `backend: "wps"` or `backend: "libreoffice"` and a
fallback reason. Set `TIWATER_OFFICE_PDF_BACKEND=wps` to require the
native WPS backend and fail closed when it is unavailable; use `libreoffice` to
request the auxiliary backend explicitly.

XLS and XLSX PDF conversion likewise prefers ET through
`pywpsrpc` when `et`, `xvfb-run`, `dbus-run-session`, and the configured RPC
Python are available. Set `TIWATER_OFFICE_PDF_BACKEND=et` to
require that native backend and fail closed; it never treats LibreOffice output
as ET proof.

PPT, PPTX, and ODP PDF conversion prefers WPP through `pywpsrpc`
when `wpp`, `xvfb-run`, `dbus-run-session`, and the configured RPC Python are
available. Set `TIWATER_OFFICE_PDF_BACKEND=wpp` to require that
native backend and fail closed; it never treats LibreOffice output as WPS
Presentation proof.

On macOS, set `TIWATER_WPS_OFFICE_LIMA_INSTANCE` to the name of a configured
Lima Linux instance that exposes the documented `/tmp/tiwater-wps-render` shared
directory and has the published WPS, Spreadsheets, and Presentation runtimes installed. The converter
then stages each input and output in a unique shared directory and invokes
`limactl shell <instance>` itself; it reports the requested `wps` or
`et`, or `wpp` backend and fails closed if the configured instance cannot render. This is an explicit
runtime configuration, not an arbitrary command hook.

LibreOffice-backed PDF conversion requires `soffice`. If it is not on `PATH`, set one of:

- `TIWATER_SOFFICE`
- `SOFFICE`
- `LIBREOFFICE_PATH`
