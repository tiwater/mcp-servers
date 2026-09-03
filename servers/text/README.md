# tiwater-text-mcp

`tiwater-text-mcp` publishes read-only observation for explicitly supported
plain-text files. `text_inspect` reports the exact byte identity, lossless
decoding facts, line count, and bounded opening lines. `text_read_lines` reads
one explicit zero-based line page and reports its continuation.

Supported inputs are `.txt`, `.text`, `.log`, `.csv`, `.tsv`, `.md`, and
`.markdown` files encoded as valid UTF-8, UTF-8 with BOM, UTF-16LE with BOM, or
UTF-16BE with BOM. The provider rejects binary content and never guesses an
encoding.

The tools do not interpret key-value pairs, records, sections, Markdown, or
business fields, and never modify or transcode the source.
