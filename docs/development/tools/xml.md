# XML tool (`xml_tools`)

Validates a caller-supplied XML string and saves it as a file. The smallest
tool in the server, and the reference example of the buffer-then-upload
shape. For the request path around the build step, see
[`../architecture.md`](../architecture.md).

## Entry points

| Name | Where | Used by |
|------|-------|---------|
| `create_xml_file` | MCP tool declared in `main.py` | MCP clients |
| `_create_xml_buffer()` | `xml_tools/base_xml_tool.py` | `main.py`, via `run_blocking` |
| `create_xml_file()` | `xml_tools/base_xml_tool.py` | direct library use; builds and uploads synchronously |
| `validate_xml()` | `xml_tools/base_xml_tool.py` | the buffer function; reusable on its own |

## Pipeline

```
xml_content
  │
  ▼  base_xml_tool._create_xml_buffer()
  ├─ strip surrounding whitespace
  ├─ validate_xml()             defusedxml.ElementTree.fromstring()
  │                             → XMLValidationError on parse error or a defused
  │                               security error (XXE, entity expansion, external DTD)
  ├─ read encoding= from the declaration, default UTF-8;
  │  prepend a UTF-8 declaration if there is none
  ▼
BytesIO of xml_content.encode(encoding)
```

## How it works

`defusedxml` is used instead of the standard library parser so a document
with an external entity, a DTD fetch or a billion-laughs expansion is
rejected rather than parsed. The parse result is discarded; only success or
failure matters. The bytes written are the caller's string, not a
re-serialisation, so formatting and comments survive.

Two exception types are defined: `XMLValidationError` for bad input and
`XMLFileCreationError` for anything after validation. Both reach the MCP
client as a `ToolError` through the handler in `main.py`.

## Extension points

This tool is deliberately minimal. The natural additions are schema
validation (XSD or RELAX NG) inside `validate_xml()`, or pretty-printing
before the encode step. Either belongs in `base_xml_tool.py`.

## Invariants and gotchas

- **The declared encoding is trusted.** A declaration naming an encoding
  Python does not know raises `LookupError` inside the encode step, which is
  wrapped as `XMLFileCreationError`. Nothing validates the encoding name
  against the bytes.
- **Well-formedness only.** No schema, no namespace checks.

## Tests

`tests/test_xml_creation.py` covers validation, the security rejections, the
encoding declaration handling and the upload path.
