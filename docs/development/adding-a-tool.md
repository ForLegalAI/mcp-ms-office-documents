# Adding a document tool

The shape every existing tool follows. Read [`architecture.md`](architecture.md)
first; [`tools/xml.md`](tools/xml.md) is the smallest worked example.

## 1. Create the package

```
<type>_tools/
  __init__.py          re-export the two entry points
  base_<type>_tool.py  _<build>_buffer() and the public wrapper
  helpers.py           optional; parsing and formatting
```

Add a module docstring to each file saying what it owns and which entry
point `main.py` uses.

## 2. Write the buffer function

```python
def _create_<type>_buffer(content: str, ...) -> io.BytesIO:
    """Build the document and return it positioned at the start."""
```

Rules:

- Input to bytes only. No upload, no request context, no `get_http_request()`.
  It runs on a worker thread.
- Raise `ValueError` for input the caller can fix (the handler passes the
  message through) and `RuntimeError` for anything else.
- If the builder ever has to work around something rather than fail, return
  `(BytesIO, warnings)` so the handler can attach the warnings to the result
  with `main._with_warnings()` — records with a stable `code` and `severity`,
  not sentences. Do not log and move on. Build the channel on
  `warning_channel.WarningChannel` and put the codes and their severities in
  your package's `warnings.py`, as `docx_tools` and `xlsx_tools` do; see
  [`shared-modules.md`](shared-modules.md#warning_channelpy).
- Templates come from `template_utils.find_file_in_template_dirs()`, never
  a hard-coded path.
- Images come from `image_utils.load_image()`, never `requests` directly.

## 3. Write the public wrapper

```python
def create_<type>(content, file_name=None, ...) -> str:
    buffer = _create_<type>_buffer(content, ...)
    try:
        return upload_file(buffer, "<ext>", filename=file_name)
    finally:
        buffer.close()
```

This exists for direct library use and tests. It cannot return a LibreChat
artifact; that is expected.

## 4. Register the MIME type

Add the extension to `MIME_TYPES` in `upload_tools/utils.py` — one table,
which `backends/librechat.py` imports. Without it, every cloud upload of the
new type raises `ValueError`.

## 5. Declare the tool in `main.py`

```python
@mcp.tool(
    name="create_<type>_from_...",
    description="One or two sentences the model reads when choosing a tool.",
    tags={"<type>", ...},
    annotations={"title": "Human-readable title"},
)
async def create_<type>_document(
    content: Annotated[str, Field(description="Everything the model must know about the input format.")],
    file_name: Annotated[Optional[str], Field(description="...", default=None)] = None,
    add_unique_prefix: Annotated[Optional[bool], Field(description="...", default=None)] = None,
) -> Union[str, dict]:
    try:
        file_buffer = await run_blocking(_create_<type>_buffer, content, ...)
        user_context = extract_user_context_from_request()
        result = await upload_and_format_response(
            file_buffer, "<ext>", file_name, user_context,
            f"<Type> '{file_name or 'document'}' created successfully.",
            add_unique_prefix=add_unique_prefix,
        )
        file_buffer.close()
        return result
    except Exception as e:
        logger.error("Error creating <type>: %s", e, exc_info=True)
        raise ToolError(f"Error creating <type>: {e}")
```

Copy an existing handler rather than typing this; the Excel handler is the
shortest. Points that matter:

- `await run_blocking(...)` for the build. Never call the buffer function
  directly.
- `extract_user_context_from_request()` on the event loop, before the
  upload, not inside the buffer function.
- `upload_and_format_response()` for the upload. Do not call `upload_file()`
  here; the helper decides between a URL and an artifact and offloads the
  sync upload.
- `add_unique_prefix` defaults to `None` and is passed through unchanged.
- Every parameter has a `Field(description=...)`. The description is the
  documentation the calling model sees; for a format the model has to
  learn, it is the **only** documentation. Keep it, the user reference page
  and the tests in step.
- Avoid `Optional[X]` in types the model must see clearly described; several
  clients drop the description next to an `anyOf`. Use a plain type with a
  `None` default where you can.

## 6. Tests

Create `tests/test_<type>_creation.py`. Call the buffer function directly
and inspect the bytes; patch `<type>_tools.base_<type>_tool.upload_file` if
you test the public wrapper. Save inspection copies under
`tests/output/<type>/` if the format benefits from opening by hand. See
[`testing.md`](testing.md).

## 7. Documentation

- Add `docs/development/tools/<type>.md` in the same shape as the others:
  entry points, pipeline, module map, how it works, extension points,
  invariants, tests, known limitations.
- Add a user reference section to `docs/` if the input format needs one, and
  link it from the README's feature table.
- Add a row to the feature table in `Readme.md`.
