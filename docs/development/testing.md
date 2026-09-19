# Testing

```bash
pip install -r requirements-dev.txt     # requirements.txt + ruff + pytest-asyncio
ruff check .
pytest                                  # everything, incl. network-marked tests
pytest -m "not network"                 # what CI runs
pytest tests/test_docx_base.py -q       # one module
```

`pytest.ini` sets `asyncio_mode = auto`, so `async def` tests need no
decorator, and declares the `network` marker. Tests marked `network` reach a
public image host; CI excludes them so a third-party outage cannot fail an
unrelated pull request. No `.env` is needed: `config.py` defaults to the
`LOCAL` strategy and INFO logging.

## Conventions

**Build, don't upload.** Every package has a way to get bytes without a
backend:

| Package | Pattern |
|---------|---------|
| Word | call `_markdown_to_doc()` and inspect the `Document`, or `_markdown_to_word_buffer()`, which returns `(buffer, warnings)`, and reopen the bytes |
| Excel | `tests/test_xlsx_creation.py::_create_workbook_from_markdown()` patches `xlsx_tools.base_xlsx_tool.upload_file`, captures the buffer and returns `load_workbook()` of it; `_markdown_to_excel_buffer()` returns `(buffer, warnings)` |
| PowerPoint | instantiate `PowerpointPresentation(slides, ...)`, call `.save()`, reopen with `Presentation()`; read `.warnings` for the warning records (`code`, `slide`, `severity`, `message`) or `.warning_messages` to assert on their text |
| Email, XML | call the buffer function directly. For email, parse the result with `email.message_from_bytes()` and base64-decode the payload to see the body — `tests/test_email_creation.py` shows the shape |
| Dynamic templates | register against a fresh `FastMCP()` instance and call the tool function; patch `upload_file` in the tool module |
| Template resolution | monkeypatch the four directory constants on `template_utils` at temp dirs and place real files; see `tests/test_template_resolution.py` |

**Warnings without a build.** To assert on what a builder worked around, take
the warnings half of the buffer function's return value, or pass a channel of
your own (`docx_tools.warnings.channel()`, `xlsx_tools.warnings.channel()`)
into the function under test — every site takes one as a keyword argument.
Assert on the `code` and the location, not on the English.

Patch `upload_file` where it is *looked up*, i.e. in the tool module
(`docx_tools.base_docx_tool.upload_file`), not in `upload_tools`.

**Inspection output.** Tests that produce whole documents save a copy under
`tests/output/{docx,pptx,xlsx}/`, named after the test, so a rendering
change can be opened in Word, Excel or PowerPoint. The directory is
git-ignored. Files named `VISUAL_INSPECTION_*` are the ones meant for a
human look.

**Configuration.** `get_config()` is a singleton read at import time by some
modules. To test a different configuration, set the environment and reload:
`tests/test_run_blocking.py::_reload_main_with_flag()` shows the full dance
(`config`, `async_runner`, `librechat_integration`, `main` popped from
`sys.modules`, `config._CONFIG` reset). For a narrower test, build a
`Config` directly as `tests/test_s3_upload.py` does.

**Registries.** The PowerPoint template registry is cached on file mtimes;
call `pptx_tools.templates.clear_cache()` after writing a template in a
test. `metrics.reset()` clears counters.

**Mutation-testing a change.** A test-only change has no bug to regress
against, so the way to show a new test has teeth is to break the source and
watch it fail. Run those with `PYTHONDONTWRITEBYTECODE=1` and clear
`__pycache__` between rounds. Python validates a `.pyc` on the source's
*whole-second* mtime and its size, and a mutation that reorders lines changes
neither: edit, run, restore within the same second and the interpreter serves
the stale bytecode instead. It fails in the direction that matters — a real
gap looks covered, and a mutation that was never actually applied looks
killed. This is not hypothetical; it produced two wrong results while #112's
review was being worked, one of them reported before it was caught.

**Network.** Anything that would download an image is either marked
`@pytest.mark.network` or patches `image_utils.download_image`. The SSRF
guard is tested by resolving hostnames to fixed addresses, not by
connecting.

## Where tests live

One file per behaviour, named for it: `test_docx_list_restart.py`,
`test_pptx_sections.py`, `test_xlsx_tier4_robustness.py`. Regression tests
say so in the module docstring and name the issue. When you fix a bug, add
the test to the file that covers that behaviour rather than a new
`test_fixes.py`.

The per-tool pages under [`tools/`](tools/) each end with a table of which
test file covers what.
