# Adding an upload backend

Read [`upload-backends.md`](upload-backends.md) first. `backends/minio.py` is
the shortest complete example.

## 1. Settings

In `config.py`:

1. Add a `<Name>Settings(BaseModel)` with the required fields and a
   `@model_validator(mode="after")` that rejects empty ones with a clear
   message. Look at `MinioSettings`.
2. Add the member to `StorageStrategy`.
3. Add an `Optional[<Name>Settings]` field to `StorageSettings` and a branch
   in `validate_strategy_requirements()` so the strategy fails at startup
   when its settings are missing.
4. Populate it in `Config.from_env()` under `elif strategy == StorageStrategy.<NAME>.value:`,
   reading only the variables you document.

## 2. The backend module

`upload_tools/backends/<name>.py`:

```python
def upload_to_<name>(file_object, file_name: str, cfg, signed_url_expires_in: int):
    if not cfg:
        logger.error("<Name> configuration not provided")
        return None
    try:
        import <sdk>  # lazy: not every deployment installs it
    except ImportError:
        logger.error("<sdk> is not installed ...")
        return None
    content_type = get_content_type(file_name)
    try:
        file_object.seek(0)
        ...upload...
        url = ...signed URL valid for signed_url_expires_in...
        return (
            "Link to created document to be shared with user in markdown format: "
            f"{url} . Link is valid for {signed_url_expires_in} seconds."
        )
    except Exception as e:
        logger.error("Error uploading to <Name>: %s", e)
        return None
```

Return a string on success. On failure return `None` after logging, or
raise `RuntimeError`. Never return an error message as a string; the
dispatcher would relay it as success.

## 3. Wire the dispatcher

In `upload_tools/main.py`, import the function, add an
`elif UPLOAD_STRATEGY == StorageStrategy.<NAME>:` branch to **both**
`upload_file()` and `upload_file_async()`, and add a line to the strategy
announcement block. Export it from `upload_tools/backends/__init__.py`.

If the backend must be async (it talks to a service with an async client
only), follow `librechat.py` instead: handle it in `upload_file_async()`
only and make `upload_file()` raise for it. Then also update
`is_librechat_strategy()` and the branch in
`librechat_integration.upload_and_format_response()`, and know that the
dynamic template tools will not support it until they stop calling the
sync entry point.

## 4. Dependencies and docs

- Add the SDK to `requirements.txt`.
- Add the variables to `.env.example` and to
  [`../configuration.md`](../configuration.md). The drift test in
  `tests/test_config_docs.py` fails until both list every variable
  `from_env()` reads.
- Add a row to the backends table in [`upload-backends.md`](upload-backends.md).

## 5. Tests

Add `tests/test_<name>_upload.py` in the style of `tests/test_s3_upload.py`:
build a `Config` from patched environment variables, assert the validator
accepts and rejects what it should, and patch the SDK client to check the
call and the returned string. Do not hit the network.
