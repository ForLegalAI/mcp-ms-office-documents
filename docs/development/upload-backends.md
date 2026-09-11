# Upload backends

Every tool ends by handing a `BytesIO` to the upload layer, which stores it
somewhere and returns what the MCP client should be told. This page covers
the dispatch, the six backends, naming, and the rules a new backend must
follow. The configuration variables for each are listed in
[`../configuration.md`](../configuration.md).

## Dispatch

`upload_tools/main.py` reads `UPLOAD_STRATEGY` once at import into
`UPLOAD_STRATEGY`, and exposes:

| Function | Sync/async | Handles | Called from |
|----------|------------|---------|-------------|
| `upload_file()` | sync | LOCAL, S3, GCS, AZURE, MINIO; **raises** `RuntimeError` for LIBRECHAT | dynamic template tools, the public `*_to_*` wrappers, and `upload_and_format_response()` through `run_blocking` |
| `upload_file_async()` | async | all six; LibreChat natively, the others by calling the sync backend inline | `upload_and_format_response()` for LibreChat only |
| `is_librechat_strategy()` | sync | the branch decision | `librechat_integration.py` |

Both entry points do the same three things: resolve the unique-prefix
default, generate the object name, call the backend. A backend signals
failure by raising or by returning `None`; the dispatcher converts `None`
into `RuntimeError("Upload to <strategy> failed. Check server logs")`. Any
other exception is wrapped as `RuntimeError("Error uploading document: …")`.

`upload_file_async()` calls the sync backends directly, without offloading.
That is fine only because its one caller, `upload_and_format_response()`,
uses it for LibreChat alone and routes everything else through
`run_blocking(upload_file)`. Do not call `upload_file_async()` for a
traditional backend from the event loop.

## Object naming

`upload_tools/utils.py`:

- `generate_unique_object_name(suffix)` → `<uuid4>.<suffix>` when the caller
  gave no `file_name`.
- `generate_named_object_name(filename, suffix, add_unique_prefix)` →
  `<safe>.<suffix>` or `<8 hex>_<safe>.<suffix>`. `sanitize_filename()` keeps
  word characters, whitespace (collapsed to `_`), hyphens and dots, cuts at
  100 characters, and falls back to `document`.
- `get_content_type(file_name)` maps the extension to a MIME type for the
  cloud backends and **raises `ValueError` for an unknown extension**. A new
  document type must be added here or every cloud upload of it fails.
  Matching is by substring, so a name such as `notes.pptx_v2.docx` is
  detected as PowerPoint; the sanitiser keeps dots, so this can happen
  ([#116](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/116)).

The unique-prefix default lives in the dispatcher: `None` becomes `True` for
traditional backends and `False` for LibreChat, which prefixes files itself.

## The backends

| Strategy | Module | Client | Auth | Returns |
|----------|--------|--------|------|---------|
| `LOCAL` | `backends/local.py` | `open()` | none | `"Document saved to <cwd>/output/<name>"` |
| `S3` | `backends/s3.py` | boto3 | explicit key/secret/region, or the default credential chain (IRSA, SSO, instance profile) when none are set | sentence with a pre-signed GET URL |
| `GCS` | `backends/gcs.py` | google-cloud-storage | service-account JSON at `GCS_CREDENTIALS_PATH`, else Application Default Credentials; federated credentials sign through the IAM signBlob API | sentence with a v4 signed URL |
| `AZURE` | `backends/azure.py` | azure-storage-blob | account key; optional custom endpoint for sovereign clouds | sentence with a SAS URL |
| `MINIO` | `backends/minio.py` | boto3 against a custom endpoint | access key and secret; path-style addressing and signature v4; SSL verification configurable | sentence with a pre-signed GET URL |
| `LIBRECHAT` | `backends/librechat.py` | httpx (async) | `X-Service-Token` plus the user's `X-User-Id` | a dict of file metadata, formatted by `format_file_artifact()` |

The signed-URL lifetime is `SIGNED_URL_EXPIRES_IN` seconds for all four cloud
backends, and the sentence returned to the model states it. The sentence is
deliberate: it is what the model relays to the user.

Every cloud backend imports its SDK inside the function, so a deployment
that uses one backend does not need the others installed. `s3.py` explains
the two credential modes in its module docstring.

### LibreChat

`upload_to_librechat()` posts the file as multipart form data to
`LIBRECHAT_SERVICE_URL` with the service token and the user headers taken
from the MCP request. It distinguishes a 401 (token mismatch) and a 400
(server-side validation) from other HTTP errors, checks the JSON `success`
flag, and returns normalised metadata including a `download_url` built from
the user id and file id. `format_file_artifact()` wraps that as
`{"result": {"message", "file": {…}}}`, the shape LibreChat's MCP client
renders as an attachment. The `download_url` is computed but not included
in the artifact.

This backend is experimental and depends on a LibreChat build with MCP file
artifact support. See [`../deployment.md`](../deployment.md).

## Rules for a backend

1. Take `(file_object, object_name, settings, signed_url_expires_in)`; call
   `file_object.seek(0)` before reading.
2. Import the SDK inside the function.
3. Return a string on success, `None` on failure after logging the reason,
   or raise `RuntimeError`. Never return an error message as a string.
4. Set the content type from `get_content_type()`.
5. Add a `Settings` model and its validator in `config.py`, populate it in
   `Config.from_env()` under the new strategy, add the enum member, and
   document the variables. See [`adding-a-backend.md`](adding-a-backend.md).

## Tests

| File | Covers |
|------|--------|
| `tests/test_upload_unique_prefix.py` | Object naming and the prefix default through the dispatcher |
| `tests/test_s3_upload.py` | The two credential modes and their validation |
| `tests/test_librechat_integration.py` | Header extraction, `upload_and_format_response()` branching, artifact shape |
