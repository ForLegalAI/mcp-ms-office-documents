# Security

## Reporting a vulnerability

Please do not open a public issue for a security problem. Email the
maintainers at the address on the
[ForLegalAI GitHub organisation page](https://github.com/ForLegalAI), or use
GitHub's private vulnerability reporting on this repository if it is
enabled. Include the version or image tag and a way to reproduce.

## What the server protects, and what it assumes

**API key.** When `API_KEY` is set, every MCP request must carry it as a
`Bearer` token, a plain `Authorization` value, or `x-api-key`. Comparison is
constant-time and failures are rate-limited in the log. The health routes
`/healthz`, `/readyz` and `/livez` are deliberately outside this gate; they
return static strings only. Without an API key the server accepts any
caller, which is fine on localhost and nowhere else.

**Admin UI.** Opt-in, gated by `ADMIN_PASSWORD` (falling back to `API_KEY`),
CSRF-protected, with a 10 MB upload cap. It writes files into
`custom_templates/` and `config/`, so anyone with the password can change
what the tools produce.

**Image fetching.** The Word and PowerPoint tools download images from
caller-supplied URLs. Every hostname is resolved and rejected unless
globally routable, and every redirect hop is re-checked, so a URL cannot be
used to reach loopback, private networks or the cloud metadata address.
`SSRF_ALLOW_PRIVATE_ADDRESSES=true` disables this; set it only on a network
you trust. DNS rebinding between the check and the connection remains a
known residual risk. Downloads are capped at 10 MB and limited to image
MIME types.

**XML.** Caller-supplied XML is parsed with `defusedxml`, which rejects
external entities, DTD fetches and entity-expansion attacks.

**LibreChat mode.** With `UPLOAD_STRATEGY=LIBRECHAT`, the `X-User-Id` header
is trusted as-is to attribute uploads. The assumption is that only LibreChat
can reach the server, which the API key enforces. Do not expose a
LibreChat-mode server without one.

**Generated files.** Cloud backends return time-limited signed URLs
(`SIGNED_URL_EXPIRES_IN`). Filenames supplied by the caller are sanitised to
word characters, hyphens, dots and underscores before use as object names.

**Not sanitised.** The HTML body of an email draft is inserted verbatim into
the `.eml`. Markdown and slide text are rendered as document content, not
executed.

## Supported versions

Only the latest release receives fixes. Docker images are rebuilt on Alpine
with the OS packages upgraded at build time; pull a new tag to pick up base
image fixes.
