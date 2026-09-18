# Deployment

## Docker Compose

`docker-compose.yml` runs the published image `georgx22/mcp-office-docs`
(tag `latest`, or a release tag) on port 8958 and mounts three directories:

| Host path | Container path | Purpose |
|-----------|----------------|---------|
| `./output` | `/app/output` | Generated files when `UPLOAD_STRATEGY=LOCAL` |
| `./custom_templates` | `/app/custom_templates` | Your Word, PowerPoint, email and Excel templates, plus any assets the admin UI uploads |
| `./config` | `/app/config` | `docx_templates.yaml`, `email_templates.yaml`, `pptx_templates.yaml`, the admin-written `*_templates.d/` directories, and credential files such as a GCS service-account JSON |

Configuration comes from `.env` in the same directory; see
[Configuration](configuration.md). Template and config directories are
re-read when they change, so adding a template does not need a restart.

The image is built from `Dockerfile` (Python 3.12 on Alpine, non-root user,
plus the metric-compatible font packages the PowerPoint fit estimate measures
text against) and published by the release workflow: a GitHub release builds and pushes
`georgx22/mcp-office-docs:<tag>`, and a full release also moves `latest`.
A prerelease never moves `latest`.

## Running without Docker

```bash
pip install -r requirements.txt
cp .env.example .env        # optional; defaults are LOCAL storage, INFO logging
python main.py
```

Everything works without the image's font packages; the PowerPoint tool's
overflow warnings are then measured against whichever font the host has, and
fall back to an arithmetic estimate on a host with none. Install
`fonts-crosextra-carlito`, `fonts-liberation`, `fonts-crosextra-caladea` and
`fonts-gelasio` (Debian/Ubuntu names) to match what the image measures.

The server listens on `0.0.0.0:8958` with the MCP endpoint at `/mcp`.
Template and config directories are resolved relative to the checkout
when the `/app/...` paths do not exist.

## Health probes and the thread pool

The server exposes health-check endpoints that Kubernetes (or any orchestrator) can use for startup/readiness/liveness probes. Each returns `200` with a short plain-text body:

| Endpoint | Probe | Body |
|----------|-------|------|
| `GET /healthz` | `startupProbe` — the pod has started | `ok` |
| `GET /readyz` | `readinessProbe` — the pod can receive traffic | `ready` |
| `GET /livez` | `livenessProbe` — the event loop is responsive; restart on failure | `alive` |

> **These three routes are not protected by `API_KEY`.** They are registered at the Starlette layer (via FastMCP's `@custom_route`), so they sit outside the MCP middleware stack and deliberately bypass the API-key check described under [Authentication](configuration.md#authentication) — this lets kubelet poll them without credentials. They expose no data beyond the static strings above.

Use HTTP probes rather than TCP ones: a TCP probe only proves the socket still accepts connections, so a wedged Python process with a bound listener would never be restarted. An HTTP probe forces the application itself to answer.

**Thread-pool offloading:** By default (`RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED=true`), all blocking document-generation work is dispatched to a bounded thread pool (`RUN_BLOCKING_MAX_WORKERS` threads, default 4). This keeps the asyncio event loop free to respond to health probes and handle concurrent requests — critical for Kubernetes deployments where blocked probes lead to pod restarts.

Set `RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED=false` only for local debugging or to rule out threading-related issues.

## Running more than one replica

The streamable-HTTP transport is **stateful by default**: each client gets a
session id that lives in one process's memory, so every request for that
session must reach the same replica. For two or more replicas set
`STATELESS_HTTP=true`, which makes every request independent.

Two more things assume a single instance:

- **Live template registration.** The admin UI writes template files and
  registers tools in the process that handled the request. With replicas,
  put `custom_templates/` and `config/` on shared storage and restart the
  pods after a change.
- **The LOCAL storage strategy** writes to the container's own disk. Use a
  cloud backend for anything beyond a single host.

## LibreChat

The `LIBRECHAT` upload strategy uploads generated files to LibreChat's
service endpoint so they appear as attachments in the chat. It is
experimental and requires a LibreChat build with MCP file artifact support;
the variables and the `librechat.yaml` snippet are in
[Configuration](configuration.md#librechat-integration-experimental). Note
that YAML-defined template tools are not available under this strategy
today; the built-in tools are.

## Security checklist

- Set `API_KEY` on any server reachable beyond localhost. The health routes
  stay open by design; everything else is gated.
- Set `ADMIN_PASSWORD` if the admin UI is enabled, or it falls back to the
  API key.
- Leave `SSRF_ALLOW_PRIVATE_ADDRESSES` unset unless images must come from
  inside your network.
- Signed download links expire after `SIGNED_URL_EXPIRES_IN` seconds; keep
  it as short as your clients allow.

See also [SECURITY.md](../SECURITY.md).
