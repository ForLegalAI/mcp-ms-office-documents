# Configuration

The server is configured entirely through environment variables, normally
placed in a `.env` file next to `docker-compose.yml` (copy `.env.example` to
start). Every variable below is read once at startup.

## Basic settings

| Variable | Description | Default |
|----------|-------------|---------|
| `DEBUG` | Enable debug logging (`1`, `true`, `yes`, `on`) | _(off)_ |
| `API_KEY` | Protect the server with an API key (see Authentication below) | _(disabled)_ |
| `UPLOAD_STRATEGY` | Where to save files: `LOCAL`, `S3`, `GCS`, `AZURE`, `MINIO`, `LIBRECHAT` | `LOCAL` |
| `SIGNED_URL_EXPIRES_IN` | How long cloud download links stay valid (seconds) | `3600` |
| `RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED` | Offload blocking tool work to a thread pool, keeping the event loop free for health probes & concurrent requests | `true` |
| `RUN_BLOCKING_MAX_WORKERS` | Maximum concurrent worker threads for blocking tool calls | `4` |
| `STATELESS_HTTP` | Run the streamable-HTTP transport without server-side sessions, so requests may land on any replica. Required for two or more replicas; see [Deployment](deployment.md) | `false` |
| `SSRF_ALLOW_PRIVATE_ADDRESSES` | Allow image URLs that resolve to private, loopback or link-local addresses. Removes the SSRF guard entirely; enable only when images are served from inside your own network | `false` |
| `EMAIL_DEFAULT_LANGUAGE` | BCP-47 tag stamped on an email draft when the call does not give one, setting the proofing language Outlook checks it in. Set your own (e.g. `cs-CZ`, `de-DE`) if your drafts are not in English. The tool's `language` argument still wins per call. **Was a hard-coded `cs-CZ` until [#116](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/116)** — a Czech deployment upgrading past it should set this | `en-US` |

## Authentication

Set `API_KEY` in your `.env` to require an API key for all requests:

```
API_KEY=your-secret-key
```

Clients can send the key in any of these headers:

| Header | Format |
|--------|--------|
| `Authorization` | `Bearer your-secret-key` |
| `Authorization` | `your-secret-key` |
| `x-api-key` | `your-secret-key` |

Leave `API_KEY` empty or unset to allow all requests without authentication.

The health-probe routes (`/healthz`, `/readyz`, `/livez`) are the one exception — they stay reachable without a key so orchestrators can poll them. See [Deployment](deployment.md#health-probes-and-the-thread-pool).

## Admin UI

| Variable | Description | Default |
|----------|-------------|---------|
| `ADMIN_ENABLED` | Mount the template-admin UI (`1`, `true`, `yes`, `on`) | _(off)_ |
| `ADMIN_PASSWORD` | Password for the admin UI. Falls back to `API_KEY` when unset | _(falls back)_ |
| `ADMIN_PATH` | URL prefix the admin UI is mounted under | `/admin` |

See [Template Admin UI](admin-ui.md).

## Storage backends

`UPLOAD_STRATEGY` selects one backend for the whole server. Each backend reads only its own variables; a missing required variable stops the server at startup with a configuration error.

### AWS S3 Storage

Set `UPLOAD_STRATEGY=S3` and provide:

| Variable | Description | Required |
|----------|-------------|----------|
| `S3_BUCKET` | S3 bucket name | ✅ Always |
| `AWS_ACCESS_KEY` | AWS access key ID | ⚠️ See below |
| `AWS_SECRET_ACCESS_KEY` | AWS secret access key | ⚠️ See below |
| `AWS_REGION` | AWS region (e.g., `us-east-1`) | ⚠️ See below |

**Credential modes:**

- **Explicit credentials** — Set all three of `AWS_ACCESS_KEY`, `AWS_SECRET_ACCESS_KEY`, and `AWS_REGION`. Recommended for simple setups.

- **AWS default credential chain** — Leave the credential variables unset and boto3 will automatically discover credentials from the standard chain:
  - `AWS_ACCESS_KEY_ID` / `AWS_SECRET_ACCESS_KEY` environment variables
  - Shared credential / config files (`~/.aws/credentials`)
  - AWS SSO sessions (`aws sso login`) — useful for local development
  - **IRSA (IAM Roles for Service Accounts)** — for AWS EKS deployments
  - ECS container credentials / EC2 instance metadata (IMDSv2)

  In this mode only `S3_BUCKET` is required; region is resolved automatically.

### Google Cloud Storage

Set `UPLOAD_STRATEGY=GCS` and provide:

| Variable | Description |
|----------|-------------|
| `GCS_BUCKET` | GCS bucket name |
| `GCS_CREDENTIALS_PATH` | Path to service account JSON (default: `/app/config/gcs-credentials.json`) |

Mount the credentials file via `docker-compose.yml` volumes.

### Azure Blob Storage

Set `UPLOAD_STRATEGY=AZURE` and provide:

| Variable | Description |
|----------|-------------|
| `AZURE_STORAGE_ACCOUNT_NAME` | Storage account name |
| `AZURE_STORAGE_ACCOUNT_KEY` | Storage account key |
| `AZURE_CONTAINER` | Blob container name |
| `AZURE_BLOB_ENDPOINT` | _(Optional)_ Custom endpoint for sovereign clouds |

### MinIO / S3-Compatible Storage

Set `UPLOAD_STRATEGY=MINIO` and provide:

| Variable | Description | Default |
|----------|-------------|---------|
| `MINIO_ENDPOINT` | MinIO server URL (e.g., `https://minio.example.com`) | _(required)_ |
| `MINIO_ACCESS_KEY` | Access key | _(required)_ |
| `MINIO_SECRET_KEY` | Secret key | _(required)_ |
| `MINIO_BUCKET` | Bucket name | _(required)_ |
| `MINIO_REGION` | Region | `us-east-1` |
| `MINIO_VERIFY_SSL` | Verify SSL certificates | `true` |
| `MINIO_PATH_STYLE` | Use path-style URLs (recommended for MinIO) | `true` |

Make sure the bucket exists and your credentials have `PutObject`/`GetObject` permissions.

### LibreChat Integration (Experimental)

> ⚠️ **Note:** This feature is **not yet implemented in the official LibreChat repository**. It requires a custom LibreChat build with MCP file artifact support. A test version is available at:
> 
> 👉 https://github.com/geodanchev/LibreChat/tree/feature/mcp-ms-office-docs-integration

Set `UPLOAD_STRATEGY=LIBRECHAT` to enable seamless document generation within LibreChat conversations. Documents are uploaded to LibreChat's service endpoint and returned as MCP file artifacts that appear as attachments in the chat UI.

| Variable | Description | Required |
|----------|-------------|----------|
| `LIBRECHAT_SERVICE_URL` | Full URL to LibreChat's service files endpoint | ✅ |
| `LIBRECHAT_SERVICE_TOKEN` | Service token for authentication (must match `MCP_SERVICE_TOKEN` in LibreChat) | ✅ |

**Typical URLs:**
- Inside Docker network: `http://api:3080/api/service/files`
- Local development: `http://localhost:3080/api/service/files`

**Generate a service token:**
```bash
openssl rand -hex 32
```

**LibreChat `librechat.yaml` configuration:**
```yaml
mcpServers:
  Office-documents:
    type: streamable-http
    url: http://mcp-office-docs:8958/mcp
    headers:
      X-User-Id: "{{LIBRECHAT_USER_ID}}"
      X-User-Email: "{{LIBRECHAT_USER_EMAIL}}"
```

The `X-User-Id` and `X-User-Email` headers are automatically populated by LibreChat and used to associate uploaded files with the correct user.

## Performance and health probes

Thread-pool offloading, the health-check routes and the stateless transport
are described in [Deployment](deployment.md).
