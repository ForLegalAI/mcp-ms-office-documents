# Connecting an AI client

Point your MCP-compatible client to the server endpoint:

```
http://localhost:8958/mcp
```

**Examples for popular clients:**

## Claude Desktop

Add to your Claude Desktop MCP config:

```json
{
  "mcpServers": {
    "office-documents": {
      "url": "http://localhost:8958/mcp"
    }
  }
}
```

## LibreChat

Add the server to your `librechat.yaml` configuration under `mcpServers`:

```yaml
mcpServers:
  office-documents:
    type: streamable-http
    url: http://mcp-office-docs:8958/mcp
```

> **Note:** If LibreChat and this server run in the same Docker network, use the container name (`mcp-office-docs`) as the hostname. If they run separately, use `http://localhost:8958/mcp` instead.

To place both services on the same network, add a shared network in your `docker-compose.yml`:

```yaml
services:
  mcp-office-docs:
    # ...existing config...
    networks:
      - shared

  librechat:
    # ...existing config...
    networks:
      - shared

networks:
  shared:
    driver: bridge
```

## Cursor / Other MCP Clients

Use the SSE/streamable HTTP transport and set the endpoint URL to:

```
http://localhost:8958/mcp
```

If you have authentication enabled, add the API key header as required by your client.

