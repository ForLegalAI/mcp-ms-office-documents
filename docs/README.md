# Documentation

**Using the server**

- [Configuration](configuration.md) — every environment variable, by backend
- [Deployment](deployment.md) — Docker Compose, running locally, health probes, replicas, LibreChat
- [Connecting an AI client](clients.md) — Claude Desktop, LibreChat, Cursor
- [Markdown reference](markdown-reference.md) — everything the Word and Excel tools accept
- [PowerPoint slide reference](powerpoint-slides.md) — the fourteen slide types and their fields
- [Templates](templates.md) — custom designs, named PowerPoint templates, dynamic Word and email tools
- [Template Admin UI](admin-ui.md) — manage templates from the browser

**Working on the code**

- [Architecture](development/architecture.md) — the request path, threading, errors, config, security boundaries
- Tools: [Word](development/tools/word.md) · [Excel](development/tools/excel.md) · [PowerPoint](development/tools/powerpoint.md) · [Email](development/tools/email.md) · [XML](development/tools/xml.md)
- [Dynamic template tools](development/dynamic-templates.md)
- [Upload backends](development/upload-backends.md)
- [Admin UI internals](development/admin-ui.md) — the component layer, the kind table, the no-external-assets rule
- [Shared root modules](development/shared-modules.md) — config, threading, uploads, images, the warnings channel
- How to: [add a tool](development/adding-a-tool.md) · [add a backend](development/adding-a-backend.md) · [test](development/testing.md)

See also [CONTRIBUTING.md](../CONTRIBUTING.md), [SECURITY.md](../SECURITY.md) and, for coding agents, [AGENTS.md](../AGENTS.md).
