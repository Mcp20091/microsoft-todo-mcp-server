# Self-hosting over the network

`dist/http-server.js` serves the MCP server over Streamable HTTP at `http://<host>:8787/mcp`, so several devices
can share one instance and one set of Microsoft tokens. Every request must carry
`Authorization: Bearer <MCP_AUTH_TOKEN>`.

| Variable            | Default                              | Purpose                                             |
| ------------------- | ------------------------------------ | --------------------------------------------------- |
| `MCP_AUTH_TOKEN`    | (required, 16+ chars)                | Shared secret clients send as a bearer token        |
| `PORT`              | `8787`                               | Listen port                                         |
| `HOST`              | `127.0.0.1` (`0.0.0.0` in the image) | Listen address                                      |
| `MSTODO_TOKEN_FILE` | `/data/tokens.json` in the image     | Microsoft tokens; refreshed tokens are written back |

The image is published to `ghcr.io/mcp20091/microsoft-todo-mcp-server` by `.github/workflows/docker.yml`
(make the package public in GitHub, or `docker login ghcr.io` on the host first).

## 1. Get Microsoft tokens into the data folder

Either copy an existing `tokens.json` (from `pnpm run auth:device`) into the data folder, or sign in from inside the
running container:

```bash
docker exec -it mstodo-mcp node dist/device-auth.js <CLIENT_ID>
```

No restart is needed; the token file is read on each request.

## 2a. TrueNAS SCALE (24.10+)

1. Create a dataset, e.g. `tank/apps/mstodo-mcp`, and give the `apps` user (568) write access.
2. Apps → Discover Apps → ⋮ → **Install via YAML**, name it `mstodo-mcp`, and paste `compose.yaml` with:
   - `MCP_AUTH_TOKEN` set to the output of `openssl rand -hex 32`
   - the volume changed to `/mnt/tank/apps/mstodo-mcp:/data`
   - `user: "568:568"` uncommented
3. Put `tokens.json` in the dataset, or run the device sign-in above from the TrueNAS shell
   (`sudo docker exec -it mstodo-mcp ...`).

## 2b. Proxmox LXC

Easiest: a Debian 12 LXC with Docker (`curl -fsSL https://get.docker.com | sh`), then put `compose.yaml` in
`/opt/mstodo-mcp`, edit the token, and run `docker compose up -d`. Unprivileged LXCs need `nesting=1` (and usually
`keyctl=1`) under Options → Features.

Without Docker, in a plain Debian LXC with Node 22:

```bash
git clone -b feat/http-transport https://github.com/Mcp20091/microsoft-todo-mcp-server.git /opt/mstodo-mcp
cd /opt/mstodo-mcp && corepack enable && pnpm install && pnpm run build
```

`/etc/systemd/system/mstodo-mcp.service`:

```ini
[Unit]
Description=Microsoft To Do MCP
After=network-online.target

[Service]
WorkingDirectory=/opt/mstodo-mcp
Environment=HOST=0.0.0.0 PORT=8787 MSTODO_TOKEN_FILE=/var/lib/mstodo-mcp/tokens.json
EnvironmentFile=/etc/mstodo-mcp.env
ExecStart=/usr/bin/node dist/http-server.js
Restart=on-failure
StateDirectory=mstodo-mcp
DynamicUser=yes

[Install]
WantedBy=multi-user.target
```

with `MCP_AUTH_TOKEN=...` in `/etc/mstodo-mcp.env` (`chmod 600`), then `systemctl enable --now mstodo-mcp`.

## 3. Connect clients

Claude Code:

```bash
claude mcp add --transport http --scope user microsoft-todo http://nas.lan:8787/mcp --header "Authorization: Bearer <MCP_AUTH_TOKEN>"
```

Claude Desktop (`claude_desktop_config.json`), via the `mcp-remote` bridge:

```json
{
  "mcpServers": {
    "microsoft-todo": {
      "command": "npx",
      "args": ["-y", "mcp-remote", "http://nas.lan:8787/mcp", "--allow-http", "--header", "Authorization:${AUTH}"],
      "env": { "AUTH": "Bearer <MCP_AUTH_TOKEN>" }
    }
  }
}
```

Claude on the web and mobile adds custom connectors from Anthropic's cloud, so it needs a public HTTPS URL with OAuth;
a LAN-only bearer-token server like this one is not reachable from there.

Health check: `curl http://nas.lan:8787/healthz`.
