# Running on TrueNAS SCALE

How to run the Microsoft To Do MCP server as a TrueNAS SCALE (24.10+) custom app, sign it in, reach it from Claude, and
roll out new versions. For the general reference (environment variables, Proxmox, client config), see
[self-hosting.md](self-hosting.md).

## 1. Install the app

1. Make a secret for MCP clients. In **System → Shell** run `openssl rand -hex 32` and keep the output somewhere safe
   (a password manager). Every client needs it.
2. Go to **Apps → Discover Apps → ⋮ (top right) → Install via YAML**, name the app `mstodo-mcp`, and paste:

   ```yaml
   services:
     mstodo-mcp:
       image: ghcr.io/mcp20091/microsoft-todo-mcp-server:latest
       container_name: mstodo-mcp
       restart: unless-stopped
       ports:
         - "8787:8787"
       environment:
         MCP_AUTH_TOKEN: PASTE_YOUR_SECRET_HERE
       volumes:
         - mstodo-data:/data

   volumes:
     mstodo-data:
   ```

3. Replace `PASTE_YOUR_SECRET_HERE` with the secret from step 1 and click **Save**.

The named volume `mstodo-data` lives in TrueNAS's apps storage, so no dataset or permissions are needed. It holds the
Microsoft sign-in and survives restarts, edits, and image updates. It is only lost if you delete the app together with
its volumes.

The image must be public on GitHub (Packages → `microsoft-todo-mcp-server` → Package settings → Danger Zone → Change
visibility → Public), otherwise TrueNAS can't pull it without logging in.

## 2. Sign in to Microsoft

Once the app shows **Running**, open a shell in the container: **Apps → Installed → mstodo-mcp → Workloads →** the
shell icon (`>_`), choosing `/bin/sh` if asked. Then run:

```sh
node dist/device-auth.js <CLIENT_ID>
```

Open the link it prints (https://microsoft.com/devicelogin), enter the code, and sign in. It ends with
`Signed in. Tokens saved to /data/tokens.json`. No restart is needed, and the server refreshes the tokens itself.

Run this in the **container** shell, which has no `sudo` and no `docker`. From the TrueNAS host shell instead, use
`sudo docker exec -it mstodo-mcp node dist/device-auth.js <CLIENT_ID>`.

You only need to sign in again if the volume is deleted, the server goes unused for about 90 days, or you revoke
sessions on your Microsoft account.

## 3. Check it works

- `curl http://<nas-ip>:8787/healthz` returns `{"ok":true}`. This only shows the server is up, not that it is signed in.
- Ask any connected Claude to "check my To Do auth status". The `auth-status` tool reports whether it is signed in and
  when the token expires.
- Opening `http://<nas-ip>:8787/` in a browser shows `Cannot GET /`. That is normal: only `/mcp` and `/healthz` exist.

## 4. Connect Claude

- **Claude Code (LAN):**
  `claude mcp add --transport http --scope user microsoft-todo http://<nas-ip>:8787/mcp --header "Authorization: Bearer <secret>"`
- **Claude on the web, desktop, and mobile:** these connect from Anthropic's cloud, so the server needs a public HTTPS
  URL, e.g. a Pangolin resource pointing at `<nas-ip>:8787`. Add it under **Settings → Connectors → Add custom
  connector** with the URL `https://<your-domain>/mcp/<secret>` and **Authentication: No sign-in**. The connector then
  appears in every Claude app on that account, including Claude Code.

Pangolin notes:

- Pangolin's basic auth (`https://user:pass@host/...`) works with Claude but not with clients that can't put credentials
  in a URL (Gemini, for example). For those, add a Pangolin rule that lets `/mcp/*` bypass auth and rely on the secret.
- Basic auth uses the `Authorization` header, so with basic auth on, the MCP secret has to go in the path
  (`/mcp/<secret>`), not in a `Bearer` header.
- A URL with the secret in it can show up in proxy logs and in `claude mcp list` output. Treat it like a password.

## 5. Change the secret

1. Generate a new one with `openssl rand -hex 32`.
2. **Apps → mstodo-mcp → Edit**, replace `MCP_AUTH_TOKEN`, then **Save**. The Microsoft sign-in is kept.
3. Update every client: the end of the connector URL, or the `Authorization` header.

## 6. Update to a new version

Pushes to the repo build a new `:latest` image on GitHub (Actions → **Docker image**, a few minutes). TrueNAS does not
pick it up by itself:

1. Pull the new image from the TrueNAS **host** shell (**System → Shell**, or SSH), not the container shell:

   ```sh
   sudo docker pull ghcr.io/mcp20091/microsoft-todo-mcp-server:latest
   ```

2. Recreate the container on it: **Apps → mstodo-mcp → Edit → Save** without changing anything. **Stop/Start** is
   not enough: it restarts the existing container, which keeps the old image.
3. Confirm the container runs the new image. These two IDs should match:

   ```sh
   sudo docker inspect mstodo-mcp --format '{{.Image}}'
   sudo docker image inspect ghcr.io/mcp20091/microsoft-todo-mcp-server:latest --format '{{.Id}}'
   ```

The sign-in in `mstodo-data` carries over, so no new device login is needed.

## Troubleshooting

| Symptom                                                    | Cause / fix                                                                        |
| ---------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `sudo: not found` or `docker: not found`                   | You are in the container shell. Use **System → Shell** for `docker` commands.      |
| `dist/device-auth.js: Permission denied`                   | Run it with `node dist/device-auth.js ...`.                                        |
| `Cannot POST /mcp/<secret>` after an update                | The container still runs the old image. Do step 6.2 (Edit → Save), then check 6.3. |
| `401 Unauthorized` (JSON body)                             | Wrong or missing MCP secret.                                                       |
| `401 Unauthorized` (plain text), or "could not be reached" | Pangolin's auth blocked the request before it reached the server.                  |
| Image pull fails with `denied` / `unauthorized`            | The GitHub package is private. Make it public (see step 1).                        |
| Tools say it is not authenticated                          | Sign in again (step 2).                                                            |
