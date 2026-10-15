#!/usr/bin/env node

// Serves the MCP server over Streamable HTTP so several devices can share one
// instance (Docker, TrueNAS app, Proxmox LXC). Every request gets a fresh,
// stateless server; the Microsoft tokens live in one token file.

import { timingSafeEqual } from "crypto"
import express, { type NextFunction, type Request, type Response } from "express"
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js"
import { createServer } from "./todo-index.js"
import { tokenManager } from "./token-manager.js"

const port = Number(process.env.PORT || 8787)
const host = process.env.HOST || "127.0.0.1"
const authToken = process.env.MCP_AUTH_TOKEN || ""

if (authToken.length < 16) {
  console.error("Set MCP_AUTH_TOKEN to a random secret of at least 16 characters (e.g. `openssl rand -hex 32`).")
  process.exit(1)
}

if (process.env.MSTODO_TOKEN_FILE) {
  tokenManager.configure({ tokenFilePath: process.env.MSTODO_TOKEN_FILE })
}
console.error(`Token file: ${tokenManager.getTokenFilePath()}`)

function matches(given: string | undefined, expected: string): boolean {
  const a = Buffer.from(given ?? "")
  const b = Buffer.from(expected)
  return a.length === b.length && timingSafeEqual(a, b)
}

// Accepts the secret as a bearer header, or as the last path segment
// (/mcp/<secret>) for clients that can't send headers, like claude.ai connectors.
function requireAuth(req: Request, res: Response, next: NextFunction): void {
  const pathToken = typeof req.params.token === "string" ? req.params.token : undefined
  if (matches(req.headers.authorization, `Bearer ${authToken}`) || matches(pathToken, authToken)) {
    next()
    return
  }
  res
    .status(401)
    .set("WWW-Authenticate", "Bearer")
    .json({
      jsonrpc: "2.0",
      error: { code: -32001, message: "Unauthorized" },
      id: null,
    })
}

const app = express()
app.use(express.json({ limit: "4mb" }))

app.get("/healthz", (_req, res) => {
  res.json({ ok: true })
})

app.post(["/mcp", "/mcp/:token"], requireAuth, async (req, res) => {
  const server = createServer()
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
    enableJsonResponse: true,
  })
  res.on("close", () => {
    transport.close()
    server.close()
  })

  try {
    await server.connect(transport)
    await transport.handleRequest(req, res, req.body)
  } catch (error) {
    console.error("Error handling MCP request:", error)
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        error: { code: -32603, message: "Internal server error" },
        id: null,
      })
    }
  }
})

// Stateless server: no SSE stream to open and no session to end.
app.all(["/mcp", "/mcp/:token"], requireAuth, (_req, res) => {
  res
    .status(405)
    .set("Allow", "POST")
    .json({
      jsonrpc: "2.0",
      error: { code: -32000, message: "Method not allowed" },
      id: null,
    })
})

app.listen(port, host, () => {
  console.error(`Microsoft To Do MCP listening on http://${host}:${port}/mcp`)
})
