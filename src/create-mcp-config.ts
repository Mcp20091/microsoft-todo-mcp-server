#!/usr/bin/env node

import fs from "fs"
import path from "path"
import { homedir } from "os"
import { fileURLToPath } from "url"

// Get the directory path for the current module
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

function getDefaultConfigTokenPath(): string {
  const configDir =
    process.platform === "win32"
      ? path.join(process.env.APPDATA || path.join(homedir(), "AppData", "Roaming"), "microsoft-todo-mcp")
      : path.join(homedir(), ".config", "microsoft-todo-mcp")

  return path.join(configDir, "tokens.json")
}

function resolveTokenPath(): string {
  if (process.argv[2]) return process.argv[2]
  if (process.env.MSTODO_TOKEN_FILE) return process.env.MSTODO_TOKEN_FILE

  const localTokenPath = path.join(process.cwd(), "tokens.json")
  if (fs.existsSync(localTokenPath)) return localTokenPath

  return getDefaultConfigTokenPath()
}

// Define paths
const tokenPath = resolveTokenPath()
const outputPath = process.argv[3] || path.join(process.cwd(), "mcp.json")

console.log(`Reading tokens from: ${tokenPath}`)
console.log(`Writing config to: ${outputPath}`)

try {
  // Read the tokens
  const tokenData = JSON.parse(fs.readFileSync(tokenPath, "utf8"))

  // Create the MCP config - only include the actual tokens
  const mcpConfig = {
    mcpServers: {
      microsoftTodo: {
        command: "npx",
        args: ["--yes", "microsoft-todo-mcp-server"],
        env: {
          MS_TODO_ACCESS_TOKEN: tokenData.accessToken,
          MS_TODO_REFRESH_TOKEN: tokenData.refreshToken,
        },
      },
    },
  }

  // Write the config
  fs.writeFileSync(outputPath, JSON.stringify(mcpConfig, null, 2), "utf8")

  console.log("MCP configuration file created successfully!")
  console.log("You can now use the service with Claude or Cursor by referencing this mcp.json file.")
} catch (error) {
  // Fix potential TypeScript error with unknown error type
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error creating MCP config:", errorMessage)
  process.exit(1)
}
