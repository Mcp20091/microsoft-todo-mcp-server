#!/usr/bin/env node

import fs from "fs"
import path from "path"
import { fileURLToPath } from "url"

// Get the directory path for the current module
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

function resolveTokenPath(): string {
  if (process.argv[2]) return process.argv[2]
  if (process.env.MSTODO_TOKEN_FILE) return process.env.MSTODO_TOKEN_FILE

  return path.join(process.cwd(), "tokens.json")
}

// Define paths
const tokenPath = resolveTokenPath()
const outputPath = process.argv[3] || path.join(process.cwd(), "mcp.json")
const cliPath = path.resolve(process.cwd(), "dist", "cli.js")

console.log(`Reading tokens from: ${tokenPath}`)
console.log(`Writing config to: ${outputPath}`)
console.log(`Using CLI path: ${cliPath}`)

try {
  // Read the tokens
  const tokenData = JSON.parse(fs.readFileSync(tokenPath, "utf8"))

  // Create MCP config for this local fork checkout.
  const mcpConfig = {
    mcpServers: {
      microsoftTodo: {
        command: "node",
        args: [cliPath],
        env: {
          MSTODO_TOKEN_FILE: tokenPath,
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
