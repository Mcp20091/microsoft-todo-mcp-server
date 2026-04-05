#!/usr/bin/env node

/**
 * Ollama Agent for Microsoft Todo MCP Server
 *
 * An interactive chat agent that connects to the Microsoft Todo MCP server
 * and uses a local Ollama model to process natural language requests.
 *
 * Usage:
 *   pnpm run ollama                          # Use default model (llama3.2)
 *   OLLAMA_MODEL=qwen2.5 pnpm run ollama     # Use a specific model
 *   OLLAMA_HOST=http://localhost:11434 pnpm run ollama
 *
 * Requirements:
 *   - Ollama installed and running (https://ollama.com)
 *   - A model with tool/function calling support pulled in Ollama
 *     e.g.: ollama pull llama3.2
 *   - Microsoft Todo tokens configured (tokens.json or env vars)
 */

import { Client } from "@modelcontextprotocol/sdk/client/index.js"
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js"
import { Ollama, type Message, type Tool } from "ollama"
import * as readline from "readline"
import path from "path"
import { fileURLToPath } from "url"

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const OLLAMA_HOST = process.env.OLLAMA_HOST || "http://localhost:11434"
const OLLAMA_MODEL = process.env.OLLAMA_MODEL || "llama3.2"

// Resolve path to the CLI entry point (sibling file after build)
const CLI_PATH = path.join(__dirname, "cli.js")

async function createMcpClient(): Promise<Client> {
  const transport = new StdioClientTransport({
    command: "node",
    args: [CLI_PATH],
    env: { ...process.env } as Record<string, string>,
  })

  const client = new Client({ name: "ollama-agent", version: "1.0.0" }, { capabilities: {} })
  await client.connect(transport)
  return client
}

type McpTool = {
  name: string
  description?: string
  inputSchema: {
    type: string
    properties?: Record<string, unknown>
    required?: string[]
  }
}

function convertMcpToolToOllama(mcpTool: McpTool): Tool {
  return {
    type: "function",
    function: {
      name: mcpTool.name,
      description: mcpTool.description || "",
      parameters: {
        type: mcpTool.inputSchema.type || "object",
        properties: (mcpTool.inputSchema.properties as Record<string, { type: string; description?: string }>) || {},
        required: mcpTool.inputSchema.required || [],
      },
    },
  }
}

async function callMcpTool(client: Client, toolName: string, args: Record<string, unknown>): Promise<string> {
  try {
    const result = await client.callTool({ name: toolName, arguments: args })
    if (result.content && Array.isArray(result.content)) {
      return result.content
        .map((item: { type: string; text?: string }) => (item.type === "text" ? (item.text ?? "") : ""))
        .join("\n")
    }
    return JSON.stringify(result)
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error)
    return `Error calling tool ${toolName}: ${msg}`
  }
}

async function chat(client: Client, ollamaTools: Tool[], history: Message[], userMessage: string): Promise<string> {
  const ollama = new Ollama({ host: OLLAMA_HOST })

  history.push({ role: "user", content: userMessage })

  let response = await ollama.chat({
    model: OLLAMA_MODEL,
    messages: history,
    tools: ollamaTools,
  })

  // Agentic loop: keep processing tool calls until the model gives a final text response
  while (response.message.tool_calls && response.message.tool_calls.length > 0) {
    // Add the assistant's tool-calling message to history
    history.push(response.message)

    for (const toolCall of response.message.tool_calls) {
      const toolName = toolCall.function.name
      const toolArgs = toolCall.function.arguments as Record<string, unknown>

      process.stderr.write(`[tool] calling ${toolName}(${JSON.stringify(toolArgs)})\n`)

      const toolResult = await callMcpTool(client, toolName, toolArgs)

      process.stderr.write(`[tool] result: ${toolResult.slice(0, 200)}${toolResult.length > 200 ? "..." : ""}\n`)

      history.push({
        role: "tool",
        content: toolResult,
      })
    }

    // Get next response from the model
    response = await ollama.chat({
      model: OLLAMA_MODEL,
      messages: history,
      tools: ollamaTools,
    })
  }

  const assistantMessage = response.message.content || ""
  history.push({ role: "assistant", content: assistantMessage })
  return assistantMessage
}

async function main(): Promise<void> {
  console.log(`Microsoft Todo Ollama Agent`)
  console.log(`Model: ${OLLAMA_MODEL}  Host: ${OLLAMA_HOST}`)
  console.log(`Connecting to MCP server...`)

  let client: Client
  try {
    client = await createMcpClient()
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error)
    console.error(`Failed to connect to MCP server: ${msg}`)
    console.error(`Make sure tokens.json is configured. Run: pnpm run setup`)
    process.exit(1)
  }

  // Fetch available tools from MCP server
  const toolsResult = await client.listTools()
  const mcpTools = toolsResult.tools as McpTool[]
  const ollamaTools = mcpTools.map(convertMcpToolToOllama)
  console.log(`Loaded ${ollamaTools.length} tools: ${mcpTools.map((t) => t.name).join(", ")}`)

  // Verify Ollama is reachable
  const ollama = new Ollama({ host: OLLAMA_HOST })
  try {
    await ollama.show({ model: OLLAMA_MODEL })
  } catch {
    console.error(`\nCould not reach Ollama or find model "${OLLAMA_MODEL}".`)
    console.error(`Make sure Ollama is running and the model is pulled:`)
    console.error(`  ollama pull ${OLLAMA_MODEL}`)
    console.error(`\nTo use a different model: OLLAMA_MODEL=<name> pnpm run ollama`)
    await client.close()
    process.exit(1)
  }

  const systemPrompt = `You are a helpful assistant that manages Microsoft To Do tasks and lists.
You have access to tools to create, read, update, and delete tasks and task lists.
Be concise and confirm actions after completing them.`

  const history: Message[] = [{ role: "system", content: systemPrompt }]

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout })
  console.log(`\nReady. Type your request (or "exit" to quit).\n`)

  const prompt = (): void => {
    rl.question("You: ", async (input: string) => {
      const userInput = input.trim()
      if (!userInput) {
        prompt()
        return
      }
      if (userInput.toLowerCase() === "exit" || userInput.toLowerCase() === "quit") {
        console.log("Goodbye!")
        await client.close()
        rl.close()
        process.exit(0)
      }

      try {
        const reply = await chat(client, ollamaTools, history, userInput)
        console.log(`\nAssistant: ${reply}\n`)
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error)
        console.error(`Error: ${msg}`)
      }

      prompt()
    })
  }

  prompt()
}

main().catch((error) => {
  console.error("Fatal error:", error)
  process.exit(1)
})
