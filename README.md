# Microsoft To Do MCP

[![CI](https://github.com/Mcp20091/microsoft-todo-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/Mcp20091/microsoft-todo-mcp-server/actions/workflows/ci.yml)

A Model Context Protocol (MCP) server that enables AI assistants like Claude and Cursor to interact with Microsoft To Do via the Microsoft Graph API. This service provides comprehensive task management capabilities through a secure OAuth 2.0 authentication flow.

## Features

- **27 MCP Tools**: Comprehensive Microsoft To Do management across authentication, lists, tasks, linked resources, attachments, checklist items, delta sync, and utility workflows
- **Seamless Authentication**: Automatic token refresh with zero manual intervention
- **OAuth 2.0 Authentication**: Secure authentication with automatic token refresh
- **Microsoft Graph API Integration**: Direct integration with Microsoft's official API
- **Multi-tenant Support**: Works with personal, work, and school Microsoft accounts
- **TypeScript**: Fully typed for reliability and developer experience
- **ESM Modules**: Modern JavaScript module system

## Prerequisites

- Node.js 18 or higher (tested with Node.js 18.x, 20.x, and 22.x)
- npm or pnpm
- A Microsoft account (personal, work, or school)
- Azure App Registration (see setup below)

## Installation

### Recommended: Clone and Run This Fork Locally

```bash
git clone https://github.com/Mcp20091/microsoft-todo-mcp-server.git
cd microsoft-todo-mcp-server
pnpm install
pnpm run build
```

This fork is documented primarily for local use from your own checkout. The setup helpers generate MCP config that points at your local built `dist/cli.js`.

## Azure App Registration

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to "App registrations" and create a new registration
3. Name your application (e.g., "To Do MCP")
4. For "Supported account types", select one of the following based on your needs:
   - **Accounts in this organizational directory only (Single tenant)** - For use within a single organization
   - **Accounts in any organizational directory (Any Azure AD directory - Multitenant)** - For use across multiple organizations
   - **Accounts in any organizational directory and personal Microsoft accounts** - For both work accounts and personal accounts
5. Set the Redirect URI to `http://localhost:3000/callback`
6. After creating the app, go to "Certificates & secrets" and create a new client secret
7. Go to "API permissions" and add the following permissions:
   - Microsoft Graph > Delegated permissions:
     - Tasks.Read
     - Tasks.ReadWrite
     - User.Read
8. Click "Grant admin consent" for these permissions

## Configuration

### Environment Setup

Create a `.env` file in the project root (required for authentication):

```env
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
TENANT_ID=your_tenant_setting
REDIRECT_URI=http://localhost:3000/callback
```

### TENANT_ID Options

- `organizations` - For multi-tenant organizational accounts (default if not specified)
- `consumers` - For personal Microsoft accounts only
- `common` - For both organizational and personal accounts
- `your-specific-tenant-id` - For single-tenant configurations

**Examples:**

```env
# For multi-tenant organizational accounts (default)
TENANT_ID=organizations

# For personal Microsoft accounts
TENANT_ID=consumers

# For both organizational and personal accounts
TENANT_ID=common

# For a specific organization tenant
TENANT_ID=00000000-0000-0000-0000-000000000000
```

### Token Storage

The server stores authentication tokens with automatic refresh 5 minutes before expiration.

Default token location:
- Windows: `%APPDATA%\microsoft-todo-mcp\tokens.json`
- macOS/Linux: `~/.config/microsoft-todo-mcp/tokens.json`

`pnpm run create-config` resolves tokens in this order:
1. explicit CLI argument
2. `MSTODO_TOKEN_FILE`
3. `./tokens.json` in the current directory, if present
4. the default per-user config path above

You can override the token file location:

```bash
# Using environment variable
export MSTODO_TOKEN_FILE=/path/to/custom/tokens.json

# Or pass tokens directly
export MS_TODO_ACCESS_TOKEN=your_access_token
export MS_TODO_REFRESH_TOKEN=your_refresh_token
```

## Usage

### Complete Setup Workflow

For this fork, the expected workflow is to run from a local clone:

```bash
pnpm install
pnpm run build
pnpm run auth
pnpm run create-config
```

Authentication opens a browser window and creates a token file. Configuration generation creates an `mcp.json` file from those tokens.

#### Step 1: Authenticate with Microsoft

```bash
pnpm run auth
```

#### Step 2: Create MCP Configuration

```bash
pnpm run create-config
```

This creates an `mcp.json` file that points to your local built `dist/cli.js`.

#### Step 3: Configure Your AI Assistant

**For Claude Desktop:**

Add to your configuration file:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoftTodo": {
      "command": "node",
      "args": ["/ABSOLUTE/PATH/TO/microsoft-todo-mcp-server/dist/cli.js"]
    }
  }
}
```

**For Cursor:**

```bash
# Copy to Cursor's global configuration
cp mcp.json ~/.cursor/mcp-servers.json
```

### Available Scripts

```bash
# Development & Building
pnpm run build        # Build TypeScript to JavaScript
pnpm run dev          # Build and run CLI in one command

# Running the Server
pnpm start            # Run MCP server directly
pnpm run cli          # Run MCP server via CLI wrapper

# Authentication & Configuration
pnpm run auth         # Start OAuth authentication server
pnpm run setup        # Run setup helper
pnpm run create-config # Generate mcp.json from the resolved token source

# Code Quality
pnpm run format       # Format code with Prettier
pnpm run format:check # Check code formatting
pnpm run lint         # Run linting checks
pnpm run typecheck    # TypeScript type checking
```

## MCP Tools

The server provides 27 tools for Microsoft To Do management.

### Authentication

- **`auth-status`** - Check authentication status, token expiration, and account type

### Task Lists (Top-level Containers)

- **`get-task-lists`** - Retrieve all task lists with metadata (default, shared, etc.)
- **`get-task-list`** - Retrieve a single task list by ID
- **`get-task-lists-delta`** - Track changes to task lists using Microsoft Graph delta queries
- **`get-task-lists-organized`** - Group lists into a more human-friendly organized view
- **`create-task-list`** - Create a new task list
- **`update-task-list`** - Rename an existing task list
- **`delete-task-list`** - Delete a task list and all its contents

### Tasks (Main Todo Items)

- **`get-tasks`** - Get tasks from a list with filtering, sorting, and pagination
  - Supports OData query parameters: `$filter`, `$select`, `$orderby`, `$top`, `$skip`, `$count`
- **`get-task`** - Retrieve a single task by ID
- **`get-tasks-delta`** - Track changes to tasks in a list using Microsoft Graph delta queries
- **`create-task`** - Create a new task with full property support
  - Title, description, due date, start date, completed date, importance, reminders, recurrence, status, categories, linked resources
- **`update-task`** - Update any task properties, including clearing due dates, reminders, start dates, and recurrence
- **`delete-task`** - Delete a task and all its checklist items

### Linked Resources

- **`get-linked-resources`** - List linked resources associated with a task
- **`create-linked-resource`** - Create a linked resource for a task

### Attachments

- **`get-attachments`** - List file attachments for a task
- **`get-attachment`** - Retrieve a single file attachment
- **`create-attachment`** - Add a small file attachment to a task
- **`create-attachment-upload-session`** - Create an upload session for large file attachments
- **`delete-attachment`** - Remove an attachment from a task

### Checklist Items (Subtasks)

- **`get-checklist-items`** - Get subtasks for a specific task
- **`create-checklist-item`** - Add a new subtask to a task
- **`update-checklist-item`** - Update subtask text, completion status, and checklist timestamps
- **`delete-checklist-item`** - Remove a specific subtask

### Utilities

- **`archive-completed-tasks`** - Move completed tasks older than a specified age to another list
- **`test-graph-api-exploration`** - Explore Microsoft Graph To Do endpoints and response shapes for troubleshooting

## Architecture

### Project Structure

- **MCP Server** (`src/todo-index.ts`) - Core server implementing the MCP protocol
- **CLI Wrapper** (`src/cli.ts`) - Executable entry point with token management
- **Auth Server** (`src/auth-server.ts`) - Express server for OAuth 2.0 flow
- **Config Generator** (`src/create-mcp-config.ts`) - Helper to create MCP configurations

### Technical Details

- **Microsoft Graph API**: Uses v1.0 endpoints
- **Authentication**: MSAL (Microsoft Authentication Library) with PKCE flow
- **Token Management**: Automatic refresh 5 minutes before expiration
- **Build System**: tsup targeting Node.js 18 with ESM output
- **Module System**: ESM (ECMAScript modules)

## Limitations & Known Issues

### Personal Microsoft Accounts

- **MailboxNotEnabledForRESTAPI Error**: Personal Microsoft accounts (outlook.com, hotmail.com, live.com) have limited access to the To Do API through Microsoft Graph
- This is a Microsoft service limitation, not an issue with this application
- Work/school accounts have full API access

### API Limitations

- Rate limits apply according to Microsoft's policies
- Some features may be unavailable for personal accounts
- Shared lists have limited functionality

## Troubleshooting

### Authentication Issues

**Token acquisition failures**

- Verify `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` in your `.env` file
- Ensure redirect URI matches exactly: `http://localhost:3000/callback`
- Check Azure App permissions are granted with admin consent

**Permission issues**

- Ensure all required Graph API permissions are added and consented
- For organizational accounts, admin consent may be required

### Account Type Configuration

**Work/School Accounts**

```env
TENANT_ID=organizations  # Multi-tenant
# Or use your specific tenant ID
```

**Personal Accounts**

```env
TENANT_ID=consumers  # Personal only
# Or TENANT_ID=common for both types
```

### Debugging

**Check authentication status:**

```bash
# Using the MCP tool
# In your AI assistant: "Check auth status"
```

**Inspect token storage:**

Windows PowerShell:

```powershell
$tokenPath = Join-Path $env:APPDATA 'microsoft-todo-mcp\tokens.json'
Get-Content $tokenPath
```

macOS/Linux:

```bash
cat ~/.config/microsoft-todo-mcp/tokens.json
```

**Enable verbose logging:**

```bash
# The server logs to stderr for debugging
pnpm run cli 2> debug.log
```

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Run `pnpm run lint` and `pnpm run typecheck` before submitting
4. Submit a pull request

## License

MIT License - See [LICENSE](LICENSE) file for details

## Acknowledgments

- Fork of [@jhirono/todomcp](https://github.com/jhirono/todomcp)
- Built on the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/sdk)
- Uses [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)

## Support

- [GitHub Issues](https://github.com/Mcp20091/microsoft-todo-mcp-server/issues)
