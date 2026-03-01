import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { z } from "zod"
import { readFileSync, writeFileSync, existsSync } from "fs"
import { join } from "path"
import dotenv from "dotenv"
import { tokenManager } from "./token-manager.js"

// Load environment variables
dotenv.config()

// Log the current working directory
console.error("Current working directory:", process.cwd())

// Microsoft Graph API endpoints
const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const USER_AGENT = "microsoft-todo-mcp-server/1.0"
const LOG_PREVIEW_LIMIT = 4000

// Create server instance
const server = new McpServer({
  name: "mstodo",
  version: "1.0.0",
})

type GraphRequestErrorInfo = {
  url: string
  method: string
  status?: number
  responseBody?: string
  requestBody?: string
  requestId?: string
  clientRequestId?: string
  message: string
}

let lastGraphRequestError: GraphRequestErrorInfo | null = null

// Helper function for making Microsoft Graph API requests
async function makeGraphRequest<T>(url: string, token: string, method = "GET", body?: any): Promise<T | null> {
  const headers = {
    "User-Agent": USER_AGENT,
    Accept: "application/json",
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  }
  const serializedBody = body && (method === "POST" || method === "PATCH") ? JSON.stringify(body) : undefined

  try {
    lastGraphRequestError = null

    const options: RequestInit = {
      method,
      headers,
    }

    if (serializedBody) {
      options.body = serializedBody
    }

    console.error(`Making request to: ${url}`)
    console.error(
      `Request options: ${JSON.stringify({
        method,
        headers: {
          ...headers,
          Authorization: "Bearer [REDACTED]",
        },
        body: serializedBody ? formatBodyForLog(body) : undefined,
      })}`,
    )

    let response = await fetch(url, options)

    // If we get a 401, try to refresh the token and retry once
    if (response.status === 401) {
      console.error("Got 401, attempting token refresh...")
      const newToken = await getAccessToken() // This will trigger refresh
      if (newToken && newToken !== token) {
        // Retry with new token
        headers.Authorization = `Bearer ${newToken}`
        response = await fetch(url, { ...options, headers })
      }
    }

    if (!response.ok) {
      const errorText = await response.text()
      lastGraphRequestError = extractGraphErrorInfo(url, method, serializedBody, response.status, errorText)

      console.error(`HTTP error! status: ${response.status}, body: ${errorText}`)
      console.error(`Response headers: ${JSON.stringify(formatHeadersForLog(response.headers))}`)
      if (serializedBody) {
        console.error(`Request body: ${formatBodyForLog(body)}`)
      }
      console.error(`Request URL: ${url}`)

      // Check for the specific MailboxNotEnabledForRESTAPI error
      if (errorText.includes("MailboxNotEnabledForRESTAPI")) {
        console.error(`
=================================================================
ERROR: MailboxNotEnabledForRESTAPI

The Microsoft To Do API is not available for personal Microsoft accounts 
(outlook.com, hotmail.com, live.com, etc.) through the Graph API.

This is a limitation of the Microsoft Graph API, not an authentication issue.
Microsoft only allows To Do API access for Microsoft 365 business accounts.

You can still use Microsoft To Do through the web interface or mobile apps,
but API access is restricted for personal accounts.
=================================================================
        `)

        throw new Error(
          "Microsoft To Do API is not available for personal Microsoft accounts. See console for details.",
        )
      }

      throw new Error(`HTTP error! status: ${response.status}, body: ${errorText}`)
    }

    if (response.status === 204) {
      console.error("Response received: 204 No Content")
      return null
    }

    const responseText = await response.text()
    if (!responseText.trim()) {
      console.error("Response received: empty body")
      return null
    }

    const data = JSON.parse(responseText)
    const responsePreview = formatBodyForLog(data)
    console.error(`Response headers: ${JSON.stringify(formatHeadersForLog(response.headers))}`)
    console.error(`Response received: ${responsePreview}`)
    return data as T
  } catch (error) {
    console.error("Error making Graph API request:", error)
    if (!lastGraphRequestError && error instanceof Error) {
      lastGraphRequestError = {
        url,
        method,
        requestBody: serializedBody,
        message: error.message,
      }
    }
    return null
  }
}

// Authentication helper using delegated flow with token manager
async function getAccessToken(): Promise<string | null> {
  try {
    console.error("getAccessToken called")

    // Use the token manager to get tokens (handles all sources and refresh)
    const tokens = await tokenManager.getTokens()

    if (tokens) {
      console.error(`Successfully retrieved valid token`)
      return tokens.accessToken
    }

    console.error("No valid tokens available")
    return null
  } catch (error) {
    console.error("Error getting access token:", error)
    return null
  }
}

// Server configuration type
interface ServerConfig {
  accessToken?: string
  refreshToken?: string
  tokenFilePath?: string
}

// Function to check if the account is a personal Microsoft account
async function isPersonalMicrosoftAccount(): Promise<boolean> {
  try {
    const token = await getAccessToken()
    if (!token) return false

    // Make a request to get user info
    const url = `${MS_GRAPH_BASE}/me`
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    })

    if (!response.ok) {
      console.error(`Error getting user info: ${response.status}`)
      return false
    }

    const userData = await response.json()
    const email = userData.mail || userData.userPrincipalName || ""

    // Check if the email domain indicates a personal account
    const personalDomains = ["outlook.com", "hotmail.com", "live.com", "msn.com", "passport.com"]
    const domain = email.split("@")[1]?.toLowerCase()

    if (domain && personalDomains.some((d) => domain.includes(d))) {
      console.error(`
=================================================================
WARNING: Personal Microsoft Account Detected

Your Microsoft account (${email}) appears to be a personal account.
Microsoft To Do API access is typically not available for personal accounts
through the Microsoft Graph API, only for Microsoft 365 business accounts.

You may encounter the "MailboxNotEnabledForRESTAPI" error when trying to
access To Do lists or tasks. This is a limitation of the Microsoft Graph API,
not an issue with your authentication or this application.

You can still use Microsoft To Do through the web interface or mobile apps,
but API access is restricted for personal accounts.
=================================================================
      `)
      return true
    }

    return false
  } catch (error) {
    console.error("Error checking account type:", error)
    return false
  }
}

// Server tool to check authentication status
server.tool(
  "auth-status",
  "Check if you're authenticated with Microsoft Graph API. Shows current token status and expiration time, and indicates if the token needs to be refreshed.",
  {},
  async () => {
    const tokens = await tokenManager.getTokens()

    if (!tokens) {
      return {
        content: [
          {
            type: "text",
            text: "Not authenticated. Please run 'npx mstodo-setup' or 'pnpm run setup' to authenticate with Microsoft.",
          },
        ],
      }
    }

    const isExpired = Date.now() > tokens.expiresAt
    const expiryTime = new Date(tokens.expiresAt).toLocaleString()

    // Check if it's a personal account
    const isPersonal = await isPersonalMicrosoftAccount()
    let accountMessage = ""

    if (isPersonal) {
      accountMessage =
        "\n\n⚠️ WARNING: You are using a personal Microsoft account. " +
        "Microsoft To Do API access is typically not available for personal accounts " +
        "through the Microsoft Graph API. You may encounter 'MailboxNotEnabledForRESTAPI' errors. " +
        "This is a Microsoft limitation, not an authentication issue."
    }

    if (isExpired) {
      return {
        content: [
          {
            type: "text",
            text: `Authentication expired at ${expiryTime}. Will attempt to refresh when you call any API.${accountMessage}`,
          },
        ],
      }
    } else {
      return {
        content: [
          {
            type: "text",
            text: `Authenticated. Token expires at ${expiryTime}.${accountMessage}`,
          },
        ],
      }
    }
  },
)

interface TaskList {
  id: string
  displayName: string
  isOwner?: boolean
  isShared?: boolean
  wellknownListName?: string // 'none', 'defaultList', 'flaggedEmails', 'unknownFutureValue'
}

interface DateTimeTimeZone {
  dateTime: string
  timeZone: string
}

interface RecurrencePattern {
  type: string
  interval: number
  month?: number
  dayOfMonth?: number
  daysOfWeek?: string[]
  firstDayOfWeek?: string
  index?: string
}

interface RecurrenceRange {
  type: string
  startDate: string
  endDate?: string
  recurrenceTimeZone?: string
  numberOfOccurrences?: number
}

interface PatternedRecurrence {
  pattern: RecurrencePattern
  range: RecurrenceRange
}

interface LinkedResource {
  id?: string
  webUrl?: string
  applicationName?: string
  displayName?: string
  externalId?: string
}

interface TaskFileAttachment {
  id: string
  name: string
  contentType?: string
  size?: number
  lastModifiedDateTime?: string
  contentBytes?: string
}

interface DeltaResponse<T> {
  value: T[]
  "@odata.count"?: number
  "@odata.nextLink"?: string
  "@odata.deltaLink"?: string
}

interface Task {
  id: string
  title: string
  status: string
  importance: string
  dueDateTime?: DateTimeTimeZone
  startDateTime?: DateTimeTimeZone
  completedDateTime?: DateTimeTimeZone
  reminderDateTime?: DateTimeTimeZone
  isReminderOn?: boolean
  recurrence?: PatternedRecurrence | null
  hasAttachments?: boolean
  createdDateTime?: string
  lastModifiedDateTime?: string
  bodyLastModifiedDateTime?: string
  body?: {
    content: string
    contentType: string
  }
  categories?: string[]
  linkedResources?: LinkedResource[]
}

interface ChecklistItem {
  id: string
  displayName: string
  isChecked: boolean
  checkedDateTime?: string
  createdDateTime?: string
}

interface UploadSession {
  uploadUrl: string
  expirationDateTime: string
  nextExpectedRanges: string[]
}

const recurrenceSchema = z.object({
  pattern: z.object({
    type: z.string().describe("Recurrence type such as daily, weekly, absoluteMonthly, relativeMonthly"),
    interval: z.number().int().min(1).describe("Repeat interval"),
    month: z.number().int().min(1).max(12).optional(),
    dayOfMonth: z.number().int().min(1).max(31).optional(),
    daysOfWeek: z.array(z.string()).optional(),
    firstDayOfWeek: z.string().optional(),
    index: z.string().optional(),
  }),
  range: z.object({
    type: z.string().describe("Range type such as noEnd, endDate, or numbered"),
    startDate: z.string().describe("Start date in YYYY-MM-DD format"),
    endDate: z.string().optional().describe("End date in YYYY-MM-DD format"),
    recurrenceTimeZone: z.string().optional(),
    numberOfOccurrences: z.number().int().min(1).optional(),
  }),
})

type RecurrenceInput = z.infer<typeof recurrenceSchema>

const linkedResourceSchema = z.object({
  webUrl: z.string().optional().describe("Deep link to the linked item"),
  applicationName: z.string().optional().describe("Source application name"),
  displayName: z.string().optional().describe("Display title for the linked item"),
  externalId: z.string().optional().describe("External identifier from the source system"),
})

function buildDateTimeTimeZone(dateTime: string): DateTimeTimeZone {
  return {
    dateTime,
    timeZone: "UTC",
  }
}

function buildRecurrencePayload(recurrence: RecurrenceInput): PatternedRecurrence {
  const pattern: PatternedRecurrence["pattern"] = {
    type: recurrence.pattern.type,
    interval: recurrence.pattern.interval,
  }

  if (recurrence.pattern.month !== undefined) pattern.month = recurrence.pattern.month
  if (recurrence.pattern.dayOfMonth !== undefined) pattern.dayOfMonth = recurrence.pattern.dayOfMonth
  if (recurrence.pattern.daysOfWeek !== undefined) pattern.daysOfWeek = recurrence.pattern.daysOfWeek
  if (recurrence.pattern.firstDayOfWeek !== undefined) pattern.firstDayOfWeek = recurrence.pattern.firstDayOfWeek
  if (recurrence.pattern.index !== undefined) pattern.index = recurrence.pattern.index

  const range: PatternedRecurrence["range"] = {
    type: recurrence.range.type,
    startDate: recurrence.range.startDate,
  }

  if (recurrence.range.endDate !== undefined) range.endDate = recurrence.range.endDate
  if (recurrence.range.numberOfOccurrences !== undefined) range.numberOfOccurrences = recurrence.range.numberOfOccurrences
  if (recurrence.range.recurrenceTimeZone !== undefined) range.recurrenceTimeZone = recurrence.range.recurrenceTimeZone

  return {
    pattern,
    range,
  }
}

function buildRecurrencePatchPayload(recurrence: RecurrenceInput): PatternedRecurrence {
  return {
    pattern: buildRecurrencePayload(recurrence).pattern,
    range: {} as PatternedRecurrence["range"],
  }
}

function formatBodyForLog(body: unknown): string {
  if (body === undefined) return ""

  try {
    const serialized = JSON.stringify(body)
    return serialized.length > LOG_PREVIEW_LIMIT
      ? `${serialized.slice(0, LOG_PREVIEW_LIMIT)}... [truncated ${serialized.length - LOG_PREVIEW_LIMIT} chars]`
      : serialized
  } catch {
    return "[unserializable body]"
  }
}

function formatHeadersForLog(headers: Headers): Record<string, string> {
  return Object.fromEntries(Array.from(headers.entries()))
}

function extractGraphErrorInfo(
  url: string,
  method: string,
  requestBody: string | undefined,
  status: number,
  responseBody: string,
): GraphRequestErrorInfo {
  let requestId: string | undefined
  let clientRequestId: string | undefined

  try {
    const parsed = JSON.parse(responseBody)
    requestId = parsed?.error?.innerError?.["request-id"]
    clientRequestId = parsed?.error?.innerError?.["client-request-id"]
  } catch {}

  return {
    url,
    method,
    status,
    responseBody,
    requestBody,
    requestId,
    clientRequestId,
    message: `HTTP error! status: ${status}, body: ${responseBody}`,
  }
}

function isRecurrencePatchDateError(info: GraphRequestErrorInfo | null): boolean {
  if (!info) return false

  return (
    info.method === "PATCH" &&
    info.status === 400 &&
    info.responseBody?.includes("recurrence.range.startDate") === true &&
    info.responseBody?.includes("Microsoft.OData.Edm.Date") === true
  )
}

function isFutureOrCurrentDateTime(value: DateTimeTimeZone): boolean {
  if (!value.dateTime) return false

  const dueDateOnly = value.dateTime.slice(0, 10)
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDateOnly)) return false

  const today = new Date()
  const todayDateOnly = `${today.getUTCFullYear()}-${String(today.getUTCMonth() + 1).padStart(2, "0")}-${String(
    today.getUTCDate(),
  ).padStart(2, "0")}`

  return dueDateOnly >= todayDateOnly
}

function formatDateTime(value?: DateTimeTimeZone | null): string | null {
  if (!value?.dateTime) return null

  const date = new Date(value.dateTime)
  if (Number.isNaN(date.getTime())) {
    return `${value.dateTime} (${value.timeZone})`
  }

  return `${date.toLocaleString()} (${value.timeZone})`
}

function formatRecurrence(recurrence?: PatternedRecurrence | null): string | null {
  if (!recurrence) return null

  const details = [`${recurrence.pattern.type} every ${recurrence.pattern.interval}`]
  const hasSentinelEndDate = recurrence.range.endDate === "0001-01-01"

  if (recurrence.pattern.daysOfWeek?.length) {
    details.push(`on ${recurrence.pattern.daysOfWeek.join(", ")}`)
  }

  if (recurrence.pattern.dayOfMonth) {
    details.push(`day ${recurrence.pattern.dayOfMonth}`)
  }

  if (recurrence.range.startDate) {
    details.push(`starting ${recurrence.range.startDate}`)
  }

  if (recurrence.range.type !== "noEnd" && recurrence.range.endDate && !hasSentinelEndDate) {
    details.push(`until ${recurrence.range.endDate}`)
  } else if (recurrence.range.numberOfOccurrences) {
    details.push(`for ${recurrence.range.numberOfOccurrences} occurrence(s)`)
  }

  return details.join(", ")
}

function formatTask(task: Task): string {
  let taskInfo = `ID: ${task.id}\nTitle: ${task.title}`

  if (task.status) {
    const status = task.status === "completed" ? "✓" : "○"
    taskInfo = `${status} ${taskInfo}`
  }

  const start = formatDateTime(task.startDateTime)
  const due = formatDateTime(task.dueDateTime)
  const reminder = formatDateTime(task.reminderDateTime)
  const completed = formatDateTime(task.completedDateTime)
  const recurrence = formatRecurrence(task.recurrence)

  if (start) taskInfo += `\nStart: ${start}`
  if (due) taskInfo += `\nDue: ${due}`
  if (reminder) taskInfo += `\nReminder: ${reminder}`
  if (completed) taskInfo += `\nCompleted: ${completed}`
  if (task.importance) taskInfo += `\nImportance: ${task.importance}`
  if (task.isReminderOn !== undefined) taskInfo += `\nReminder Enabled: ${task.isReminderOn ? "Yes" : "No"}`
  if (task.hasAttachments !== undefined) taskInfo += `\nHas Attachments: ${task.hasAttachments ? "Yes" : "No"}`
  if (recurrence) taskInfo += `\nRecurrence: ${recurrence}`
  if (task.categories && task.categories.length > 0) taskInfo += `\nCategories: ${task.categories.join(", ")}`

  if (task.linkedResources && task.linkedResources.length > 0) {
    const linkedSummary = task.linkedResources
      .map((resource) => resource.displayName || resource.applicationName || resource.webUrl || resource.id || "Linked item")
      .join(", ")
    taskInfo += `\nLinked Resources: ${linkedSummary}`
  }

  if (task.createdDateTime) taskInfo += `\nCreated: ${new Date(task.createdDateTime).toLocaleString()}`
  if (task.lastModifiedDateTime) taskInfo += `\nLast Modified: ${new Date(task.lastModifiedDateTime).toLocaleString()}`
  if (task.bodyLastModifiedDateTime) taskInfo += `\nBody Modified: ${new Date(task.bodyLastModifiedDateTime).toLocaleString()}`

  if (task.body && task.body.content && task.body.content.trim() !== "") {
    const previewLength = 120
    const contentPreview =
      task.body.content.length > previewLength ? task.body.content.substring(0, previewLength) + "..." : task.body.content
    taskInfo += `\nDescription: ${contentPreview}`
  }

  return `${taskInfo}\n---`
}

function isTaskFileAttachment(value: unknown): value is TaskFileAttachment {
  return Boolean(
    value &&
      typeof value === "object" &&
      "id" in value &&
      "name" in value &&
      typeof (value as { id?: unknown }).id === "string" &&
      typeof (value as { name?: unknown }).name === "string",
  )
}

// Register tools
server.tool(
  "get-task-lists",
  "Get all Microsoft Todo task lists (the top-level containers that organize your tasks). Shows list names, IDs, and indicates default or shared lists.",
  {},
  async () => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve task lists",
            },
          ],
        }
      }

      const lists = response.value || []
      if (lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No task lists found.",
            },
          ],
        }
      }

      const formattedLists = lists.map((list) => {
        // Add well-known list name if applicable
        let wellKnownInfo = ""
        if (list.wellknownListName && list.wellknownListName !== "none") {
          if (list.wellknownListName === "defaultList") {
            wellKnownInfo = " (Default Tasks List)"
          } else if (list.wellknownListName === "flaggedEmails") {
            wellKnownInfo = " (Flagged Emails)"
          }
        }

        // Add sharing info if applicable
        let sharingInfo = ""
        if (list.isShared) {
          sharingInfo = list.isOwner ? " (Shared by you)" : " (Shared with you)"
        }

        return `ID: ${list.id}\nName: ${list.displayName}${wellKnownInfo}${sharingInfo}\n---`
      })

      return {
        content: [
          {
            type: "text",
            text: `Your task lists:\n\n${formattedLists.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching task lists: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-task-list",
  "Get a single Microsoft Todo task list by ID.",
  {
    listId: z.string().describe("ID of the task list"),
    select: z.string().optional().describe("Comma-separated list of properties to include"),
  },
  async ({ listId, select }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const queryParams = new URLSearchParams()
      if (select) queryParams.append("$select", select)

      const queryString = queryParams.toString()
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}${queryString ? "?" + queryString : ""}`
      const list = await makeGraphRequest<TaskList>(url, token)

      if (!list) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve task list: ${listId}`,
            },
          ],
        }
      }

      const metadata = []
      if (list.wellknownListName && list.wellknownListName !== "none") metadata.push(`Type: ${list.wellknownListName}`)
      if (list.isShared !== undefined) metadata.push(`Shared: ${list.isShared ? "Yes" : "No"}`)
      if (list.isOwner !== undefined) metadata.push(`Owner: ${list.isOwner ? "Yes" : "No"}`)

      return {
        content: [
          {
            type: "text",
            text: `Task list details:\n\nID: ${list.id}\nName: ${list.displayName}${metadata.length ? `\n${metadata.join("\n")}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching task list: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-task-lists-delta",
  "Track changes to Microsoft Todo task lists using the Graph delta API.",
  {
    deltaUrl: z.string().optional().describe("Full @odata.nextLink or @odata.deltaLink URL from a previous delta response"),
    deltaToken: z.string().optional().describe("Delta token from a previous delta response"),
    skipToken: z.string().optional().describe("Skip token from a previous delta response"),
    select: z.string().optional().describe("Comma-separated list of properties to include on the initial request"),
    maxPageSize: z.number().int().min(1).optional().describe("Preferred maximum number of lists returned"),
  },
  async ({ deltaUrl, deltaToken, skipToken, select, maxPageSize }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      let url = deltaUrl || `${MS_GRAPH_BASE}/me/todo/lists/delta`
      if (!deltaUrl) {
        const queryParams = new URLSearchParams()
        if (deltaToken) queryParams.append("$deltatoken", deltaToken)
        if (skipToken) queryParams.append("$skiptoken", skipToken)
        if (select) queryParams.append("$select", select)
        const queryString = queryParams.toString()
        if (queryString) url += `?${queryString}`
      }

      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      }
      if (maxPageSize) {
        headers.Prefer = `odata.maxpagesize=${maxPageSize}`
      }

      const response = await fetch(url, { method: "GET", headers })
      if (!response.ok) {
        const errorText = await response.text()
        throw new Error(`HTTP error! status: ${response.status}, body: ${errorText}`)
      }

      const data = (await response.json()) as DeltaResponse<TaskList>
      const formattedLists = (data.value || []).map((list) => `ID: ${list.id}\nName: ${list.displayName}\n---`).join("\n")

      return {
        content: [
          {
            type: "text",
            text:
              `Task list delta results:\n\n${formattedLists || "No changed lists returned."}` +
              `${data["@odata.nextLink"] ? `\n\nNext Link:\n${data["@odata.nextLink"]}` : ""}` +
              `${data["@odata.deltaLink"] ? `\n\nDelta Link:\n${data["@odata.deltaLink"]}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching task list delta: ${error}`,
          },
        ],
      }
    }
  },
)

// Enhanced organized view of task lists
server.tool(
  "get-task-lists-organized",
  "Get all task lists organized into logical folders/categories based on naming patterns, emoji prefixes, and sharing status. Provides a hierarchical view similar to folder organization.",
  {
    includeIds: z.boolean().optional().describe("Include list IDs in output (default: false)"),
    groupBy: z
      .enum(["category", "shared", "type"])
      .optional()
      .describe("Grouping strategy - 'category' (default), 'shared', or 'type'"),
  },
  async ({ includeIds, groupBy }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve task lists",
            },
          ],
        }
      }

      const lists = response.value || []
      if (lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No task lists found.",
            },
          ],
        }
      }

      // Group by shared status
      if (groupBy === "shared") {
        const sharedLists = lists.filter((l) => l.isShared)
        const personalLists = lists.filter((l) => !l.isShared)

        let output = "📂 Microsoft To Do Lists - By Sharing Status\n"
        output += "=".repeat(50) + "\n\n"

        output += `👥 Shared Lists (${sharedLists.length})\n`
        sharedLists.forEach((list) => {
          const ownership = list.isOwner ? "Shared by you" : "Shared with you"
          output += `   ├─ ${list.displayName} [${ownership}]\n`
        })

        output += `\n🔒 Personal Lists (${personalLists.length})\n`
        personalLists.forEach((list) => {
          output += `   ├─ ${list.displayName}\n`
        })

        return { content: [{ type: "text", text: output }] }
      }

      // Helper function to organize lists
      const organizeLists = (lists: TaskList[]): { [category: string]: TaskList[] } => {
        const organized: { [category: string]: TaskList[] } = {}

        // Patterns for categorizing lists
        const patterns = {
          archived: /\(([^)]+)\s*-\s*Archived\)$/i,
          archive: /^📦\s*Archive/i,
          shopping: /^🛒/,
          property: /^🏡/,
          family: /^👪/,
          seasonal: /^(🎄|🎉)/,
          work: /^(Work|SBIR)/i,
          travel: /^(🚗|Rangeley)/i,
          reading: /^📰/,
        }

        lists.forEach((list) => {
          let placed = false

          // Check archived pattern
          const archiveMatch = list.displayName.match(patterns.archived)
          if (archiveMatch) {
            const category = `📦 Archived - ${archiveMatch[1]}`
            if (!organized[category]) organized[category] = []
            organized[category].push(list)
            placed = true
          }
          // Check archive prefix
          else if (patterns.archive.test(list.displayName)) {
            if (!organized["📦 Archives"]) organized["📦 Archives"] = []
            organized["📦 Archives"].push(list)
            placed = true
          }
          // Check shopping lists
          else if (patterns.shopping.test(list.displayName)) {
            if (!organized["🛒 Shopping Lists"]) organized["🛒 Shopping Lists"] = []
            organized["🛒 Shopping Lists"].push(list)
            placed = true
          }
          // Check property lists
          else if (patterns.property.test(list.displayName)) {
            if (!organized["🏡 Properties"]) organized["🏡 Properties"] = []
            organized["🏡 Properties"].push(list)
            placed = true
          }
          // Check family lists
          else if (patterns.family.test(list.displayName)) {
            if (!organized["👪 Family"]) organized["👪 Family"] = []
            organized["👪 Family"].push(list)
            placed = true
          }
          // Check seasonal lists
          else if (patterns.seasonal.test(list.displayName)) {
            if (!organized["🎉 Seasonal & Events"]) organized["🎉 Seasonal & Events"] = []
            organized["🎉 Seasonal & Events"].push(list)
            placed = true
          }
          // Check work lists
          else if (patterns.work.test(list.displayName)) {
            if (!organized["💼 Work"]) organized["💼 Work"] = []
            organized["💼 Work"].push(list)
            placed = true
          }
          // Check travel lists
          else if (patterns.travel.test(list.displayName)) {
            if (!organized["🚗 Travel & Rangeley"]) organized["🚗 Travel & Rangeley"] = []
            organized["🚗 Travel & Rangeley"].push(list)
            placed = true
          }
          // Check reading lists
          else if (patterns.reading.test(list.displayName)) {
            if (!organized["📚 Reading"]) organized["📚 Reading"] = []
            organized["📚 Reading"].push(list)
            placed = true
          }
          // Special lists
          else if (list.wellknownListName && list.wellknownListName !== "none") {
            if (!organized["⭐ Special Lists"]) organized["⭐ Special Lists"] = []
            organized["⭐ Special Lists"].push(list)
            placed = true
          }
          // Shared lists (only if not already placed)
          else if (list.isShared && !placed) {
            if (!organized["👥 Shared Lists"]) organized["👥 Shared Lists"] = []
            organized["👥 Shared Lists"].push(list)
            placed = true
          }
          // Everything else
          else {
            if (!organized["📋 Other Lists"]) organized["📋 Other Lists"] = []
            organized["📋 Other Lists"].push(list)
          }
        })

        return organized
      }

      // Default: organize by category
      const organized = organizeLists(lists)

      let output = "📂 Microsoft To Do Lists - Organized View\n"
      output += "=".repeat(50) + "\n\n"

      // Sort categories for consistent display
      const sortedCategories = Object.keys(organized).sort((a, b) => {
        // Priority order for categories
        const priority: { [key: string]: number } = {
          "⭐ Special Lists": 1,
          "👥 Shared Lists": 2,
          "💼 Work": 3,
          "👪 Family": 4,
          "🏡 Properties": 5,
          "🛒 Shopping Lists": 6,
          "🚗 Travel & Rangeley": 7,
          "🎉 Seasonal & Events": 8,
          "📚 Reading": 9,
          "📋 Other Lists": 10,
          "📦 Archives": 11,
        }

        // Check if categories start with "📦 Archived -"
        const aIsArchived = a.startsWith("📦 Archived -")
        const bIsArchived = b.startsWith("📦 Archived -")

        if (aIsArchived && !bIsArchived) return 1
        if (!aIsArchived && bIsArchived) return -1
        if (aIsArchived && bIsArchived) return a.localeCompare(b)

        const aPriority = priority[a] || 999
        const bPriority = priority[b] || 999

        if (aPriority !== bPriority) return aPriority - bPriority
        return a.localeCompare(b)
      })

      sortedCategories.forEach((category) => {
        const categoryLists = organized[category]
        output += `${category} (${categoryLists.length})\n`

        categoryLists.forEach((list, index) => {
          const isLast = index === categoryLists.length - 1
          const prefix = isLast ? "└─" : "├─"

          let listInfo = `${prefix} ${list.displayName}`

          // Add metadata
          const metadata = []
          if (list.wellknownListName === "defaultList") metadata.push("Default")
          if (list.wellknownListName === "flaggedEmails") metadata.push("Flagged Emails")
          if (list.isShared && list.isOwner) metadata.push("Shared by you")
          if (list.isShared && !list.isOwner) metadata.push("Shared with you")

          if (metadata.length > 0) {
            listInfo += ` [${metadata.join(", ")}]`
          }

          output += `   ${listInfo}\n`

          if (!isLast) {
            output += "   │\n"
          }
        })

        output += "\n"
      })

      // Add summary
      const totalLists = Object.values(organized).reduce((sum, l) => sum + l.length, 0)
      const totalCategories = Object.keys(organized).length

      output += "-".repeat(50) + "\n"
      output += `Summary: ${totalLists} lists in ${totalCategories} categories\n`

      if (includeIds) {
        // Add a section with IDs
        output += "\n\n📋 List IDs Reference:\n" + "-".repeat(50) + "\n"
        lists.forEach((list) => {
          output += `${list.displayName}: ${list.id}\n`
        })
      }

      return { content: [{ type: "text", text: output }] }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching organized task lists: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-task-list",
  "Create a new task list (top-level container) in Microsoft Todo to help organize your tasks into categories or projects.",
  {
    displayName: z.string().describe("Name of the new task list"),
  },
  async ({ displayName }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody = {
        displayName,
      }

      // Make the API request to create the task list
      const response = await makeGraphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token, "POST", requestBody)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create task list: ${displayName}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list created successfully!\nName: ${response.displayName}\nID: ${response.id}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating task list: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "update-task-list",
  "Update the name of an existing task list (top-level container) in Microsoft Todo.",
  {
    listId: z.string().describe("ID of the task list to update"),
    displayName: z.string().describe("New name for the task list"),
  },
  async ({ listId, displayName }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody = {
        displayName,
      }

      // Make the API request to update the task list
      const response = await makeGraphRequest<TaskList>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update task list with ID: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list updated successfully!\nNew name: ${response.displayName}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating task list: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "delete-task-list",
  "Delete a task list (top-level container) from Microsoft Todo. This will remove the list and all tasks within it.",
  {
    listId: z.string().describe("ID of the task list to delete"),
  },
  async ({ listId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}`
      console.error(`Deleting task list: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
      return {
        content: [
          {
            type: "text",
            text: `Task list with ID: ${listId} was successfully deleted.`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting task list: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-tasks",
  "Get tasks from a specific Microsoft Todo list. These are the main todo items that can contain checklist items (subtasks).",
  {
    listId: z.string().describe("ID of the task list"),
    filter: z.string().optional().describe("OData $filter query (e.g., 'status eq \\'completed\\'')"),
    select: z.string().optional().describe("Comma-separated list of properties to include (e.g., 'id,title,status')"),
    orderby: z.string().optional().describe("Property to sort by (e.g., 'createdDateTime desc')"),
    top: z.number().optional().describe("Maximum number of tasks to retrieve"),
    skip: z.number().optional().describe("Number of tasks to skip"),
    count: z.boolean().optional().describe("Whether to include a count of tasks"),
  },
  async ({ listId, filter, select, orderby, top, skip, count }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Build the query parameters
      const queryParams = new URLSearchParams()

      if (filter) queryParams.append("$filter", filter)
      if (select) queryParams.append("$select", select)
      if (orderby) queryParams.append("$orderby", orderby)
      if (top !== undefined) queryParams.append("$top", top.toString())
      if (skip !== undefined) queryParams.append("$skip", skip.toString())
      if (count !== undefined) queryParams.append("$count", count.toString())

      // Construct the URL with query parameters
      const queryString = queryParams.toString()
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks${queryString ? "?" + queryString : ""}`

      console.error(`Making request to: ${url}`)

      const response = await makeGraphRequest<{ value: Task[]; "@odata.count"?: number }>(url, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve tasks for list: ${listId}`,
            },
          ],
        }
      }

      const tasks = response.value || []
      if (tasks.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No tasks found in list with ID: ${listId}`,
            },
          ],
        }
      }

      const formattedTasks = tasks.map((task) => formatTask(task))

      // Add count information if requested and available
      let countInfo = ""
      if (count && response["@odata.count"] !== undefined) {
        countInfo = `Total count: ${response["@odata.count"]}\n\n`
      }

      return {
        content: [
          {
            type: "text",
            text: `Tasks in list ${listId}:\n\n${countInfo}${formattedTasks.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching tasks: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-task",
  "Get a single Microsoft Todo task by ID.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    select: z.string().optional().describe("Comma-separated list of properties to include"),
  },
  async ({ listId, taskId, select }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const queryParams = new URLSearchParams()
      if (select) queryParams.append("$select", select)

      const queryString = queryParams.toString()
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}${queryString ? "?" + queryString : ""}`
      const task = await makeGraphRequest<Task>(url, token)

      if (!task) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve task: ${taskId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task details:\n\n${formatTask(task)}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching task: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-tasks-delta",
  "Track changes to tasks in a Microsoft Todo list using the Graph delta API.",
  {
    listId: z.string().describe("ID of the task list"),
    deltaUrl: z.string().optional().describe("Full @odata.nextLink or @odata.deltaLink URL from a previous delta response"),
    deltaToken: z.string().optional().describe("Delta token from a previous delta response"),
    skipToken: z.string().optional().describe("Skip token from a previous delta response"),
    select: z.string().optional().describe("Comma-separated list of properties to include on the initial request"),
    top: z.number().int().min(1).optional().describe("Maximum number of tasks to retrieve on the initial request"),
    expand: z.string().optional().describe("OData $expand expression for the initial request"),
    maxPageSize: z.number().int().min(1).optional().describe("Preferred maximum number of tasks returned"),
  },
  async ({ listId, deltaUrl, deltaToken, skipToken, select, top, expand, maxPageSize }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      let url = deltaUrl || `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/delta`
      if (!deltaUrl) {
        const queryParams = new URLSearchParams()
        if (deltaToken) queryParams.append("$deltatoken", deltaToken)
        if (skipToken) queryParams.append("$skiptoken", skipToken)
        if (select) queryParams.append("$select", select)
        if (top !== undefined) queryParams.append("$top", top.toString())
        if (expand) queryParams.append("$expand", expand)
        const queryString = queryParams.toString()
        if (queryString) url += `?${queryString}`
      }

      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      }
      if (maxPageSize) {
        headers.Prefer = `odata.maxpagesize=${maxPageSize}`
      }

      const response = await fetch(url, { method: "GET", headers })
      if (!response.ok) {
        const errorText = await response.text()
        throw new Error(`HTTP error! status: ${response.status}, body: ${errorText}`)
      }

      const data = (await response.json()) as DeltaResponse<Task>
      const formattedTasks = (data.value || []).map((task) => formatTask(task)).join("\n")

      return {
        content: [
          {
            type: "text",
            text:
              `Task delta results for list ${listId}:\n\n${formattedTasks || "No changed tasks returned."}` +
              `${data["@odata.nextLink"] ? `\n\nNext Link:\n${data["@odata.nextLink"]}` : ""}` +
              `${data["@odata.deltaLink"] ? `\n\nDelta Link:\n${data["@odata.deltaLink"]}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching task delta: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-task",
  "Create a new task in a specific Microsoft Todo list. A task is the main todo item that can have a title, description, due date, and other properties.",
  {
    listId: z.string().describe("ID of the task list"),
    title: z.string().describe("Title of the task"),
    body: z.string().optional().describe("Description or body content of the task"),
    dueDateTime: z.string().optional().describe("Due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    startDateTime: z.string().optional().describe("Start date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    completedDateTime: z.string().optional().describe("Completion date in ISO format"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("Task importance"),
    isReminderOn: z.boolean().optional().describe("Whether to enable reminder for this task"),
    reminderDateTime: z.string().optional().describe("Reminder date and time in ISO format"),
    recurrence: recurrenceSchema.optional().describe("Structured recurrence definition"),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("Status of the task"),
    categories: z.array(z.string()).optional().describe("Categories associated with the task"),
    linkedResources: z.array(linkedResourceSchema).optional().describe("Linked resources to create with the task"),
  },
  async ({
    listId,
    title,
    body,
    dueDateTime,
    startDateTime,
    completedDateTime,
    importance,
    isReminderOn,
    reminderDateTime,
    recurrence,
    status,
    categories,
    linkedResources,
  }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Construct the task body with all supported properties
      const taskBody: any = { title }

      // Add optional properties if provided
      if (body) {
        taskBody.body = {
          content: body,
          contentType: "text",
        }
      }

      if (dueDateTime) {
        taskBody.dueDateTime = buildDateTimeTimeZone(dueDateTime)
      }

      if (startDateTime) {
        taskBody.startDateTime = buildDateTimeTimeZone(startDateTime)
      }

      if (completedDateTime) {
        taskBody.completedDateTime = buildDateTimeTimeZone(completedDateTime)
      }

      if (importance) {
        taskBody.importance = importance
      }

      if (isReminderOn !== undefined) {
        taskBody.isReminderOn = isReminderOn
      }

      if (reminderDateTime) {
        taskBody.reminderDateTime = buildDateTimeTimeZone(reminderDateTime)
      }

      if (recurrence) {
        taskBody.recurrence = buildRecurrencePayload(recurrence)
      }

      if (status) {
        taskBody.status = status
      }

      if (categories && categories.length > 0) {
        taskBody.categories = categories
      }

      if (linkedResources && linkedResources.length > 0) {
        taskBody.linkedResources = linkedResources
      }

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
        token,
        "POST",
        taskBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create task in list: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task created successfully!\nID: ${response.id}\nTitle: ${response.title}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating task: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "update-task",
  "Update an existing task in Microsoft Todo. Allows changing any properties of the task including title, due date, importance, etc.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to update"),
    title: z.string().optional().describe("New title of the task"),
    body: z.string().optional().describe("New description or body content of the task"),
    dueDateTime: z.string().optional().describe("New due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    startDateTime: z.string().optional().describe("New start date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    completedDateTime: z
      .string()
      .optional()
      .describe("New completion date in ISO format. Pass an empty string to clear it."),
    importance: z.enum(["low", "normal", "high"]).optional().describe("New task importance"),
    isReminderOn: z.boolean().optional().describe("Whether to enable reminder for this task"),
    reminderDateTime: z.string().optional().describe("New reminder date and time in ISO format"),
    recurrence: recurrenceSchema.nullable().optional().describe("New recurrence definition. Pass null to clear it."),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("New status of the task"),
    categories: z.array(z.string()).optional().describe("New categories associated with the task"),
  },
  async ({
    listId,
    taskId,
    title,
    body,
    dueDateTime,
    startDateTime,
    completedDateTime,
    importance,
    isReminderOn,
    reminderDateTime,
    recurrence,
    status,
    categories,
  }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      let existingTask: Task | null = null

      // Construct the task update body with all provided properties
      const taskBody: any = {}

      // Add optional properties if provided
      if (title !== undefined) {
        taskBody.title = title
      }

      if (body !== undefined) {
        taskBody.body = {
          content: body,
          contentType: "text",
        }
      }

      if (dueDateTime !== undefined) {
        if (dueDateTime === "") {
          // Remove the due date by setting it to null
          taskBody.dueDateTime = null
        } else {
          taskBody.dueDateTime = buildDateTimeTimeZone(dueDateTime)
        }
      }

      if (startDateTime !== undefined) {
        if (startDateTime === "") {
          // Remove the start date by setting it to null
          taskBody.startDateTime = null
        } else {
          taskBody.startDateTime = buildDateTimeTimeZone(startDateTime)
        }
      }

      if (completedDateTime !== undefined) {
        if (completedDateTime === "") {
          taskBody.completedDateTime = null
        } else {
          taskBody.completedDateTime = buildDateTimeTimeZone(completedDateTime)
        }
      }

      if (importance !== undefined) {
        taskBody.importance = importance
      }

      if (isReminderOn !== undefined) {
        taskBody.isReminderOn = isReminderOn
      }

      if (reminderDateTime !== undefined) {
        if (reminderDateTime === "") {
          // Remove the reminder date by setting it to null
          taskBody.reminderDateTime = null
        } else {
          taskBody.reminderDateTime = buildDateTimeTimeZone(reminderDateTime)
        }
      }

      if (recurrence !== undefined) {
        if (recurrence === null) {
          taskBody.recurrence = null
        } else {
          if (taskBody.dueDateTime === undefined) {
            existingTask = await makeGraphRequest<Task>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`, token)

            if (!existingTask) {
              return {
                content: [
                  {
                    type: "text",
                    text: `Failed to load task ${taskId} to determine the due date required for recurrence updates.`,
                  },
                ],
              }
            }

            if (!existingTask.dueDateTime) {
              return {
                content: [
                  {
                    type: "text",
                    text:
                      "Microsoft Graph requires dueDateTime when adding or updating recurrence. " +
                      "This task has no due date, so include dueDateTime in the update request.",
                  },
                ],
              }
            }

            if (!isFutureOrCurrentDateTime(existingTask.dueDateTime)) {
              return {
                content: [
                  {
                    type: "text",
                    text:
                      "Microsoft Graph requires dueDateTime when adding or updating recurrence. " +
                      "This task's existing due date is already in the past, so specify a new dueDateTime in the update request.",
                  },
                ],
              }
            }

            taskBody.dueDateTime = existingTask.dueDateTime
          } else if (taskBody.dueDateTime === null) {
            return {
              content: [
                {
                  type: "text",
                  text: "Cannot clear dueDateTime while setting recurrence. Microsoft Graph requires dueDateTime for recurrence updates.",
                },
              ],
            }
          }

          taskBody.recurrence = buildRecurrencePatchPayload(recurrence)
        }
      }

      if (status !== undefined) {
        taskBody.status = status
      }

      if (categories !== undefined) {
        taskBody.categories = categories
      }

      // Make sure we have at least one property to update
      if (Object.keys(taskBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify at least one property to change.",
            },
          ],
        }
      }

      const response = await makeGraphRequest<Task>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`, token, "PATCH", taskBody)

      if (!response) {
        if (taskBody.recurrence !== undefined && isRecurrencePatchDateError(lastGraphRequestError)) {
          const requestIdInfo = lastGraphRequestError?.requestId
            ? `\nRequest ID: ${lastGraphRequestError.requestId}`
            : ""
          const clientRequestIdInfo = lastGraphRequestError?.clientRequestId
            ? `\nClient Request ID: ${lastGraphRequestError.clientRequestId}`
            : ""

          return {
            content: [
              {
                type: "text",
                text:
                  "Microsoft Graph rejected the recurrence update even though the MCP server sent a minimal documented recurrence payload. " +
                  "This appears to be a Graph To Do API limitation or bug affecting PATCH updates to recurrence.range.startDate." +
                  "\nThe exact request and response bodies were logged to stderr for debugging." +
                  requestIdInfo +
                  clientRequestIdInfo,
              },
            ],
          }
        }

        return {
          content: [
            {
              type: "text",
              text: `Failed to update task with ID: ${taskId} in list: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task updated successfully!\nID: ${response.id}\nTitle: ${response.title}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating task: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "delete-task",
  "Delete a task from a Microsoft Todo list. This will remove the task and all its checklist items (subtasks).",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to delete"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`
      console.error(`Deleting task: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
      return {
        content: [
          {
            type: "text",
            text: `Task with ID: ${taskId} was successfully deleted from list: ${listId}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting task: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-linked-resources",
  "Get linked resources for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: LinkedResource[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources`,
        token,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve linked resources for task: ${taskId}`,
            },
          ],
        }
      }

      const linkedResources = response.value || []
      if (linkedResources.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No linked resources found for task: ${taskId}`,
            },
          ],
        }
      }

      const formattedResources = linkedResources.map((resource) => {
        const details = [
          `ID: ${resource.id || "Unknown"}`,
          `Display Name: ${resource.displayName || "Unknown"}`,
          `Application: ${resource.applicationName || "Unknown"}`,
        ]

        if (resource.webUrl) details.push(`URL: ${resource.webUrl}`)
        if (resource.externalId) details.push(`External ID: ${resource.externalId}`)

        return `${details.join("\n")}\n---`
      })

      return {
        content: [
          {
            type: "text",
            text: `Linked resources for task ${taskId}:\n\n${formattedResources.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching linked resources: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-linked-resource",
  "Create a linked resource for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    webUrl: z.string().optional().describe("Deep link to the linked item"),
    applicationName: z.string().optional().describe("Source application name"),
    displayName: z.string().optional().describe("Display name for the linked item"),
    externalId: z.string().optional().describe("External identifier from the source system"),
  },
  async ({ listId, taskId, webUrl, applicationName, displayName, externalId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<LinkedResource>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources`,
        token,
        "POST",
        {
          webUrl,
          applicationName,
          displayName,
          externalId,
        },
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create linked resource for task: ${taskId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Linked resource created successfully!\nID: ${response.id || "Unknown"}` +
              `${response.displayName ? `\nDisplay Name: ${response.displayName}` : ""}` +
              `${response.webUrl ? `\nURL: ${response.webUrl}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating linked resource: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-attachments",
  "List file attachments for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: TaskFileAttachment[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments`,
        token,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve attachments for task: ${taskId}`,
            },
          ],
        }
      }

      const attachments = response.value || []
      if (attachments.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No attachments found for task: ${taskId}`,
            },
          ],
        }
      }

      const formattedAttachments = attachments.map((attachment) => {
        const details = [
          `ID: ${attachment.id}`,
          `Name: ${attachment.name}`,
          `Size: ${attachment.size ?? "Unknown"} bytes`,
        ]

        if (attachment.contentType) details.push(`Content Type: ${attachment.contentType}`)
        if (attachment.lastModifiedDateTime) {
          details.push(`Last Modified: ${new Date(attachment.lastModifiedDateTime).toLocaleString()}`)
        }

        return `${details.join("\n")}\n---`
      })

      return {
        content: [
          {
            type: "text",
            text: `Attachments for task ${taskId}:\n\n${formattedAttachments.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching attachments: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-attachment",
  "Get a single file attachment for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    attachmentId: z.string().describe("ID of the attachment"),
  },
  async ({ listId, taskId, attachmentId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value?: TaskFileAttachment } | TaskFileAttachment>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}`,
        token,
      )

      let attachment: TaskFileAttachment | undefined
      if (response && typeof response === "object" && "value" in response) {
        attachment = response.value
      } else if (isTaskFileAttachment(response)) {
        attachment = response
      }
      if (!attachment) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve attachment: ${attachmentId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Attachment details:\n\nID: ${attachment.id}\nName: ${attachment.name}` +
              `${attachment.contentType ? `\nContent Type: ${attachment.contentType}` : ""}` +
              `${attachment.size !== undefined ? `\nSize: ${attachment.size} bytes` : ""}` +
              `${attachment.lastModifiedDateTime ? `\nLast Modified: ${new Date(attachment.lastModifiedDateTime).toLocaleString()}` : ""}` +
              `${attachment.contentBytes ? `\nContent Bytes Present: Yes (${attachment.contentBytes.length} base64 chars)` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching attachment: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-attachment",
  "Create a small file attachment on a Microsoft Todo task. For files larger than 3 MB, use create-attachment-upload-session.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    name: z.string().describe("Attachment display name"),
    contentBytes: z.string().describe("Base64-encoded file contents"),
    contentType: z.string().optional().describe("Attachment content type"),
  },
  async ({ listId, taskId, name, contentBytes, contentType }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<TaskFileAttachment>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments`,
        token,
        "POST",
        {
          "@odata.type": "#microsoft.graph.taskFileAttachment",
          name,
          contentBytes,
          contentType,
        },
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create attachment for task: ${taskId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Attachment created successfully!\nID: ${response.id}\nName: ${response.name}` +
              `${response.size !== undefined ? `\nSize: ${response.size} bytes` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating attachment: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-attachment-upload-session",
  "Create an upload session for a large file attachment on a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    name: z.string().describe("Attachment display name"),
    size: z.number().int().min(0).describe("Attachment size in bytes"),
  },
  async ({ listId, taskId, name, size }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<UploadSession>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/createUploadSession`,
        token,
        "POST",
        {
          attachmentInfo: {
            attachmentType: "file",
            name,
            size,
          },
        },
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create attachment upload session for task: ${taskId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Attachment upload session created successfully!\nUpload URL: ${response.uploadUrl}\nExpiration: ${response.expirationDateTime}\nNext Expected Ranges: ${response.nextExpectedRanges.join(", ")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating attachment upload session: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "delete-attachment",
  "Delete a file attachment from a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    attachmentId: z.string().describe("ID of the attachment"),
  },
  async ({ listId, taskId, attachmentId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      await makeGraphRequest<null>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}`, token, "DELETE")

      return {
        content: [
          {
            type: "text",
            text: `Attachment with ID: ${attachmentId} was successfully deleted from task: ${taskId}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting attachment: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-checklist-items",
  "Get checklist items (subtasks) for a specific task. Checklist items are smaller steps or components that belong to a parent task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Fetch the task first to get its title
      const taskResponse = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
      )

      const taskTitle = taskResponse ? taskResponse.title : "Unknown Task"

      // Fetch the checklist items
      const response = await makeGraphRequest<{ value: ChecklistItem[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve checklist items for task: ${taskId}`,
            },
          ],
        }
      }

      const items = response.value || []
      if (items.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No checklist items found for task "${taskTitle}" (ID: ${taskId})`,
            },
          ],
        }
      }

      const formattedItems = items.map((item) => {
        const status = item.isChecked ? "✓" : "○"
        let itemInfo = `${status} ${item.displayName} (ID: ${item.id})`

        // Add creation date if available
        if (item.createdDateTime) {
          const createdDate = new Date(item.createdDateTime).toLocaleString()
          itemInfo += `\nCreated: ${createdDate}`
        }

        if (item.checkedDateTime) {
          const checkedDate = new Date(item.checkedDateTime).toLocaleString()
          itemInfo += `\nChecked: ${checkedDate}`
        }

        return itemInfo
      })

      return {
        content: [
          {
            type: "text",
            text: `Checklist items for task "${taskTitle}" (ID: ${taskId}):\n\n${formattedItems.join("\n\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching checklist items: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-checklist-item",
  "Create a new checklist item (subtask) for a task. Checklist items help break down a task into smaller, manageable steps.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    displayName: z.string().describe("Text content of the checklist item"),
    isChecked: z.boolean().optional().describe("Whether the item is checked off"),
    checkedDateTime: z.string().optional().describe("Completion timestamp in ISO format"),
    createdDateTime: z.string().optional().describe("Creation timestamp in ISO format"),
  },
  async ({ listId, taskId, displayName, isChecked, checkedDateTime, createdDateTime }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody: any = {
        displayName,
      }

      if (isChecked !== undefined) {
        requestBody.isChecked = isChecked
      }

      if (checkedDateTime !== undefined) {
        requestBody.checkedDateTime = checkedDateTime
      }

      if (createdDateTime !== undefined) {
        requestBody.createdDateTime = createdDateTime
      }

      // Make the API request to create the checklist item
      const response = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
        "POST",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create checklist item for task: ${taskId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Checklist item created successfully!\nContent: ${response.displayName}\nID: ${response.id}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "update-checklist-item",
  "Update an existing checklist item (subtask). Allows changing the text content or completion status of the subtask.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item to update"),
    displayName: z.string().optional().describe("New text content of the checklist item"),
    isChecked: z.boolean().optional().describe("Whether the item is checked off"),
    checkedDateTime: z
      .string()
      .optional()
      .describe("Completion timestamp in ISO format. Pass an empty string to clear it."),
    createdDateTime: z.string().optional().describe("Creation timestamp in ISO format"),
  },
  async ({ listId, taskId, checklistItemId, displayName, isChecked, checkedDateTime, createdDateTime }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the update body, including only the fields that are provided
      const requestBody: any = {}

      if (displayName !== undefined) {
        requestBody.displayName = displayName
      }

      if (isChecked !== undefined) {
        requestBody.isChecked = isChecked
      }

      if (checkedDateTime !== undefined) {
        requestBody.checkedDateTime = checkedDateTime === "" ? null : checkedDateTime
      }

      if (createdDateTime !== undefined) {
        requestBody.createdDateTime = createdDateTime
      }

      // Make sure we have at least one property to update
      if (Object.keys(requestBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify at least one checklist item property to change.",
            },
          ],
        }
      }

      // Make the API request to update the checklist item
      const response = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update checklist item with ID: ${checklistItemId}`,
            },
          ],
        }
      }

      const statusText = response.isChecked ? "Checked" : "Not checked"

      return {
        content: [
          {
            type: "text",
            text: `Checklist item updated successfully!\nContent: ${response.displayName}\nStatus: ${statusText}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "delete-checklist-item",
  "Delete a checklist item (subtask) from a task. This removes just the specific subtask, not the parent task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item to delete"),
  },
  async ({ listId, taskId, checklistItemId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`
      console.error(`Deleting checklist item: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
      return {
        content: [
          {
            type: "text",
            text: `Checklist item with ID: ${checklistItemId} was successfully deleted from task: ${taskId}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

// Bulk archive completed tasks
server.tool(
  "archive-completed-tasks",
  "Move completed tasks older than a specified number of days from one list to another (archive) list. Useful for cleaning up active lists while preserving historical tasks.",
  {
    sourceListId: z.string().describe("ID of the source list to archive tasks from"),
    targetListId: z.string().describe("ID of the target archive list"),
    olderThanDays: z
      .number()
      .min(0)
      .default(90)
      .describe("Archive tasks completed more than this many days ago (default: 90)"),
    dryRun: z
      .boolean()
      .optional()
      .default(false)
      .describe("If true, only preview what would be archived without making changes"),
  },
  async ({ sourceListId, targetListId, olderThanDays, dryRun }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Calculate cutoff date
      const cutoffDate = new Date()
      cutoffDate.setDate(cutoffDate.getDate() - olderThanDays)

      // Get all completed tasks from source list
      const tasksResponse = await makeGraphRequest<{ value: Task[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks?$filter=status eq 'completed'`,
        token,
      )

      if (!tasksResponse || !tasksResponse.value) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve tasks from source list",
            },
          ],
        }
      }

      // Filter tasks older than cutoff
      const tasksToArchive = tasksResponse.value.filter((task) => {
        if (!task.completedDateTime?.dateTime) return false
        const completedDate = new Date(task.completedDateTime.dateTime)
        return completedDate < cutoffDate
      })

      if (tasksToArchive.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No completed tasks found older than ${olderThanDays} days.`,
            },
          ],
        }
      }

      if (dryRun) {
        // Preview mode - just show what would be archived
        let preview = `📋 Archive Preview\n`
        preview += `Would archive ${tasksToArchive.length} tasks completed before ${cutoffDate.toLocaleDateString()}\n\n`

        tasksToArchive.forEach((task) => {
          const completedDate = task.completedDateTime?.dateTime
            ? new Date(task.completedDateTime.dateTime).toLocaleDateString()
            : "Unknown"
          preview += `- ${task.title} (completed: ${completedDate})\n`
        })

        return { content: [{ type: "text", text: preview }] }
      }

      // Actually archive the tasks
      let successCount = 0
      let failedTasks: string[] = []

      for (const task of tasksToArchive) {
        try {
          // Create task in target list
          const createResponse = await makeGraphRequest(
            `${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks`,
            token,
            "POST",
            {
              title: task.title,
              status: "completed",
              body: task.body,
              importance: task.importance,
              completedDateTime: task.completedDateTime,
              dueDateTime: task.dueDateTime,
              startDateTime: task.startDateTime,
              reminderDateTime: task.reminderDateTime,
              isReminderOn: task.isReminderOn,
              recurrence: task.recurrence,
              categories: task.categories,
              linkedResources: task.linkedResources,
            },
          )

          if (createResponse) {
            // Delete from source list
            await makeGraphRequest(`${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${task.id}`, token, "DELETE")
            successCount++
          } else {
            failedTasks.push(task.title)
          }
        } catch (error) {
          failedTasks.push(task.title)
        }
      }

      let result = `📦 Archive Complete\n`
      result += `Successfully archived ${successCount} of ${tasksToArchive.length} tasks\n`
      result += `Tasks completed before ${cutoffDate.toLocaleDateString()} were moved.\n`

      if (failedTasks.length > 0) {
        result += `\n⚠️ Failed to archive ${failedTasks.length} tasks:\n`
        failedTasks.forEach((title) => {
          result += `- ${title}\n`
        })
      }

      return { content: [{ type: "text", text: result }] }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error archiving tasks: ${error}`,
          },
        ],
      }
    }
  },
)

// Test tool to explore Graph API for hidden properties
server.tool(
  "test-graph-api-exploration",
  "Test various Graph API queries to discover hidden properties or endpoints for folder/group organization in Microsoft To Do.",
  {
    testType: z.enum(["odata-select", "odata-expand", "headers", "extensions", "all"]).describe("Type of test to run"),
  },
  async ({ testType }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      let results = "🔍 Graph API Exploration Results\n" + "=".repeat(50) + "\n\n"

      // Test 1: Try with $select=* to get all properties
      if (testType === "odata-select" || testType === "all") {
        results += "📊 Test 1: Using $select=* to retrieve all properties\n"
        try {
          const response = await makeGraphRequest<any>(`${MS_GRAPH_BASE}/me/todo/lists?$select=*`, token)
          if (response && response.value && response.value.length > 0) {
            const firstList = response.value[0]
            const properties = Object.keys(firstList)
            results += `Found ${properties.length} properties: ${properties.join(", ")}\n`

            // Show full first list as example
            results += "\nExample list object:\n"
            results += JSON.stringify(firstList, null, 2).substring(0, 1000) + "...\n"
          }
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 2: Try various $expand options
      if (testType === "odata-expand" || testType === "all") {
        results += "📊 Test 2: Using $expand to retrieve related data\n"
        const expandOptions = [
          "extensions",
          "singleValueExtendedProperties",
          "multiValueExtendedProperties",
          "openExtensions",
          "parent",
          "children",
          "folder",
          "parentFolder",
          "group",
          "category",
        ]

        for (const expand of expandOptions) {
          try {
            const response = await makeGraphRequest<any>(
              `${MS_GRAPH_BASE}/me/todo/lists?$expand=${expand}&$top=1`,
              token,
            )
            if (response && response.value) {
              results += `✓ $expand=${expand}: Success - `
              if (response.value.length > 0 && response.value[0][expand]) {
                results += `Found data!\n`
                results += JSON.stringify(response.value[0][expand], null, 2).substring(0, 500) + "...\n"
              } else {
                results += `No additional data returned\n`
              }
            }
          } catch (error: any) {
            results += `✗ $expand=${expand}: ${error.message || "Failed"}\n`
          }
        }
        results += "\n"
      }

      // Test 3: Check response headers for additional info
      if (testType === "headers" || testType === "all") {
        results += "📊 Test 3: Checking response headers\n"
        try {
          const response = await fetch(`${MS_GRAPH_BASE}/me/todo/lists`, {
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: "application/json",
              Prefer: "return=representation",
            },
          })

          results += "Response headers:\n"
          response.headers.forEach((value, key) => {
            results += `${key}: ${value}\n`
          })
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 4: Try extensions endpoint
      if (testType === "extensions" || testType === "all") {
        results += "📊 Test 4: Checking for extensions\n"
        try {
          const listsResponse = await makeGraphRequest<{ value: TaskList[] }>(
            `${MS_GRAPH_BASE}/me/todo/lists?$top=1`,
            token,
          )

          if (listsResponse && listsResponse.value && listsResponse.value.length > 0) {
            const listId = listsResponse.value[0].id

            // Try to get extensions
            try {
              const extResponse = await makeGraphRequest<any>(
                `${MS_GRAPH_BASE}/me/todo/lists/${listId}/extensions`,
                token,
              )
              results += `Extensions found: ${JSON.stringify(extResponse, null, 2)}\n`
            } catch (error: any) {
              results += `No extensions endpoint: ${error.message}\n`
            }
          }
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 5: Check if there's a separate folders or groups endpoint
      if (testType === "all") {
        results += "📊 Test 5: Checking for folder/group endpoints\n"
        const endpoints = [
          "/me/todo/folders",
          "/me/todo/groups",
          "/me/todo/listGroups",
          "/me/todo/listFolders",
          "/me/todo/categories",
        ]

        for (const endpoint of endpoints) {
          try {
            const response = await makeGraphRequest<any>(`${MS_GRAPH_BASE}${endpoint}`, token)
            results += `✓ ${endpoint}: Found! Response: ${JSON.stringify(response).substring(0, 200)}...\n`
          } catch (error: any) {
            results += `✗ ${endpoint}: Not found (${error.message || "Failed"})\n`
          }
        }
      }

      results += "\n" + "=".repeat(50) + "\n"
      results += "Analysis complete. Check results above for any discovered properties or endpoints."

      return {
        content: [
          {
            type: "text",
            text: results,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error during Graph API exploration: ${error}`,
          },
        ],
      }
    }
  },
)

// Main function to start the server
export async function startServer(config?: ServerConfig): Promise<void> {
  try {
    if (config?.tokenFilePath) {
      process.env.MSTODO_TOKEN_FILE = config.tokenFilePath
      tokenManager.configure({ tokenFilePath: config.tokenFilePath })
    }

    if (config?.accessToken) {
      process.env.MS_TODO_ACCESS_TOKEN = config.accessToken
    }

    if (config?.refreshToken) {
      process.env.MS_TODO_REFRESH_TOKEN = config.refreshToken
    }

    // Check if using a personal Microsoft account and show warning if needed
    await isPersonalMicrosoftAccount()

    // Start the server
    const transport = new StdioServerTransport()
    await server.connect(transport)

    console.error("Server started and listening")
  } catch (error) {
    console.error("Error starting server:", error)
    throw error
  }
}

// Main entry point when executed directly
if (import.meta.url === `file://${process.argv[1]}`) {
  startServer().catch((error) => {
    console.error("Fatal error in main():", error)
    process.exit(1)
  })
}
