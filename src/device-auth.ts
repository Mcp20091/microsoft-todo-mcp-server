// src/device-auth.ts
// Sign in with the OAuth device code flow: only a client ID is needed (no
// client secret, no redirect URI). The app registration must allow public
// client flows. Usage: pnpm run auth:device [CLIENT_ID]
import { tokenManager, TOKEN_SCOPES } from "./token-manager.js"

const AUTHORITY = "https://login.microsoftonline.com"
// /common serves apps for work and personal accounts; apps registered for
// personal accounts only must use /consumers (AADSTS9002331).
const CONSUMERS_ONLY_ERROR = "AADSTS9002331"

interface DeviceCodeResponse {
  device_code: string
  user_code: string
  verification_uri: string
  interval: number
  expires_in: number
  message?: string
}

async function postForm(url: string, form: Record<string, string>): Promise<{ ok: boolean; data: any }> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(form),
  })
  return { ok: response.ok, data: await response.json() }
}

async function requestCode(
  clientId: string,
  tenantId: string,
): Promise<{ code: DeviceCodeResponse; tenantId: string }> {
  const result = await postForm(`${AUTHORITY}/${tenantId}/oauth2/v2.0/devicecode`, {
    client_id: clientId,
    scope: TOKEN_SCOPES,
  })
  if (result.ok) return { code: result.data as DeviceCodeResponse, tenantId }
  const description: string = result.data?.error_description ?? ""
  if (tenantId === "common" && description.includes(CONSUMERS_ONLY_ERROR)) {
    return requestCode(clientId, "consumers")
  }
  throw new Error(description.split(/\r?\n/)[0] || result.data?.error || "Could not get a sign-in code")
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

async function main(): Promise<void> {
  const clientId = process.argv[2] || process.env.CLIENT_ID
  if (!clientId) {
    console.error("Usage: pnpm run auth:device <CLIENT_ID>   (or set CLIENT_ID)")
    process.exit(1)
  }

  const { code, tenantId } = await requestCode(clientId, process.env.TENANT_ID || "common")
  console.log(`\nOpen ${code.verification_uri} and enter the code: ${code.user_code}\n`)
  console.log("Waiting for you to sign in and approve...")

  let interval = (code.interval || 5) * 1000
  const deadline = Date.now() + code.expires_in * 1000
  while (Date.now() < deadline) {
    await sleep(interval)
    const result = await postForm(`${AUTHORITY}/${tenantId}/oauth2/v2.0/token`, {
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: clientId,
      device_code: code.device_code,
    })
    if (result.ok) {
      const data = result.data
      const tokenFile = tokenManager.getTokenFilePath()
      tokenManager.saveTokens({
        accessToken: data.access_token,
        refreshToken: data.refresh_token,
        expiresAt: Date.now() + data.expires_in * 1000 - 5 * 60 * 1000,
        clientId,
        tenantId,
      })
      console.log(`\nSigned in. Tokens saved to ${tokenFile}`)
      return
    }
    const error: string = result.data?.error ?? ""
    if (error === "authorization_pending") continue
    if (error === "slow_down") {
      interval += 5000
      continue
    }
    throw new Error((result.data?.error_description ?? error).split(/\r?\n/)[0])
  }
  throw new Error("The code expired before sign-in finished. Run this again.")
}

main().catch((error) => {
  console.error(`\nSign-in failed: ${error instanceof Error ? error.message : error}`)
  process.exit(1)
})
