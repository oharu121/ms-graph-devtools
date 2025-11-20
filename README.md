# Azure Microsoft Graph Utility

[![npm version](https://badge.fury.io/js/ms-graph-devtools.svg)](https://badge.fury.io/js/ms-graph-devtools)
![License](https://img.shields.io/npm/l/ms-graph-devtools)
![Types](https://img.shields.io/npm/types/ms-graph-devtools)
![NPM Downloads](https://img.shields.io/npm/dw/ms-graph-devtools)
![Last Commit](https://img.shields.io/github/last-commit/oharu121/ms-graph-devtools)
![GitHub Stars](https://img.shields.io/github/stars/oharu121/ms-graph-devtools?style=social)

A TypeScript utility for Microsoft Graph API operations with automatic token management and cross-platform storage support.

## ✨ Features

- ✅ **Global instance pattern** - Configure once, export, use everywhere in your app
- ✅ **AWS SDK-style API** - Familiar `Azure.service` pattern (e.g., `Azure.outlook.sendMail()`)
- ✅ **Automatic token refresh** - Handles token expiration transparently
- ✅ **Multiple auth modes** - Access token, refresh token, or custom token provider
- ✅ **Cross-platform storage** - Works on Windows, Mac, and Linux using XDG standards
- ✅ **TypeScript support** - Full type safety and autocomplete
- ✅ **Fluent builders** - Clean, chainable APIs for emails and Teams cards
- ✅ **Zero config imports** - Import configured instance, no setup needed at call site

## Major Achievement: Three-Tier User System

Successfully designed and implemented a progressive complexity model for Azure authentication that supports all user types from beginners to enterprise.

### Tier 1: Light User (Access Token Only)

**Purpose:** Quick testing, POC, experimentation

**API:**

```typescript
Azure.config({ accessToken: "eyJ0eX..." });
await Azure.outlook.getMe();
```

**Key Features:**

- Zero setup - paste token and go
- No credentials required
- Perfect for testing
- Expires in ~1 hour
- No automatic renewal

**Critical Design Decision:**

- ✅ Don't assume `expiredAt` - only set when Azure provides it
- ✅ Fail gracefully with helpful error message
- ✅ Don't force users to "upgrade" - respect their choice

### Tier 2: Medium User (Refresh Token)

**Purpose:** Production automation, scheduled tasks, single-tenant apps

**API:**

```typescript
Azure.config({
  refreshToken: "xxx",
  clientId: "...",
  clientSecret: "...",
  tenantId: "...",
});
await Azure.outlook.sendMail({...});
```

**Key Features:**

- Automatic token refresh for ~90 days
- Cross-platform storage (XDG standard)
- Works for 90% of users

**Storage Strategy:**

- ✅ Store: `refreshToken`, `accessToken`, `clientId`, `tenantId`, `expiresAt`
- ❌ NEVER store: `clientSecret` (security best practice)
- Client secret must be provided via init config

### Tier 3: Super User (Token Provider)

**Purpose:** Enterprise apps, multi-tenant SaaS, long-running services

**API:**

```typescript
Azure.config({
  clientId: "...",
  clientSecret: "...",
  tenantId: "...",
  tokenProvider: async (callback) => await getAuthCode(callback),
});
await Azure.teams.postMessage({...});
```

**Key Features:**

- Never expires (provider handles rotation)
- Multi-tenant support
- Integration with secret vaults (HashiCorp Vault, AWS Secrets Manager, etc.)
- Provider called only when needed
- Skips storage (provider is source of truth)

## 🚀 Quick Start

### Installation

```bash
npm install ms-graph-devtools
```

### Best Practice: Global Instance Pattern (Recommended)

**Create a single config file and export the configured instance:**

```typescript
// config/azure.ts (or src/lib/azure.ts, etc.)
import Azure from 'ms-graph-devtools';

Azure.config({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  tokenProvider: async (callback) => {
    return await Playwright.getAzureCode(callback);
  },
});

// Export configured instance - this is your single source of truth
export default Azure;
```

**Then import and use everywhere in your app:**

```typescript
// anywhere in your codebase
import Azure from './config/azure';

// Use services directly - already configured!
await Azure.outlook.sendMail({...});
await Azure.teams.postMessage({...});
await Azure.calendar.getCalendars();
await Azure.sharePoint.getLists();
```

**Why this pattern?**
- ✅ Configure once, use everywhere
- ✅ Single source of truth for credentials
- ✅ No config duplication across files
- ✅ Easy to test (mock one import)
- ✅ Consistent auth across your entire app
- ✅ Clean imports - no config needed at call site

## 📖 API Reference

### `Azure.config(config): void`

Set global configuration for all Azure services. Call once at app startup.

**Parameters:**

```typescript
interface AzureConfig {
  // Authentication options (choose one):
  accessToken?: string;              // Light mode: temporary token (~1 hour)
  refreshToken?: string;             // Medium mode: 90-day auto-renewal
  tokenProvider?: (callback: string) => Promise<string> | string; // Super mode: infinite renewal

  // Required credentials:
  clientId?: string;                 // Azure app client ID
  clientSecret?: string;             // Azure app client secret
  tenantId?: string;                 // Azure tenant ID

  // Optional:
  scopes?: string[];                 // Custom OAuth scopes
  allowInsecure?: boolean;           // Allow insecure SSL (dev only)
}
```

**Examples:**

```typescript
// Medium User: Refresh token (90-day renewal)
Azure.config({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});

// Super User: Token provider (infinite renewal)
Azure.config({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  tokenProvider: async (callback) => {
    // Your custom auth logic (e.g., Playwright, browser automation)
    return await getAuthorizationCode(callback);
  },
});

// Light User: Access token only (temporary)
Azure.config({
  accessToken: "eyJ0eX...", // Expires in ~1 hour
});
```

### Service Getters

After calling `Azure.config()`, access services directly:

- **`Azure.outlook`** - Email operations
- **`Azure.teams`** - Teams messaging and adaptive cards
- **`Azure.calendar`** - Calendar and holidays
- **`Azure.sharePoint`** - SharePoint lists and sites

### `Azure.reset(): void`

Reset global configuration and clear all service instances. Useful for testing or switching accounts.

```typescript
Azure.reset();
Azure.config({ refreshToken: "new-token" });
```

### `Azure.listStoredCredentials(): Promise<Array>`

List all stored credential files. Useful for debugging multi-tenant setups.

```typescript
const stored = await Azure.listStoredCredentials();
console.log(stored);
// [
//   { tenantId: 'abc123', clientId: 'def456', file: 'tokens.abc123.def456.json' },
//   { tenantId: 'xyz789', clientId: 'app2', file: 'tokens.xyz789.app2.json' },
//   { file: 'tokens.json' } // Legacy file
// ]
```

### `Azure.clearStoredCredentials(tenantId?, clientId?): Promise<void>`

Clear stored credentials for a specific tenant/client or all.

```typescript
// Clear specific tenant/client
await Azure.clearStoredCredentials("abc123", "def456");

// Clear all stored credentials
await Azure.clearStoredCredentials();
```

## 🔑 Token Management

### Token Priority

The utility loads tokens in this priority order:

1. **Memory** - Token provided in `init()` config
2. **Storage** - Saved token file (cross-platform)
3. **Provider** - Custom `tokenProvider` callback
4. **Error** - If none available, throws helpful error

### Storage Locations

Tokens are automatically saved to platform-specific locations with **multi-tenant support**:

**Storage Directory:**

- **Windows**: `C:\Users\{user}\AppData\Local\ms-graph-devtools\`
- **Mac/Linux**: `~/.config/ms-graph-devtools/`

**File Naming (Automatic):**

- With tenant+client: `tokens.{tenantId}.{clientId}.json`
- Legacy/fallback: `tokens.json`

**Example:**

```
~/.config/ms-graph-devtools/
  ├── tokens.abc123-tenant.def456-client.json
  ├── tokens.xyz789-tenant.app2-client.json
  └── tokens.json  (fallback)
```

**Benefits:**

- ✅ Multiple tenants/clients supported automatically
- ✅ No overwrites when switching configurations
- ✅ Backward compatible with single-tenant setups
- ✅ Files created with secure permissions (owner read/write only)

## 📋 Usage Examples

### Example 1: Global Configuration Pattern (Recommended)

```typescript
// config/azure.ts - Configure once, use everywhere
import Azure from 'ms-graph-devtools';

Azure.config({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});

export default Azure;
```

```typescript
// anywhere in your app
import Azure from './config/azure';

await Azure.outlook.sendMail({...});
await Azure.teams.postMessage({...});
```

### Example 2: Scheduled Automation

```typescript
// scheduled-task.ts
import Azure from './config/azure'; // Pre-configured instance

async function dailyReport() {
  const emails = await Azure.outlook.getMails(
    new Date().toISOString(),
    "Invoice"
  );

  await Azure.outlook.sendMail({
    message: {
      subject: "Daily Report",
      body: {
        contentType: "Text",
        content: `Found ${emails.length} invoices today`,
      },
      toRecipients: [
        {
          emailAddress: { address: "admin@example.com" },
        },
      ],
    },
  });
}

dailyReport();
```

**Add to cron:**

```bash
0 9 * * * cd /path/to/project && node scheduled-task.js
```

### Example 3: Token Provider with Playwright

```typescript
// config/azure.ts
import Azure from "ms-graph-devtools";
import Playwright from "./playwright-helper";

Azure.config({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  tokenProvider: async (callback) => {
    // Playwright automates browser to get auth code
    return await Playwright.getAzureCode(callback);
  },
});

export default Azure;
```

```typescript
// anywhere in your app
import Azure from './config/azure';

// Just works - tokenProvider handles auth automatically
await Azure.outlook.getMe();
await Azure.calendar.getCalendars();
```

### Example 4: Multi-Tenant Usage

```typescript
import Azure from "ms-graph-devtools";

// Company A
Azure.config({
  tenantId: "company-a-tenant-id",
  clientId: "app-1-client-id",
  clientSecret: process.env.COMPANY_A_SECRET,
  refreshToken: "company-a-token",
});
await Azure.outlook.getMe();
// Saved to: tokens.company-a-tenant-id.app-1-client-id.json

// Switch to Company B (reset first)
Azure.reset();
Azure.config({
  tenantId: "company-b-tenant-id",
  clientId: "app-2-client-id",
  clientSecret: process.env.COMPANY_B_SECRET,
  refreshToken: "company-b-token",
});
await Azure.teams.getTeams();
// Saved to: tokens.company-b-tenant-id.app-2-client-id.json

// Both token files are preserved!
// Switch between them by calling reset() and config()
```

### Example 5: CI/CD Environment

```typescript
// automation-script.js
import Azure from 'ms-graph-devtools';

// Load credentials from GitHub Secrets or environment
Azure.config({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});

await Azure.outlook.sendMail({...});
await Azure.teams.postMessage({...});
```

```yaml
# .github/workflows/automation.yml
steps:
  - run: node automation-script.js
    env:
      AZURE_REFRESH_TOKEN: ${{ secrets.AZURE_REFRESH_TOKEN }}
      AZURE_CLIENT_ID: ${{ secrets.AZURE_CLIENT_ID }}
      AZURE_CLIENT_SECRET: ${{ secrets.AZURE_CLIENT_SECRET }}
      AZURE_TENANT_ID: ${{ secrets.AZURE_TENANT_ID }}
```

## 📚 Service Classes

This library provides specialized service classes for different Microsoft Graph APIs.

**Two usage patterns:**

1. **Global instance (Recommended)**: `Azure.outlook`, `Azure.teams`, etc. after calling `Azure.config()`
2. **Direct instantiation**: `new Outlook()` for advanced use cases (multiple configs, dependency injection)

### 🔷 Outlook Service

Email operations with fluent builder pattern.

```typescript
import Azure from "ms-graph-devtools";

// Pattern 1: Via Azure global instance (recommended)
const user = await Azure.outlook.getMe();

// Get emails
const emails = await Azure.outlook.getMails("2024-01-15", "invoice");

// Send email with fluent builder
await Azure.outlook
  .compose()
  .subject("Meeting Reminder")
  .body("Don't forget our meeting tomorrow!", "Text")
  .to(["colleague@example.com"])
  .cc(["manager@example.com"])
  .importance("high")
  .send();

// With attachments
await Azure.outlook
  .compose()
  .subject("Monthly Report")
  .body("<h1>Report</h1>", "HTML")
  .to(["boss@example.com"])
  .attachments(["./report.pdf", "./charts.xlsx"])
  .send();

// Pattern 2: Direct instantiation (advanced)
import { Outlook } from "ms-graph-devtools";
const outlook = new Outlook({ refreshToken: "..." });
await outlook.sendMail({...});
```

**Available Methods:**

- `getMe()` - Get current user profile
- `sendMail(payload)` - Send email (low-level)
- `getMails(date, subjectFilter?)` - Get emails by date
- `compose()` - Create fluent email builder

**Email Builder Methods:**

- `subject(text)` - Set subject
- `body(content, type?)` - Set body (Text/HTML)
- `to(recipients[])` - Set recipients
- `cc(recipients[])` - Set CC
- `bcc(recipients[])` - Set BCC
- `attachments(files[])` - Add attachments
- `importance(level)` - Set priority (low/normal/high)
- `requestReadReceipt(bool)` - Request read receipt
- `send()` - Send the email

---

### 🔷 Teams Service

Teams messaging, adaptive cards, and channel management.

```typescript
import Azure from "ms-graph-devtools";

// Get user's teams
const myTeams = await Azure.teams.getTeams();
// [{ id: '...', displayName: 'Engineering Team', description: '...' }]

// Get channels for a team
const channels = await Azure.teams.getChannels(myTeams[0].id);
// [{ id: '...', displayName: 'General', membershipType: 'standard' }]

// Get tags for mentions
const tags = await Azure.teams.getTags(myTeams[0].id);
// [{ id: '...', displayName: 'Backend Team', memberCount: 5 }]

// Send adaptive card with builder
await Azure.teams
  .compose()
  .team("team-id")
  .channel("channel-id")
  .card({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "Deployment Complete!",
        size: "Large",
        weight: "Bolder",
      },
    ],
    actions: [
      {
        type: "Action.OpenUrl",
        title: "View Dashboard",
        url: "https://dashboard.example.com",
      },
    ],
  })
  .mentionTeam("team-id", "Engineering Team")
  .mentionUser("user-id", "John Doe")
  .send();
```

**Available Methods:**

- `getTeams()` - Get all teams user has joined
- `getChannels(teamId)` - Get channels in a team
- `getTags(teamId)` - Get tags for mentions
- `postAdaptiveCard(teamId, channelId, card, tags?)` - Post card (low-level)
- `compose()` - Create fluent adaptive card builder

**Adaptive Card Builder Methods:**

- `team(teamId)` - Set target team
- `channel(channelId)` - Set target channel
- `card(adaptiveCard)` - Set card JSON
- `mentionTeam(id, name)` - Add team mention
- `mentionTag(id, name)` - Add tag mention
- `mentionUser(id, name)` - Add user mention
- `mentions(tags[])` - Add multiple mentions
- `send()` - Post the card

---

### 🔷 SharePoint Service

SharePoint list operations and site management.

```typescript
import Azure from "ms-graph-devtools";

// Search for sites
const sites = await Azure.sharePoint.searchSites("Engineering");
Azure.sharePoint.setSiteId(sites[0].id);

// Get all lists in site
const lists = await Azure.sharePoint.getLists();
// [{ id: '...', displayName: 'Tasks', webUrl: '...' }]

// Get list columns
const columns = await Azure.sharePoint.getListColumns("Tasks");

// Create list item
await Azure.sharePoint.createListItem("Tasks", {
  Title: "New Task",
  Status: "Active",
  Priority: "High",
  DueDate: "2024-12-31",
});

// Query items with filtering
const items = await Azure.sharePoint.getListItems("Tasks", {
  filter: "fields/Status eq 'Active'",
  orderby: "createdDateTime desc",
  top: 10,
  expand: "fields",
});

// Update item
await Azure.sharePoint.updateListItem("Tasks", "item-id", {
  Status: "Completed",
});

// Delete item
await Azure.sharePoint.deleteListItem("Tasks", "item-id");

// Advanced: Query and process (task queue pattern)
const tasks = await Azure.sharePoint.queryAndProcess(
  "TaskQueue",
  "fields/taskName eq 'send-notification'",
  (item) => ({
    id: item.id,
    recipient: item.fields.recipient,
    message: item.fields.message,
  }),
  true // delete after processing
);

// Get latest item
const latest = await Azure.sharePoint.getLatestItem("Tasks");
```

**Available Methods:**

- `setSiteId(siteId)` - Set default site ID
- `searchSites(query)` - Search for sites
- `getSiteByPath(hostname, path)` - Get specific site
- `getLists(siteId?)` - Get all lists
- `getList(listId, siteId?)` - Get specific list
- `getListColumns(listId, siteId?)` - Get list columns
- `getListItems(listId, options?, siteId?)` - Query list items
- `getListItem(listId, itemId, expand?, siteId?)` - Get single item
- `createListItem(listId, fields, siteId?)` - Create item
- `updateListItem(listId, itemId, fields, siteId?)` - Update item
- `deleteListItem(listId, itemId, siteId?)` - Delete item
- `deleteListItems(listId, itemIds[], siteId?)` - Bulk delete
- `queryAndProcess(listId, filter, processor, deleteAfter?, siteId?)` - Query and process pattern
- `getLatestItem(listId, orderBy?, filter?, siteId?)` - Get most recent item

---

### 🔷 Calendar Service

Calendar and holiday management.

```typescript
import Azure from "ms-graph-devtools";

// Get all calendars
const calendars = await Azure.calendar.getCalendars();

// Get holidays
const indiaHolidays = await Azure.calendar.getIndiaHolidays(
  "2024-01-01",
  "2024-12-31"
);
const japanHolidays = await Azure.calendar.getJapanHolidays(
  "2024-01-01",
  "2024-12-31"
);
```

**Available Methods:**

- `getCalendars()` - Get user's calendars
- `getIndiaHolidays(start, end)` - Get India holidays
- `getJapanHolidays(start, end)` - Get Japan holidays

---

## 🎯 Service Usage Patterns

### Pattern 1: Global Instance (Recommended)

```typescript
// config/azure.ts
import Azure from 'ms-graph-devtools';

Azure.config({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});

export default Azure;
```

```typescript
// anywhere in your app
import Azure from './config/azure';

await Azure.outlook.compose().subject('Test').to(['user@example.com']).send();
await Azure.teams.compose().team('id').channel('id').card({...}).send();
await Azure.sharePoint.createListItem('Tasks', { Title: 'New Task' });
```

### Pattern 2: Direct Service Instantiation (Advanced)

For cases where you need multiple configurations or dependency injection:

```typescript
import { Outlook, Teams } from "ms-graph-devtools";

const outlook = new Outlook({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});

const teams = new Teams(); // Uses global config
```

### Pattern 3: Shared Authentication

```typescript
import { AzureAuth, Outlook, Teams } from "ms-graph-devtools";

// Create shared auth instance
const auth = new AzureAuth({ refreshToken: "your-token" });

// Share across services
const outlook = new Outlook(auth);
const teams = new Teams(auth);

await outlook.getMe();
await teams.getTeams();
```

### Pattern 4: Task Automation

```typescript
import Azure from './config/azure';

// Process task queue
const tasks = await Azure.sharePoint.queryAndProcess(
  "NotificationQueue",
  "fields/status eq 'pending'",
  async (item) => {
    // Send notification to Teams
    await Azure.teams
      .compose()
      .team(item.fields.teamId)
      .channel(item.fields.channelId)
      .card(JSON.parse(item.fields.cardData))
      .send();

    return { id: item.id, processed: true };
  },
  true // delete after processing
);

console.log(`Processed ${tasks.length} notifications`);
```

## 💡 Best Practices

### 1. Use Global Instance Export Pattern

**The recommended way to use this library:**

```typescript
// ✅ GOOD: Single config file pattern
// config/azure.ts
import Azure from 'ms-graph-devtools';

Azure.config({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  tokenProvider: async (callback) => await getAuthCode(callback),
});

export default Azure;
```

```typescript
// Everywhere else in your app
import Azure from './config/azure';
await Azure.outlook.sendMail({...});
```

**Why this is best:**
- Single source of truth for configuration
- No config duplication across files
- Easier to maintain and update credentials
- Consistent authentication across entire app
- Simpler imports (just the instance, no config needed)

**Avoid this pattern:**

```typescript
// ❌ BAD: Repeating config everywhere
import { Outlook } from 'ms-graph-devtools';

const outlook = new Outlook({
  clientId: process.env.AZURE_CLIENT_ID,  // Repeated in every file
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});
```

### 2. Keep Config Separate from Business Logic

**File structure:**
```
your-project/
├── config/
│   └── azure.ts          # Azure config + export
├── src/
│   ├── email.ts          # import Azure from '../config/azure'
│   ├── notifications.ts  # import Azure from '../config/azure'
│   └── reports.ts        # import Azure from '../config/azure'
└── .env                  # Credentials (gitignored)
```

### 3. Use Environment Variables

Always load credentials from environment variables, never hardcode:

```typescript
// config/azure.ts
import Azure from 'ms-graph-devtools';
import 'dotenv/config'; // If using dotenv

Azure.config({
  clientId: process.env.AZURE_CLIENT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!,
  tenantId: process.env.AZURE_TENANT_ID!,
  tokenProvider: async (callback) => await getAuthCode(callback),
});

export default Azure;
```

### 4. Type Safety with TypeScript

```typescript
// config/azure.ts
import Azure from 'ms-graph-devtools';

// Validate required env vars at startup
const requiredEnvVars = ['AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_TENANT_ID'];
for (const envVar of requiredEnvVars) {
  if (!process.env[envVar]) {
    throw new Error(`Missing required environment variable: ${envVar}`);
  }
}

Azure.config({
  clientId: process.env.AZURE_CLIENT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!,
  tenantId: process.env.AZURE_TENANT_ID!,
  tokenProvider: async (callback) => await getAuthCode(callback),
});

export default Azure;
```

## 🔒 Security Best Practices

### 1. Never Commit Tokens

Add to `.gitignore`:

```
.env
tokens.json
*.token
```

### 2. Use Secrets in CI/CD

Store credentials as GitHub Secrets or CI/CD environment variables and pass them explicitly to the library:

```typescript
const outlook = new Outlook({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});
```

### 3. Secure File Permissions

The utility automatically sets secure permissions on token files:

- Directory: `0o700` (owner only)
- File: `0o600` (owner read/write only)

### 4. Token Rotation

Refresh tokens typically expire after 90 days. Set a reminder to rotate them.

### 5. Least Privilege

Only request the scopes you need in your Azure App Registration.

## 🛠️ Getting Your Refresh Token

See [get-refresh-token.js](./get-refresh-token.js) for an interactive tool to obtain your refresh token.

```bash
node get-refresh-token.js
```

Or see [REFRESH_TOKEN_SETUP.md](./REFRESH_TOKEN_SETUP.md) for manual instructions.

## ❓ FAQ

**Q: What's the recommended way to use this library?**

A: **Export a configured global instance** - this is the best practice for 95% of use cases:

```typescript
// config/azure.ts - Configure once
import Azure from 'ms-graph-devtools';
Azure.config({...});
export default Azure;

// everywhere else - just import and use
import Azure from './config/azure';
await Azure.outlook.sendMail({...});
```

This gives you:
- Single source of truth
- No config duplication
- Clean imports everywhere
- Easy to maintain

**Q: Do I need to call `Azure.config()` every time I use a service?**

A: No! That's the beauty of the global instance pattern. Call `Azure.config()` once in a config file, export the instance, then import and use it everywhere.

```typescript
// config/azure.ts - Call config() ONCE
Azure.config({...});
export default Azure;

// other files - NO config() needed, just import
import Azure from './config/azure';
await Azure.outlook.sendMail({...});
await Azure.teams.postMessage({...});
```

**Q: What if I call `config()` multiple times with different tokens?**

A: Each call to `config()` replaces the previous configuration and resets all service instances. Use `Azure.reset()` first for clarity.

**Q: Where are tokens stored?**

A: Platform-specific secure locations:

- Windows: `%LOCALAPPDATA%\ms-graph-devtools\tokens.json`
- Mac/Linux: `~/.config/ms-graph-devtools/tokens.json`

**Q: Can I use this in serverless/Lambda?**

A: Yes! Use a custom `tokenProvider` that fetches from your secret store (AWS Secrets Manager, etc.).

**Q: Will this work with my organization's SSO (Okta, 1Password, etc.)?**

A: Yes! You obtain the refresh token once through your SSO, then the utility handles everything automatically.

**Q: How do I update the token?**

A: Either:

1. `Azure.reset()` then `Azure.config({ refreshToken: 'new-token' })`
2. Delete the storage file and call `Azure.config()` with new credentials

## 🐛 Troubleshooting

### "No refresh token available"

**Solution:** Provide a token via `Azure.config()`:

```typescript
Azure.config({ refreshToken: 'xxx' })
// OR
Azure.config({ tokenProvider: async (callback) => getAuthCode(callback) })
// OR
Azure.config({ accessToken: 'xxx' }) // Light mode
```

### "Invalid grant" error

Your refresh token has expired (typically 90 days). Get a new one:

```bash
node get-refresh-token.js
```

### "Insufficient privileges"

Your Azure App Registration needs API permissions. Go to:
Azure Portal → App Registrations → API permissions → Add permissions

### Storage file not found

Normal on first run. The file is created when you first provide a token.

## 📦 For Package Publishers

When publishing to npm:

1. Remove `Playwright` dependency (user-specific auth)
2. Include `get-refresh-token.js` helper script
3. Document in README how users get their token
4. Add to `bin` in package.json:

```json
{
  "bin": {
    "get-refresh-token": "./get-refresh-token.js"
  }
}
```

Users can then run:

```bash
npx your-package get-refresh-token
```

## 📄 License

MIT

## 🤝 Contributing

Contributions welcome! Please open an issue first for major changes.

---

Built with ❤️ for automated Microsoft Graph workflows
