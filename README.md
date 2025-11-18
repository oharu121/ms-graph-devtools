# Azure Microsoft Graph Utility

[![npm version](https://badge.fury.io/js/ms-graph-devtools.svg)](https://badge.fury.io/js/ms-graph-devtools)
![License](https://img.shields.io/npm/l/ms-graph-devtools)
![Types](https://img.shields.io/npm/types/ms-graph-devtools)
![NPM Downloads](https://img.shields.io/npm/dw/ms-graph-devtools)
![Last Commit](https://img.shields.io/github/last-commit/oharu121/ms-graph-devtools)
![GitHub Stars](https://img.shields.io/github/stars/oharu121/ms-graph-devtools?style=social)

A TypeScript utility for Microsoft Graph API operations with automatic token management and cross-platform storage support.

## ✨ Features

- ✅ **Singleton pattern** - One instance shared across your entire application
- ✅ **Automatic token refresh** - Handles token expiration transparently
- ✅ **Multiple token sources** - Direct config, storage, or custom provider
- ✅ **Cross-platform storage** - Works on Windows, Mac, and Linux using XDG standards
- ✅ **Idempotent initialization** - Safe to call `init()` multiple times
- ✅ **TypeScript support** - Full type safety and autocomplete
- ✅ **Method chaining** - Clean, fluent API

## Major Achievement: Three-Tier User System

Successfully designed and implemented a progressive complexity model for Azure authentication that supports all user types from beginners to enterprise.

### Tier 1: Light User (Access Token Only)

**Purpose:** Quick testing, POC, experimentation

**API:**

```typescript
Azure.setAccessToken("eyJ0eX...").getMe();
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
Azure.init({
  refreshToken: "xxx",
  clientId: "...",
  clientSecret: "...",
  tenantId: "...",
});
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
Azure.init({
  clientId: "...",
  clientSecret: "...",
  tenantId: "...",
  tokenProvider: async () => await vault.get("token"),
});
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
npm install your-package-name
```

### Basic Usage

```typescript
import Azure from 'ms-graph-devtools';

// Simple usage - loads token from storage
await Azure.init().getMe();

// Provide token directly
await Azure.init({ refreshToken: 'your-token' }).sendMail({...});

// Load from storage
await Azure.init().getCalendars();
```

## 📖 API Reference

### `Azure.init(config?): Azure`

Initialize the Azure singleton instance. This method is **idempotent** - calling it multiple times will only apply the first configuration.

**Parameters:**

```typescript
interface AzureConfig {
  refreshToken?: string; // Direct token
  tokenProvider?: () => Promise<string> | string; // Custom token provider
  clientId?: string; // Azure app client ID
  clientSecret?: string; // Azure app client secret
  tenantId?: string; // Azure tenant ID
}
```

**Returns:** `Azure` instance for method chaining

**Examples:**

```typescript
// Load from storage (most common in production)
const azure = Azure.init();

// Provide token directly
const azure = Azure.init({ refreshToken: "xxx" });

// Custom token provider
const azure = Azure.init({
  tokenProvider: async () => {
    return await mySecureVault.get("azure-token");
  },
});

// Full configuration
const azure = Azure.init({
  refreshToken: "xxx",
  clientId: "your-client-id",
  clientSecret: "your-client-secret",
  tenantId: "your-tenant-id",
});
```

### `Azure.reset(): void`

Reset the singleton instance. Useful for testing or when you need to reinitialize with different credentials.

```typescript
Azure.reset();
Azure.init({ refreshToken: "new-token" });
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

### Example 1: Production Usage (Storage)

```typescript
import Azure from 'ms-graph-devtools';

// First time setup - provide token
Azure.init({ refreshToken: 'your-token' });
// Token is saved to storage

// All subsequent runs - just works
await Azure.init().getMe();
await Azure.init().sendMail({...});
```

### Example 2: Scheduled Automation

```typescript
// scheduled-task.ts
import Azure from "ms-graph-devtools";

async function dailyReport() {
  // Loads token from storage automatically
  const emails = await Azure.init().getMails(
    new Date().toISOString(),
    "Invoice"
  );

  await Azure.init().sendMail({
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

### Example 3: Custom Token Provider

```typescript
import Azure from "ms-graph-devtools";
import { SecretVault } from "my-vault";

const vault = new SecretVault();

await Azure.init({
  tokenProvider: async () => {
    return await vault.getSecret("azure-refresh-token");
  },
}).getCalendars();
```

### Example 4: Multiple Scripts (Singleton)

```typescript
// script1.js
import Azure from 'ms-graph-devtools';
Azure.init({ refreshToken: 'xxx' });
await Azure.init().getMe();

// script2.js (different file)
import Azure from 'ms-graph-devtools';
// Same instance! No need to provide token again
await Azure.init().sendMail({...});
```

### Example 5: Multi-Tenant Usage

```typescript
import Azure from "ms-graph-devtools";

// Company A
Azure.init({
  tenantId: "company-a-tenant-id",
  clientId: "app-1-client-id",
  clientSecret: process.env.COMPANY_A_SECRET,
  refreshToken: "company-a-token",
});
await Azure.init().getMe();
// Saved to: tokens.company-a-tenant-id.app-1-client-id.json

// Switch to Company B (reset first)
Azure.reset();
Azure.init({
  tenantId: "company-b-tenant-id",
  clientId: "app-2-client-id",
  clientSecret: process.env.COMPANY_B_SECRET,
  refreshToken: "company-b-token",
});
await Azure.init().getMe();
// Saved to: tokens.company-b-tenant-id.app-2-client-id.json

// Both token files are preserved!
// You can switch between them by calling reset() and init()
```

### Example 5: CI/CD Environment

```typescript
// automation-script.js
import Azure from 'ms-graph-devtools';

// Load credentials from GitHub Secrets or environment
const outlook = new Outlook({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});

await outlook.sendMail({...});
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

This library provides specialized service classes for different Microsoft Graph APIs. Each service can be used independently or through the main Azure class.

### 🔷 Outlook Service

Email operations with fluent builder pattern.

```typescript
import { Outlook } from "smart-azure";

const outlook = new Outlook();

// Get current user
const user = await outlook.getMe();

// Get emails
const emails = await outlook.getMails("2024-01-15", "invoice");

// Send email with fluent builder
await outlook
  .compose()
  .subject("Meeting Reminder")
  .body("Don't forget our meeting tomorrow!", "Text")
  .to(["colleague@example.com"])
  .cc(["manager@example.com"])
  .importance("high")
  .send();

// With attachments
await outlook
  .compose()
  .subject("Monthly Report")
  .body("<h1>Report</h1>", "HTML")
  .to(["boss@example.com"])
  .attachments(["./report.pdf", "./charts.xlsx"])
  .send();
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
import { Teams } from "smart-azure";

const teams = new Teams();

// Get user's teams
const myTeams = await teams.getTeams();
// [{ id: '...', displayName: 'Engineering Team', description: '...' }]

// Get channels for a team
const channels = await teams.getChannels(myTeams[0].id);
// [{ id: '...', displayName: 'General', membershipType: 'standard' }]

// Get tags for mentions
const tags = await teams.getTags(myTeams[0].id);
// [{ id: '...', displayName: 'Backend Team', memberCount: 5 }]

// Send adaptive card with builder
await teams
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
import { SharePoint } from "smart-azure";

const sharepoint = new SharePoint();

// Search for sites
const sites = await sharepoint.searchSites("Engineering");
sharepoint.setSiteId(sites[0].id);

// Get all lists in site
const lists = await sharepoint.getLists();
// [{ id: '...', displayName: 'Tasks', webUrl: '...' }]

// Get list columns
const columns = await sharepoint.getListColumns("Tasks");

// Create list item
await sharepoint.createListItem("Tasks", {
  Title: "New Task",
  Status: "Active",
  Priority: "High",
  DueDate: "2024-12-31",
});

// Query items with filtering
const items = await sharepoint.getListItems("Tasks", {
  filter: "fields/Status eq 'Active'",
  orderby: "createdDateTime desc",
  top: 10,
  expand: "fields",
});

// Update item
await sharepoint.updateListItem("Tasks", "item-id", {
  Status: "Completed",
});

// Delete item
await sharepoint.deleteListItem("Tasks", "item-id");

// Advanced: Query and process (task queue pattern)
const tasks = await sharepoint.queryAndProcess(
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
const latest = await sharepoint.getLatestItem("Tasks");
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
import { Calendar } from "smart-azure";

const calendar = new Calendar();

// Get all calendars
const calendars = await calendar.getCalendars();

// Get holidays
const indiaHolidays = await calendar.getIndiaHolidays(
  "2024-01-01",
  "2024-12-31"
);
const japanHolidays = await calendar.getJapanHolidays(
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

### Pattern 1: Direct Service Usage

```typescript
import { Outlook, Teams, SharePoint } from 'smart-azure';

const outlook = new Outlook();
const teams = new Teams();
const sharepoint = new SharePoint();

await outlook.compose().subject('Test').to(['user@example.com']).send();
await teams.compose().team('id').channel('id').card({...}).send();
await sharepoint.createListItem('Tasks', { Title: 'New Task' });
```

### Pattern 2: Shared Authentication

```typescript
import { AzureAuth, Outlook, Teams } from "smart-azure";

// Create shared auth instance
const auth = new AzureAuth({ refreshToken: "your-token" });

// Share across services
const outlook = new Outlook(auth);
const teams = new Teams(auth);

await outlook.getMe();
await teams.getTeams();
```

### Pattern 3: Configuration-Based

```typescript
import { Outlook } from "smart-azure";

const outlook = new Outlook({
  clientId: process.env.AZURE_CLIENT_ID,
  clientSecret: process.env.AZURE_CLIENT_SECRET,
  tenantId: process.env.AZURE_TENANT_ID,
  refreshToken: process.env.AZURE_REFRESH_TOKEN,
});
```

### Pattern 4: Task Automation

```typescript
import { SharePoint, Teams } from "smart-azure";

const sharepoint = new SharePoint();
const teams = new Teams();

// Process task queue
const tasks = await sharepoint.queryAndProcess(
  "NotificationQueue",
  "fields/status eq 'pending'",
  async (item) => {
    // Send notification to Teams
    await teams
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

**Q: Do I need to call `init()` before every method?**

A: No! The singleton persists. But calling `init()` is safe and enables method chaining:

```typescript
await Azure.init().getMe(); // Clean syntax
```

**Q: What if I call `init()` multiple times with different tokens?**

A: The first configuration wins (idempotent). Use `Azure.reset()` to reinitialize.

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

1. `Azure.reset()` then `Azure.init({ refreshToken: 'new-token' })`
2. Delete the storage file and reinitialize with a new config

## 🐛 Troubleshooting

### "No refresh token available"

**Solution:** Provide a token via one of these methods:

```typescript
Azure.init({ refreshToken: 'xxx' })
// OR
Azure.init({ tokenProvider: async () => 'xxx' })
// OR load from storage (if previously saved)
Azure.init()
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
