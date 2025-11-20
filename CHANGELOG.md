# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.0]

### Added
- **AWS SDK-Style Global Instance API**
  - Added lazy-loaded service getters: `Azure.outlook`, `Azure.teams`, `Azure.calendar`, `Azure.sharePoint`
  - Services are now accessible directly after calling `Azure.config()` - no manual instantiation needed
  - Singleton pattern: same instance returned on multiple accesses
  - Lazy initialization: services only created when first accessed
  - `Azure.reset()` now clears all service instances for clean state

### Changed
- **⚠️ BREAKING: Dropped Node.js 18 support**
  - Minimum Node.js version is now 20.0.0
  - Node.js 18 reached End-of-Life on April 30, 2025
  - This change enables better compatibility with modern dependencies
- **⚠️ API Improvement: Simplified service access pattern**
  - **Before**: `Azure.config({...}); const outlook = new Outlook(); await outlook.sendMail({...});`
  - **After**: `Azure.config({...}); await Azure.outlook.sendMail({...});`
  - No more manual service instantiation required
  - Global config is automatically applied to all services
  - Calling `Azure.config()` resets all service instances to pick up new configuration
  - Old pattern (`new Outlook()`) still works for advanced use cases

### Documentation
- Complete README rewrite to showcase new API pattern
- Updated all examples to use `Azure.service` syntax
- Added comprehensive vitest test suite for new API (`tests/azure-api.test.ts`)
- Updated Three-Tier User System examples
- Updated Service Usage Patterns section
- Clarified FAQ and troubleshooting for new API

### Migration Guide

**Old Pattern (Still Supported):**
```typescript
import { Outlook, Teams } from 'ms-graph-devtools';

const outlook = new Outlook({
  clientId: "...",
  clientSecret: "...",
  tenantId: "...",
  refreshToken: "..."
});

await outlook.sendMail({...});
```

**New Pattern (Recommended):**
```typescript
import Azure from 'ms-graph-devtools';

// Configure once
Azure.config({
  clientId: "...",
  clientSecret: "...",
  tenantId: "...",
  refreshToken: "..."
});

// Export for use across your app
export default Azure;

// Use services directly - no instantiation needed!
await Azure.outlook.sendMail({...});
await Azure.teams.postMessage({...});
await Azure.calendar.getCalendars();
await Azure.sharePoint.getLists();
```

**Benefits:**
- ✅ Configure once, use everywhere
- ✅ No redundant `init()` or instantiation
- ✅ Clean, AWS SDK-like API
- ✅ Better discoverability (IDE autocomplete shows all services)
- ✅ Still tree-shakable (unused services won't be bundled)

## [1.1.1]

### Changed
- **⚠️ BREAKING: Updated tokenProvider signature**
  - `tokenProvider` now receives the OAuth callback URL as a parameter
  - Old: `tokenProvider?: () => Promise<string> | string`
  - New: `tokenProvider?: (callback: string) => Promise<string> | string`
  - The callback URL contains all OAuth parameters (tenantId, clientId, scopes, redirectUri)
  - tokenProvider should return the authorization code (not refresh token)
  - Library now handles the code-to-token exchange internally via `forgeRefreshToken()`

### Added
- **OAuth Authorization Code Flow**: New `forgeRefreshToken()` method
  - Constructs OAuth authorization URL with configured parameters
  - Exchanges authorization code for access_token, refresh_token, and expires_in
  - Enables full Playwright integration for automated authentication

### Migration Guide

If you're using tokenProvider, update your implementation:

**Before:**
```typescript
const outlook = new Outlook({
  tokenProvider: async () => {
    // Return refresh token from vault
    return await fetchFromVault('azure-refresh-token');
  }
});
```

**After:**
```typescript
const outlook = new Outlook({
  tokenProvider: async (callback: string) => {
    // callback contains the OAuth authorization URL
    // Use Playwright to navigate, login, and extract the code
    const code = await Playwright.getAzureCode(callback);
    return code;
  }
});
```

## [1.1.0]

### Added
- **ESLint Configuration**: Full ESLint setup with TypeScript support
  - Modern flat config format (`eslint.config.js`)
  - TypeScript-specific linting rules
  - Prettier integration to prevent conflicts
  - Special rules for test files (allows `any` for mocking)
  - Excludes examples directory from linting
  - npm scripts: `npm run lint` and `npm run lint:fix`

### Changed
- **⚠️ BREAKING: Converted package to pure ESM (ES Modules)**
  - Added `"type": "module"` to package.json
  - Updated TypeScript configuration to output ESNext modules
  - Changed `module` from `"commonjs"` to `"ESNext"` in tsconfig.json
  - Changed `moduleResolution` from `"node"` to `"bundler"` in tsconfig.json
  - All imports now require `.js` extensions in distributed code
  - Removed CommonJS exports (no more `require()` support)
  - Updated `index.js` and `index.d.ts` to use ESM syntax
  - Package now only supports ESM consumers (Node.js 14+ with ESM)

- **Code Quality Improvements**:
  - Replaced `any` types with proper type annotations where possible
  - Added `AxonError` type imports from axios-fluent for better error handling
  - Removed unused imports and variables throughout codebase
  - Fixed unnecessary try-catch wrappers
  - Fixed unnecessary escape characters in template literals
  - Improved type safety in error handling using `NodeJS.ErrnoException`
  - All source files now pass strict ESLint checks

### Migration Guide

If you're upgrading from a CommonJS version:

**Before (CommonJS):**
```javascript
const { Outlook, Teams } = require('ms-graph-devtools');
```

**After (ESM):**
```javascript
import { Outlook, Teams } from 'ms-graph-devtools';
```

**For TypeScript users:**
- Ensure your `tsconfig.json` has `"module": "ESNext"` or `"NodeNext"`
- Ensure your `package.json` has `"type": "module"`

**For Node.js users:**
- Requires Node.js 14+ with ESM support
- Add `"type": "module"` to your `package.json`, OR
- Use `.mjs` file extensions for your entry points

## [1.0.0] - Initial Release

### Added
- Initial release with Microsoft Graph API utilities
- Outlook service for email operations
- Teams service for messaging and adaptive cards
- SharePoint service for list and site management
- Calendar service for calendar and holiday operations
- Automatic token management and refresh
- Fluent builder APIs for composing messages
- Comprehensive TypeScript type definitions
