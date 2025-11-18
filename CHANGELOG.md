# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
