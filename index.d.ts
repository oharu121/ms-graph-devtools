/**
 * ms-graph-devtools - Microsoft Azure/Graph API utility with modular service architecture
 *
 * @license ISC
 */

// Default export (Azure class)
export { default } from "./dist/index.js";
export { default as Azure } from "./dist/index.js";

// Named exports (service classes)
export { Outlook, Calendar, Teams, SharePoint, AzureAuth } from "./dist/index.js";

// Builder classes
export { MailBuilder, AdaptiveCardBuilder } from "./dist/index.js";

// Export types
export type { AzureConfig } from "./dist/types.js";
