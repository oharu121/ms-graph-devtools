/**
 * ms-graph-devtools - Microsoft Azure/Graph API utility with modular service architecture
 *
 * @license ISC
 */

// Re-export everything from the compiled dist
export { default } from './dist/index.js';
export {
  Outlook,
  Calendar,
  Teams,
  SharePoint,
  MailBuilder,
  AdaptiveCardBuilder,
  AzureAuth,
} from './dist/index.js';

// Re-export default as named export for consistency
export { default as Azure } from './dist/index.js';
