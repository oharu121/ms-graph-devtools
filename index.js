/**
 * ms-graph-devtools - Microsoft Azure/Graph API utility with modular service architecture
 *
 * @license ISC
 */

"use strict";

// Default export (Azure class)
module.exports = require("./dist/index").default;

// Named exports (service classes)
module.exports.Outlook = require("./dist/index").Outlook;
module.exports.Calendar = require("./dist/index").Calendar;
module.exports.Teams = require("./dist/index").Teams;
module.exports.SharePoint = require("./dist/index").SharePoint;

// Builder classes
module.exports.MailBuilder = require("./dist/index").MailBuilder;
module.exports.AdaptiveCardBuilder = require("./dist/index").AdaptiveCardBuilder;

// Auth class
module.exports.AzureAuth = require("./dist/index").AzureAuth;

// Re-export default as named export for consistency
module.exports.Azure = module.exports.default;
