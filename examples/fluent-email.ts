import Azure, { Outlook } from "../src/index";

/**
 * Example: Using the fluent email builder API
 *
 * This demonstrates the new compose() method which provides a
 * chainable, intuitive way to build and send emails.
 */

async function exampleFluentEmail() {
  // Set global configuration
  Azure.config({
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    tenantId: process.env.AZURE_TENANT_ID,
    refreshToken: process.env.AZURE_REFRESH_TOKEN,
  });

  const outlook = new Outlook();

  // Example 1: Simple text email
  await outlook
    .compose()
    .subject("Meeting Reminder")
    .body("Don't forget our meeting tomorrow at 10 AM!", "Text")
    .to(["colleague@example.com"])
    .send();

  console.log("✓ Simple email sent");

  // Example 2: HTML email with multiple recipients
  await outlook
    .compose()
    .subject("Project Update")
    .body(
      `
      <h1>Q4 Project Update</h1>
      <p>Here's the latest update on our project:</p>
      <ul>
        <li>Phase 1: Complete ✓</li>
        <li>Phase 2: In Progress</li>
        <li>Phase 3: Scheduled</li>
      </ul>
    `,
      "HTML"
    )
    .to(["team@example.com"])
    .cc(["manager@example.com"])
    .importance("high")
    .send();

  console.log("✓ HTML email sent");

  // Example 3: Email with attachments
  await outlook
    .compose()
    .subject("Monthly Report")
    .body("Please find attached the monthly report.", "Text")
    .to(["boss@example.com"])
    .attachments(["./report.pdf", "./charts.xlsx"])
    .importance("high")
    .requestReadReceipt(true)
    .send();

  console.log("✓ Email with attachments sent");

  // Example 4: Email with Buffer attachments (no file system)
  const jsonData = Buffer.from(JSON.stringify({ status: "complete" }, null, 2));

  await outlook
    .compose()
    .subject("API Response Data")
    .body("Here's the response data from the API.", "Text")
    .to(["developer@example.com"])
    .attachments([
      {
        name: "response.json",
        content: jsonData,
        contentType: "application/json",
      },
    ])
    .send();

  console.log("✓ Email with Buffer attachment sent");

  // Example 5: Full-featured email
  await outlook
    .compose()
    .subject("Important: Action Required")
    .bodyPreview("This email requires your immediate attention")
    .body(
      `
      <div style="font-family: Arial, sans-serif;">
        <h2 style="color: #d32f2f;">Action Required</h2>
        <p>Please review and approve the following items by EOD:</p>
        <ol>
          <li>Budget proposal</li>
          <li>Marketing campaign</li>
          <li>Q4 roadmap</li>
        </ol>
        <p><strong>Deadline:</strong> Today, 5:00 PM</p>
      </div>
    `,
      "HTML"
    )
    .to(["approver@example.com"])
    .cc(["supervisor@example.com"])
    .bcc(["archive@example.com"])
    .replyTo(["noreply@example.com"])
    .importance("high")
    .categories(["Action Required", "Urgent"])
    .flag()
    .requestReadReceipt(true)
    .saveToSentItems(true)
    .send();

  console.log("✓ Full-featured email sent");
}

// Comparison: Old way vs New way
async function comparisonExample() {
  const outlook = new Outlook();

  // OLD WAY: Manual payload construction
  await outlook.sendMail({
    message: {
      subject: "Hello",
      body: {
        contentType: "Text",
        content: "World",
      },
      toRecipients: [
        { emailAddress: { address: "user@example.com" } },
      ],
    },
  });

  // NEW WAY: Fluent builder (much cleaner!)
  await outlook
    .compose()
    .subject("Hello")
    .body("World", "Text")
    .to(["user@example.com"])
    .send();
}

// Run examples
if (require.main === module) {
  exampleFluentEmail().catch(console.error);
}

export { exampleFluentEmail, comparisonExample };
