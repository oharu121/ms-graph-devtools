/**
 * Advanced Usage Examples
 * This file demonstrates advanced patterns and use cases
 */

import { AzureAuth, Outlook, Calendar, Teams, SharePoint } from '../src/index';

// ============================================
// Example 1: Shared Auth Instance
// ============================================
async function example1_SharedAuth() {
  // Create a single auth instance
  const auth = new AzureAuth({
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    tenantId: process.env.AZURE_TENANT_ID,
    refreshToken: process.env.AZURE_REFRESH_TOKEN
  });

  // Share auth across multiple services
  // This ensures all services use the same token
  const outlook = new Outlook(auth);
  const calendar = new Calendar(auth);
  const teams = new Teams(auth);
  const sharepoint = new SharePoint(auth);

  // All API calls use the same authenticated session
  await outlook.compose().subject('Test').to(['test@example.com']).send();
  await calendar.getCalendars();

  const myTeams = await teams.getTeams();
  if (myTeams.length > 0) {
    const channels = await teams.getChannels(myTeams[0].id);
    await teams.compose()
      .team(myTeams[0].id)
      .channel(channels[0].id)
      .card({ type: 'AdaptiveCard', version: '1.4', body: [] })
      .send();
  }

  const sites = await sharepoint.searchSites('Engineering');
  console.log('Sites:', sites);
}

// ============================================
// Example 2: Token Provider (Super User Mode)
// ============================================
async function example2_TokenProvider() {
  // For infinite token renewal
  // Useful when integrating with external token vaults
  const outlook = new Outlook({
    clientId: '...',
    clientSecret: '...',
    tenantId: '...',
    tokenProvider: async () => {
      // Fetch token from your secure vault
      // This gets called automatically when token needs refresh
      const token = await fetchFromVault('azure-refresh-token');
      return token;
    }
  });

  await outlook.compose().subject('Test').to(['test@example.com']).send();
}

// Simulated vault function
async function fetchFromVault(key: string): Promise<string> {
  // In production, this would fetch from your secure vault
  // (Azure Key Vault, HashiCorp Vault, AWS Secrets Manager, etc.)
  return process.env.AZURE_REFRESH_TOKEN || '';
}

// ============================================
// Example 3: Custom OAuth Scopes
// ============================================
async function example3_CustomScopes() {
  // If you need admin-level permissions like Sites.ReadWrite.All
  const sharepoint = new SharePoint({
    clientId: '...',
    clientSecret: '...',
    tenantId: '...',
    refreshToken: '...',
    scopes: [
      'User.Read',
      'Mail.Send',
      'Sites.ReadWrite.All', // ⚠️ Requires admin consent
      'Files.ReadWrite.All'
    ]
  });

  await sharepoint.getList('Documents');
}

// ============================================
// Example 4: SharePoint Task Queue Processing
// ============================================
async function example4_TaskQueueProcessing() {
  const sharepoint = new SharePoint();
  const teams = new Teams();

  // Search and set site
  const sites = await sharepoint.searchSites('MyProject');
  if (sites.length > 0) {
    sharepoint.setSiteId(sites[0].id);
  }

  // Process notification queue
  const notifications = await sharepoint.queryAndProcess(
    'NotificationQueue',
    "fields/status eq 'pending'",
    async (item) => {
      // Send notification to Teams
      await teams.compose()
        .team(item.fields.teamId)
        .channel(item.fields.channelId)
        .card(JSON.parse(item.fields.cardData))
        .send();

      return { id: item.id, sent: true };
    },
    true // Delete after processing
  );

  console.log(`Processed ${notifications.length} notifications`);
}

// ============================================
// Example 5: Deployment Notification Workflow
// ============================================
async function example5_DeploymentNotification() {
  const sharepoint = new SharePoint();
  const teams = new Teams();

  // Get deployment info from SharePoint
  const sites = await sharepoint.searchSites('DevOps');
  if (sites.length > 0) {
    sharepoint.setSiteId(sites[0].id);
  }

  const deployment = await sharepoint.getLatestItem('Deployments');

  if (deployment) {
    // Notify team
    const myTeams = await teams.getTeams();
    const channels = await teams.getChannels(myTeams[0].id);

    await teams.compose()
      .team(myTeams[0].id)
      .channel(channels[0].id)
      .card({
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
          {
            type: 'TextBlock',
            text: `Deployment ${deployment.fields.Version} Complete!`,
            size: 'Large',
            weight: 'Bolder',
            color: 'Good'
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Environment:', value: deployment.fields.Environment },
              { title: 'Build:', value: deployment.fields.BuildNumber },
              { title: 'Status:', value: deployment.fields.Status }
            ]
          }
        ]
      })
      .mentionTeam(myTeams[0].id, myTeams[0].displayName)
      .send();
  }
}

// ============================================
// Example 6: Multi-Calendar Holiday Aggregation
// ============================================
async function example6_MultiCalendarHolidays() {
  const calendar = new Calendar();

  // Get holidays from multiple regions with custom calendar names
  const allHolidays = await Promise.all([
    calendar.getIndiaHolidays('2024-01-01', '2024-12-31'),
    calendar.getIndiaHolidays('2024-01-01', '2024-12-31', 'Indian Holidays'), // Custom name
    calendar.getJapanHolidays('2024-01-01', '2024-12-31'),
    calendar.getJapanHolidays('2024-01-01', '2024-12-31', ['Japan holidays', '日本の休日']),
    calendar.getHolidaysByCalendarName('US Holidays', '2024-01-01', '2024-12-31'),
    calendar.getHolidaysByCalendarName('UK Holidays', '2024-01-01', '2024-12-31')
  ]);

  // Flatten and deduplicate
  const flatHolidays = allHolidays.flat().filter(Boolean);
  console.log(`Total holidays: ${flatHolidays.length}`);
}

// ============================================
// Example 7: SharePoint Multi-Site Management
// ============================================
async function example7_MultiSiteManagement() {
  const sharepoint = new SharePoint();

  // Work with multiple sites
  const sites = await sharepoint.searchSites('Project');

  for (const site of sites) {
    console.log(`\nProcessing site: ${site.displayName}`);

    // Get lists for this site
    const lists = await sharepoint.getLists(site.id);
    console.log(`  Lists: ${lists.map(l => l.displayName).join(', ')}`);

    // Query tasks in this site
    const tasks = await sharepoint.getListItems(
      'Tasks',
      {
        filter: "fields/Status eq 'Active'",
        top: 5
      },
      site.id
    );
    console.log(`  Active tasks: ${tasks.length}`);
  }
}

// ============================================
// Example 8: Email Automation with Calendar
// ============================================
async function example8_EmailAutomation() {
  const outlook = new Outlook();
  const calendar = new Calendar();

  // Get upcoming holidays
  const holidays = await calendar.getHolidaysByCalendarName(
    'Company Holidays',
    new Date().toISOString(),
    new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString() // Next 30 days
  );

  if (holidays.length > 0) {
    const nextHoliday = holidays[0];

    // Send reminder email
    await outlook.compose()
      .subject(`Upcoming Holiday: ${nextHoliday.name}`)
      .body(`
        <h2>Holiday Reminder</h2>
        <p>Our next holiday is coming up:</p>
        <ul>
          <li><strong>Holiday:</strong> ${nextHoliday.name}</li>
          <li><strong>Date:</strong> ${new Date(nextHoliday.date).toLocaleDateString()}</li>
        </ul>
        <p>Please plan your work accordingly.</p>
      `, 'HTML')
      .to(['team@example.com'])
      .importance('high')
      .send();
  }
}

// ============================================
// Example 9: Error Handling & Retry Logic
// ============================================
async function retryOperation<T>(
  operation: () => Promise<T>,
  maxRetries: number = 3,
  delayMs: number = 1000
): Promise<T> {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await operation();
    } catch (error: any) {
      if (i === maxRetries - 1) throw error;

      // Check if it's a retryable error
      const isRetryable =
        error.response?.status === 429 || // Rate limit
        error.response?.status >= 500;    // Server error

      if (isRetryable) {
        const delay = delayMs * Math.pow(2, i); // Exponential backoff
        console.log(`Retry ${i + 1}/${maxRetries} after ${delay}ms`);
        await new Promise(resolve => setTimeout(resolve, delay));
      } else {
        throw error; // Don't retry client errors
      }
    }
  }
  throw new Error('Max retries exceeded');
}

async function example9_RetryLogic() {
  const outlook = new Outlook();

  const result = await retryOperation(
    () => outlook.compose()
      .subject('Test')
      .to(['test@example.com'])
      .body('Testing retry logic')
      .send(),
    3 // Max 3 retries
  );

  console.log('Email sent successfully');
}

// ============================================
// Example 10: Service Factory Pattern
// ============================================
class AzureServiceFactory {
  private auth: AzureAuth;

  constructor(config?: any) {
    this.auth = new AzureAuth(config);
  }

  getOutlook(): Outlook {
    return new Outlook(this.auth);
  }

  getCalendar(): Calendar {
    return new Calendar(this.auth);
  }

  getTeams(): Teams {
    return new Teams(this.auth);
  }

  getSharePoint(siteId?: string): SharePoint {
    return new SharePoint(this.auth, siteId);
  }
}

async function example10_FactoryPattern() {
  const factory = new AzureServiceFactory({
    refreshToken: '...',
    clientId: '...',
    clientSecret: '...',
    tenantId: '...'
  });

  const outlook = factory.getOutlook();
  const calendar = factory.getCalendar();
  const teams = factory.getTeams();

  await outlook.compose().subject('Test').to(['test@example.com']).send();
  await calendar.getCalendars();

  const myTeams = await teams.getTeams();
  console.log('Teams:', myTeams.map(t => t.displayName));
}

// ============================================
// Example 11: Credential Management
// ============================================
async function example11_CredentialManagement() {
  // List all stored credentials
  const stored = await AzureAuth.listStoredCredentials();
  console.log('Stored credentials:', stored);

  // Clear specific tenant/client credentials
  if (stored.length > 0 && stored[0].tenantId && stored[0].clientId) {
    await AzureAuth.clearStoredCredentials(stored[0].tenantId, stored[0].clientId);
    console.log('Cleared specific credentials');
  }

  // Clear all stored credentials
  // await AzureAuth.clearStoredCredentials();
}

// Run examples
if (require.main === module) {
  (async () => {
    try {
      // Uncomment the example you want to run
      // await example1_SharedAuth();
      // await example2_TokenProvider();
      // await example3_CustomScopes();
      // await example4_TaskQueueProcessing();
      // await example5_DeploymentNotification();
      // await example6_MultiCalendarHolidays();
      // await example7_MultiSiteManagement();
      // await example8_EmailAutomation();
      // await example9_RetryLogic();
      // await example10_FactoryPattern();
      // await example11_CredentialManagement();
    } catch (error) {
      console.error('Error:', error);
    }
  })();
}
