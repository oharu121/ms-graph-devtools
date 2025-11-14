/**
 * Basic Usage Examples
 * This file demonstrates common usage patterns for the Azure utility
 */

import Azure, { Outlook, Calendar, Teams, SharePoint } from '../src/index';

// ============================================
// Example 1: Using Environment Variables
// ============================================
// Set these environment variables:
// AZURE_CLIENT_ID=your-client-id
// AZURE_CLIENT_SECRET=your-client-secret
// AZURE_TENANT_ID=your-tenant-id
// AZURE_REFRESH_TOKEN=your-refresh-token

async function example1_AutoLoadFromEnv() {
  // No config needed - auto-loads from env
  const outlook = new Outlook();

  const user = await outlook.getMe();
  console.log(`Hello, ${user.displayName}!`);

  await outlook.compose()
    .subject('Test Email')
    .body('This is a test email from Azure utility', 'Text')
    .to(['test@example.com'])
    .send();
}

// ============================================
// Example 2: Explicit Configuration
// ============================================
async function example2_ExplicitConfig() {
  const outlook = new Outlook({
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret',
    tenantId: 'your-tenant-id',
    refreshToken: 'your-refresh-token'
  });

  const emails = await outlook.getMails('2024-01-15', 'invoice');
  console.log(`Found ${emails.length} emails`);
}

// ============================================
// Example 3: Global Configuration
// ============================================
async function example3_GlobalConfig() {
  // Set config once for all services
  Azure.config({
    clientId: process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    tenantId: process.env.AZURE_TENANT_ID!,
    refreshToken: process.env.AZURE_REFRESH_TOKEN!
  });

  // All services use global config
  const outlook = new Outlook();
  const calendar = new Calendar();
  const teams = new Teams();

  await outlook.compose()
    .subject('Global Config Test')
    .body('Testing global configuration')
    .to(['test@example.com'])
    .send();

  const calendars = await calendar.getCalendars();
  console.log('Available calendars:', calendars.map(c => c.name));
}

// ============================================
// Example 4: Calendar Operations
// ============================================
async function example4_CalendarOperations() {
  const calendar = new Calendar();

  // Get all calendars
  const calendars = await calendar.getCalendars();
  console.log('Calendars:', calendars);

  // Get India holidays for 2024 (default calendar name)
  const holidays = await calendar.getIndiaHolidays(
    '2024-01-01T00:00:00Z',
    '2024-12-31T23:59:59Z'
  );

  holidays.forEach(holiday => {
    console.log(`${holiday.name}: ${holiday.date}`);
  });

  // Get holidays with custom calendar name
  const customHolidays = await calendar.getIndiaHolidays(
    '2024-01-01T00:00:00Z',
    '2024-12-31T23:59:59Z',
    'Indian Holidays' // Custom calendar name
  );

  // Get any country's holidays using generic method
  const usHolidays = await calendar.getHolidaysByCalendarName(
    'US Holidays',
    '2024-01-01T00:00:00Z',
    '2024-12-31T23:59:59Z'
  );
}

// ============================================
// Example 5: SharePoint Operations
// ============================================
async function example5_SharePointOperations() {
  const sharepoint = new SharePoint();

  // Search for your site
  const sites = await sharepoint.searchSites('Engineering');
  if (sites.length > 0) {
    sharepoint.setSiteId(sites[0].id);
  }

  // Get all lists in site
  const lists = await sharepoint.getLists();
  console.log('Available lists:', lists.map(l => l.displayName));

  // Create a task
  await sharepoint.createListItem('Tasks', {
    Title: 'New Task',
    Status: 'Active',
    Priority: 'High',
    DueDate: '2024-12-31'
  });

  // Query active tasks
  const activeTasks = await sharepoint.getListItems('Tasks', {
    filter: "fields/Status eq 'Active'",
    orderby: 'createdDateTime desc',
    top: 10,
    expand: 'fields'
  });

  console.log(`Found ${activeTasks.length} active tasks`);

  // Update a task
  if (activeTasks.length > 0) {
    await sharepoint.updateListItem('Tasks', activeTasks[0].id, {
      Status: 'In Progress'
    });
  }

  // Get latest item
  const latestTask = await sharepoint.getLatestItem('Tasks');
  console.log('Latest task:', latestTask);
}

// ============================================
// Example 6: Teams Adaptive Card with Builder
// ============================================
async function example6_TeamsAdaptiveCard() {
  const teams = new Teams();

  // Get your teams and channels
  const myTeams = await teams.getTeams();
  console.log('Available teams:', myTeams.map(t => t.displayName));

  if (myTeams.length > 0) {
    const channels = await teams.getChannels(myTeams[0].id);
    console.log('Channels:', channels.map(c => c.displayName));

    // Send adaptive card using builder
    await teams.compose()
      .team(myTeams[0].id)
      .channel(channels[0].id)
      .card({
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
          {
            type: 'TextBlock',
            text: 'Deployment Complete!',
            size: 'Large',
            weight: 'Bolder',
            color: 'Good'
          },
          {
            type: 'TextBlock',
            text: 'Version 2.0 has been deployed successfully.'
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Environment:', value: 'Production' },
              { title: 'Build:', value: '#1234' },
              { title: 'Time:', value: new Date().toLocaleString() }
            ]
          }
        ],
        actions: [
          {
            type: 'Action.OpenUrl',
            title: 'View Dashboard',
            url: 'https://dashboard.example.com'
          }
        ]
      })
      .mentionTeam(myTeams[0].id, myTeams[0].displayName)
      .send();
  }
}

// ============================================
// Example 7: Teams with Mentions
// ============================================
async function example7_TeamsWithMentions() {
  const teams = new Teams();

  const myTeams = await teams.getTeams();
  const teamId = myTeams[0].id;

  // Get tags for mentions
  const tags = await teams.getTags(teamId);
  const channels = await teams.getChannels(teamId);

  await teams.compose()
    .team(teamId)
    .channel(channels[0].id)
    .card({
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: 'Urgent: Action Required',
          size: 'Large',
          color: 'Attention'
        }
      ]
    })
    .mentionTeam(teamId, 'Engineering Team')
    .mentionTag(tags[0]?.id, tags[0]?.displayName)
    .send();
}

// ============================================
// Example 8: Fluent Email Builder
// ============================================
async function example8_FluentEmailBuilder() {
  const outlook = new Outlook();

  // Simple text email
  await outlook.compose()
    .subject('Meeting Reminder')
    .body('Don\'t forget our meeting tomorrow at 10 AM!', 'Text')
    .to(['colleague@example.com'])
    .cc(['manager@example.com'])
    .importance('high')
    .send();

  // HTML email with attachments
  await outlook.compose()
    .subject('Monthly Report')
    .body('<h1>Monthly Report</h1><p>Please review the attached documents.</p>', 'HTML')
    .to(['boss@example.com'])
    .attachments(['./report.pdf', './charts.xlsx'])
    .requestReadReceipt(true)
    .send();
}

// Run examples
if (require.main === module) {
  (async () => {
    try {
      // Uncomment the example you want to run
      // await example1_AutoLoadFromEnv();
      // await example2_ExplicitConfig();
      // await example3_GlobalConfig();
      // await example4_CalendarOperations();
      // await example5_SharePointOperations();
      // await example6_TeamsAdaptiveCard();
      // await example7_TeamsWithMentions();
      // await example8_FluentEmailBuilder();
    } catch (error) {
      console.error('Error:', error);
    }
  })();
}
