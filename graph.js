// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <graphInitSnippet>
let graphClient = undefined;

function initializeGraphClient(msalClient, account, scopes)
{
  // Create an authentication provider
  const authProvider = new MSGraphAuthCodeMSALBrowserAuthProvider
  .AuthCodeMSALBrowserAuthenticationProvider(msalClient, {
    account: account,
    scopes: scopes,
    interactionType: msal.InteractionType.PopUp
  });

  // Initialize the Graph client
  graphClient = MicrosoftGraph.Client.initWithMiddleware({authProvider});
}
// </graphInitSnippet>

// <getUserSnippet>
async function getUser() {
    return graphClient
      .api('/me')
      // Only get the fields used by the app
      .select('id,displayName,mail,userPrincipalName')
      .get();
  }
  // </getUserSnippet>

  // <getEventsSnippet>
// </getEventsSnippet>

// <createEventSnippet>
async function createNewMonitor() {
  const user = JSON.parse(sessionStorage.getItem('graphUser'));

  // Get the user's input
  const displayName = document.getElementById('mon-displayName').value;
  const description = document.getElementById('mon-description').value;
  const baseline = document.getElementById('mon-baseline').value;
  const parameters = document.getElementById('mon-parameters').value;

  // Require at least subject, start, and end
  if (!displayName || !baseline) {
    updatePage(Views.error, {
      message: 'Please provide a display name and content for the baseline.'
    });
    return;
  }

  let newMonitor = {
    displayName: displayName,
    description: description,
    baseline: JSON.parse(baseline)
  };

  if ('' != parameters)
  {
    newMonitor.parameters = JSON.parse(parameters)
  }

  try {
    await graphClient
      .api('https://graph.microsoft.com/beta/admin/configurationManagement/configurationMonitors/')
      .header('Content-Type', 'application/json')
      .post(newMonitor);

    // Return to the calendar view
    getMonitors();
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error creating monitor',
      debug: error + ": " + newMonitor.baseline
    });
  }
}

async function createNewSnapshot() {
  const user = JSON.parse(sessionStorage.getItem('graphUser'));

  // Get the user's input
  const displayName = document.getElementById('mon-displayName').value;
  const description = document.getElementById('mon-description').value;

  var result = [];
  var select = document.getElementById('mon-resources');
  var options = select && select.options;
  var opt;
  var resources = ""
  for (var i=0, iLen=options.length; i<iLen; i++) {
    opt = options[i];

    if (opt.selected)
    {
      resources += opt.value + ",";
    }
  }
  resources = resources.slice(0, -1);

  // Require at least subject, start, and end
  if (!displayName || !resources) {
    updatePage(Views.error, {
      message: 'Please provide a display name and the list of resources for the snapshot.'
    });
    return;
  }

  let newJob = {
    displayName: displayName,
    description: description,
    resources: resources.split(',')
  };

  try {
    await graphClient
      .api('https://graph.microsoft.com/beta/admin/configurationManagement/configurationSnapshots/createSnapshot')
      .header('Content-Type', 'application/json')
      .post(newJob);

    // Return to the calendar view
    getSnapshotJobs();
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error creating snapshot job',
      debug: error + ": " + newJob.resources
    });
  }
}
// </createEventSnippet>

async function getMonitors() {
  try {
    // GET /me/calendarview?startDateTime=''&endDateTime=''
    // &$select=subject,organizer,start,end
    // &$orderby=start/dateTime
    // &$top=50
    let responseMonitors = await graphClient
      .api('/admin/configurationManagement/configurationMonitors')
      // Set the Prefer=outlook.timezone header so date/times are in
      // user's preferred time zone
      .version('beta')
      //.header("Prefer", `outlook.timezone="${user.mailboxSettings.timeZone}"`)
      // Add the startDateTime and endDateTime query parameters
      //.query({ startDateTime: startOfWeek.format(), endDateTime: endOfWeek.format() })
      // Select just the fields we are interested in
      .select('displayName,id,status,createdBy')
      // Sort the results by start, earliest first
      .orderby('displayName')
      // Maximum 50 events in response
      .get();

    let responseMonitorRuns = await graphClient
      .api('/admin/configurationManagement/configurationMonitoringResults')
      // Set the Prefer=outlook.timezone header so date/times are in
      // user's preferred time zone
      .version('beta')
      //.header("Prefer", `outlook.timezone="${user.mailboxSettings.timeZone}"`)
      // Add the startDateTime and endDateTime query parameters
      //.query({ startDateTime: startOfWeek.format(), endDateTime: endOfWeek.format() })
      // Select just the fields we are interested in
      .select('id,runStatus,driftsCount,monitorId, runInitiationDateTime, runCompletionDateTime')
      // Sort the results by start, earliest first
      .orderby('runInitiationDateTime desc')
      .top(100)
      .get();


    updatePage(Views.monitors, responseMonitors.value, responseMonitorRuns.value);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}

async function getDrifts(monitorId) {

  try {
    // GET /me/calendarview?startDateTime=''&endDateTime=''
    // &$select=subject,organizer,start,end
    // &$orderby=start/dateTime
    // &$top=50
    let responseDrifts = await graphClient
      .api("https://graph.microsoft.com/beta/admin/configurationManagement/configurationDrifts?filter=MonitorId eq '" + monitorId + "'")
      // Set the Prefer=outlook.timezone header so date/times are in
      // user's preferred time zone
      .version('beta')
      //.header("Prefer", `outlook.timezone="${user.mailboxSettings.timeZone}"`)
      // Add the startDateTime and endDateTime query parameters
      //.query({ startDateTime: startOfWeek.format(), endDateTime: endOfWeek.format() })
      // Select just the fields we are interested in
      .select('id,resourceType,firstReportedDateTime,status,resourceInstanceIdentifier,driftedProperties')
      // Sort the results by start, earliest first
      //.orderby('start/dateTime')
      // Maximum 50 events in response
      .get();


    updatePage(Views.drifts, responseDrifts.value);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}

async function getAllDrifts() {

  // GET /me/calendarview?startDateTime=''&endDateTime=''
  // &$select=subject,organizer,start,end
  // &$orderby=start/dateTime
  // &$top=50
  return graphClient
    .api("https://graph.microsoft.com/beta/admin/configurationManagement/configurationDrifts")
    // Set the Prefer=outlook.timezone header so date/times are in
    // user's preferred time zone
    .version('beta')
    //.header("Prefer", `outlook.timezone="${user.mailboxSettings.timeZone}"`)
    // Add the startDateTime and endDateTime query parameters
    //.query({ startDateTime: startOfWeek.format(), endDateTime: endOfWeek.format() })
    // Select just the fields we are interested in
    .select('status,tenantId')
    // Sort the results by start, earliest first
    //.orderby('start/dateTime')
    // Maximum 50 events in response
    .top(1)
    .get();
}

async function getSnapshotJobs() {

  try {
    let responseJobs = await graphClient
      .api("https://graph.microsoft.com/beta/admin/configurationManagement/configurationSnapshotJobs")
      .version('beta')
      .responseType('raw')
      .select('id,displayName,description,status,createdDateTime,completedDateTime,resourceLocation,resources,createdBy,errorDetails')
      .orderby('createdDateTime desc')
      .get();


    updatePage(Views.snapshots, responseJobs.value);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting Snapshot Jobs',
      debug: error
    });
  }
}

async function getSnapshot(id) {

  try
  {
    let snapshot = await graphClient
      .api("https://graph.microsoft.com/beta/admin/configurationManagement/configurationSnapshots('" + id + "')")
      .version('beta')
      .get();

    updatePage(Views.snapshotInfo, snapshot);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting Snapshot Jobs',
      debug: error
    });
  }
}

async function getSnapshotErrors(jobId) {

  try
  {
    let errors = await graphClient
      .api("https://graph.microsoft.com/beta/admin/configurationManagement/configurationSnapshotJobs('" + jobId + "')")
      .select("errorDetails")
      .version('beta')
      .get();

    updatePage(Views.snapshotErrors, errors);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting Snapshot Errors',
      debug: error
    });
  }
}

async function deleteMonitor(monitorId) {

  try {
    let responseDrifts = await graphClient
      .api("https://graph.microsoft.com/beta/admin/configurationManagement/configurationMonitors('" + monitorId + "')")
      .delete();

      getMonitors();
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}

async function deleteSnapshotJob(jobId) {

  try {
    let responseDrifts = await graphClient
      .api("https://graph.microsoft.com/beta/admin/configurationManagement/configurationSnapshotJobs('" + jobId + "')")
      .delete();

      getSnapshotJobs();
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}
