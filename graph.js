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

async function getUser() {
    return graphClient
      .api('/me')
      // Only get the fields used by the app
      .select('id,displayName,mail,userPrincipalName')
      .get();
  }

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


async function getMonitors() {
  try {
    var uri = 'https://graph.microsoft.com/beta/admin/configurationManagement/configurationMonitors';
    let responseMonitors = await graphClient
      .api(uri)
      .version('beta')
      .select('displayName,id,status,createdBy')
      .orderby('displayName')
      .top(10)
      .get();

    let responseMonitorRuns = await graphClient
      .api('/admin/configurationManagement/configurationMonitoringResults')
      .version('beta')
      .select('id,runStatus,driftsCount,monitorId, runInitiationDateTime, runCompletionDateTime')
      .orderby('runInitiationDateTime desc')
      .top(100)
      .get();


    updatePage(Views.monitors, responseMonitors.value, responseMonitorRuns.value, uri);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}

async function getMonitorDetails(monitorId) {
  try {
    var uri = "https://graph.microsoft.com/beta/admin/configurationManagement/configurationMonitors('" + monitorId + "')";
    let responseMonitor = await graphClient
      .api(uri)
      .version('beta')
      .top(1)
      .get();

    uri = "https://graph.microsoft.com/beta/admin/configurationManagement/configurationMonitors('" + monitorId + "')/baseline";
    let responseMonitorBaseline = await graphClient
      .api(uri)
      .version('beta')
      .get();

    updatePage(Views.editMonitor, responseMonitor.value, responseMonitorBaseline.value);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}

async function getDrifts(monitorId) {

  try {
    var uri = "https://graph.microsoft.com/beta/admin/configurationManagement/configurationDrifts?filter=MonitorId eq '" + monitorId + "'"
    let responseDrifts = await graphClient
      .api(uri)
      .version('beta')
      .select('id,resourceType,firstReportedDateTime,status,resourceInstanceIdentifier,driftedProperties')
      .get();


    updatePage(Views.drifts, responseDrifts.value, null, uri);
  } catch (error) {
    updatePage(Views.error, {
      message: 'Error getting events',
      debug: error
    });
  }
}

async function getAllDrifts() {
  return graphClient
    .api("https://graph.microsoft.com/beta/admin/configurationManagement/configurationDrifts")
    .version('beta')
    .select('status,tenantId')
    .top(1)
    .get();
}

async function getSnapshotJobs() {

  try {
    var uri = "https://graph.microsoft.com/beta/admin/configurationManagement/configurationSnapshotJobs";
    let responseJobs = await graphClient
      .api(uri)
      .version('beta')
      .select('id,displayName,description,status,createdDateTime,completedDateTime,resourceLocation,resources,createdBy,errorDetails')
      .orderby('createdDateTime desc')
      .get();

    updatePage(Views.snapshots, responseJobs.value, null, uri);
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
    var uri = "https://graph.microsoft.com/beta/admin/configurationManagement/configurationSnapshots('" + id + "')";
    let snapshot = await graphClient
      .api(uri)
      .version('beta')
      .get();

    updatePage(Views.snapshotInfo, snapshot, null, uri);
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
    var uri = "https://graph.microsoft.com/beta/admin/configurationManagement/configurationSnapshotJobs('" + jobId + "')"
    let errors = await graphClient
      .api(uri)
      .select("errorDetails")
      .version('beta')
      .get();

    updatePage(Views.snapshotErrors, errors, null, uri);
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
