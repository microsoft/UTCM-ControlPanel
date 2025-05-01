//Copyright (c) Microsoft Corporation.
//Licensed under the MIT License.

// Select DOM elements to work with
const authenticatedNav = document.getElementById('authenticated-nav');
const accountNav = document.getElementById('account-nav');
const mainContainer = document.getElementById('main-container');

const Views = { error: 1, home: 2, monitors: 3, drifts:4, snapshots:5, snapshotInfo:6, snapshotErrors:7 };

function createElement(type, className, text) {
  var element = document.createElement(type);
  element.className = className;

  if (text) {
    var textNode = document.createTextNode(text);
    element.appendChild(textNode);
  }

  return element;
}

function showAuthenticatedNav(user, view) {
  authenticatedNav.innerHTML = '';

  if (user) {
    var monitorNav = createElement('li', 'nav-item');

    var monitorLink = createElement('button',
      `btn btn-link nav-link${view === Views.monitors ? ' active' : '' }`,
      'Monitors');
    monitorLink.setAttribute('onclick', 'getMonitors();');
    monitorNav.appendChild(monitorLink);

    authenticatedNav.appendChild(monitorNav);

    var snapshotNav = createElement('li', 'nav-item');

    var snapshotLink = createElement('button',
      `btn btn-link nav-link${view === Views.monitors ? ' active' : '' }`,
      'Snapshots');
    snapshotLink.setAttribute('onclick', 'getSnapshotJobs();');
    snapshotNav.appendChild(snapshotLink);

    authenticatedNav.appendChild(snapshotNav);
  }
}

function showAccountNav(user) {
  accountNav.innerHTML = '';

  if (user) {
    // Show the "signed-in" nav
    accountNav.className = 'nav-item dropdown';

    var dropdown = createElement('a', 'nav-link dropdown-toggle');
    dropdown.setAttribute('data-bs-toggle', 'dropdown');
    dropdown.setAttribute('role', 'button');
    accountNav.appendChild(dropdown);

    let userIcon = createElement('img', 'rounded-circle align-self-center me-2');
    userIcon.style.width = '32px';
    userIcon.src = 'utcm.png';
    userIcon.alt = 'user';
    dropdown.appendChild(userIcon);

    var menu = createElement('div', 'dropdown-menu dropdown-menu-end');
    accountNav.appendChild(menu);

    var userName = createElement('h5', 'dropdown-item-text mb-0', user.displayName);
    menu.appendChild(userName);

    var userEmail = createElement('p', 'dropdown-item-text text-muted mb-0', user.mail || user.userPrincipalName);
    menu.appendChild(userEmail);

    var divider = createElement('hr', 'dropdown-divider');
    menu.appendChild(divider);

    var signOutButton = createElement('button', 'dropdown-item', 'Sign out');
    signOutButton.setAttribute('onclick', 'signOut();');
    menu.appendChild(signOutButton);
  } else {
    // Show a "sign in" button
    accountNav.className = 'nav-item';

    var signInButton = createElement('button', 'btn btn-link nav-link', 'Sign in');
    signInButton.setAttribute('onclick', 'signIn();');
    accountNav.appendChild(signInButton);
  }
}

function showWelcomeMessage(user, drifts) {
  // Create jumbotron
  let jumbotron = createElement('div', 'p-5 mb-4 bg-light rounded-3');

  let container = createElement('div', 'container-fluid py-5');
  jumbotron.appendChild(container);

  let heading = createElement('h1', null, 'UTCM - Control Panel');
  container.appendChild(heading);

  let lead = createElement('p', 'lead',
    'View monitors, drifts and snapshots in a central location.');
    container.appendChild(lead);

  if (!user)
  {
    // Show a sign in button in the jumbotron
    let signInButton = createElement('button', 'btn btn-primary btn-large',
      'Click here to sign in');
    signInButton.setAttribute('onclick', 'signIn();')
    container.appendChild(signInButton);
  }
  else
  {
    let spanDrifts = createElement('span')
    if(null != drifts && drifts.value.length > 0)
    {
      spanDrifts.innerHTML = "<span id='" + drifts.value[0].tenantId + "'><img src='red.jpg' width='20' alt='" + drifts.value.length + " drifts detected' />&nbsp;Tenant " + drifts.value[0].tenantId + " has <strong>" + drifts.value.length + " Active Drifts</strong></span>"
    }
    else
    {
      spanDrifts.innerHTML = "<span id='NoDrift'><img src='green.png' width='20' alt='No drift detected' />&nbsp;No active drifts detected for the current tenant.</span>"
    }
    container.appendChild(spanDrifts);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(jumbotron);
}

function showError(error) {
  var alert = createElement('div', 'alert alert-danger');

  var message = createElement('p', 'mb-3', error.message);
  alert.appendChild(message);

  if (error.debug)
  {
    var pre = createElement('pre', 'alert-pre border bg-light p-2');
    alert.appendChild(pre);

    var code = createElement('code', 'text-break text-wrap',
      JSON.stringify(error.debug, null, 2));
    pre.appendChild(code);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(alert);
}

function updatePage(view, data) {
  if (!view) {
    view = Views.home;
  }

  const user = JSON.parse(sessionStorage.getItem('graphUser'));

  showAccountNav(user);
  showAuthenticatedNav(user, view);

  switch (view) {
    case Views.error:
      showError(data);
      break;
    case Views.home:
      showWelcomeMessage(user, drifts);
      break;
    case Views.monitors:
      break;
    case Views.drifts:
      break;
    case Views.snapshots:
      break;
    case Views.snapshotInfo:
      break;
    case Views.snapshotErrors:
      break;
  }
}

// <updatePageSnippet>
function updatePage(view, data) {
  if (!view) {
    view = Views.home;
  }

  const user = JSON.parse(sessionStorage.getItem('graphUser'));

  showAccountNav(user);
  showAuthenticatedNav(user, view);

  switch (view) {
    case Views.error:
      showError(data);
      break;
    case Views.home:
      showWelcomeMessage(user, data);
      break;
    case Views.monitors:
      showMonitors(data);
      break;
    case Views.drifts:
      showDrifts(data);
      break;
    case Views.snapshots:
      showSnapshotJobs(data);
      break;
    case Views.snapshotInfo:
      showSnapshot(data);
      break;
    case Views.snapshotErrors:
      showSnapshotErrors(data);
      break;
  }
}

function updatePage(view, data, data2) {
  if (!view) {
    view = Views.home;
  }
  const user = JSON.parse(sessionStorage.getItem('graphUser'));
  const drifts = JSON.parse(sessionStorage.getItem('drifts'));

  showAccountNav(user);
  showAuthenticatedNav(user, view);

  switch (view) {
    case Views.error:
      showError(data);
      break;
    case Views.home:
      showWelcomeMessage(user, drifts);
      break;
    case Views.monitors:
      showMonitors(data, data2);
      break;
    case Views.drifts:
      showDrifts(data);
      break;
    case Views.snapshots:
      showSnapshotJobs(data);
      break;
    case Views.snapshotInfo:
      showSnapshot(data);
      break;
    case Views.snapshotErrors:
      showSnapshotErrors(data);
      break;
  }
}
// </updatePageSnippet>

function showNewSnapshotForm() {
  let form = document.createElement('form');

  let displayNameGroup = createElement('div', 'form-group mb-2');
  form.appendChild(displayNameGroup);

  displayNameGroup.appendChild(createElement('label', '', 'Display Name'));

  let displayNameInput = createElement('input', 'form-control');
  displayNameInput.setAttribute('id', 'mon-displayName');
  displayNameInput.setAttribute('type', 'text');
  displayNameGroup.appendChild(displayNameInput);

  let descriptionGroup = createElement('div', 'form-group mb-2');
  form.appendChild(descriptionGroup);

  descriptionGroup.appendChild(createElement('label', '', 'Description'));

  let descriptionInput = createElement('input', 'form-control');
  descriptionInput.setAttribute('id', 'mon-description');
  descriptionInput.setAttribute('type', 'text');
  descriptionGroup.appendChild(descriptionInput);

  let resourcesGroup = createElement('div', 'form-group mb-2');
  form.appendChild(resourcesGroup);

  var allResources = ["microsoft.entra.AdministrativeUnit","microsoft.entra.AuthenticationContextClassReference","microsoft.entra.AuthenticationMethodPolicy","microsoft.entra.AuthenticationMethodPolicyAuthenticator","microsoft.entra.AuthenticationMethodPolicyEmail","microsoft.entra.AuthenticationMethodPolicyFido2","microsoft.entra.AuthenticationMethodPolicySms","microsoft.entra.AuthenticationMethodPolicySoftware","microsoft.entra.AuthenticationMethodPolicyTemporary","microsoft.entra.AuthenticationMethodPolicyVoice","microsoft.entra.AuthenticationMethodPolicyX509","microsoft.entra.AuthenticationStrengthPolicy","microsoft.entra.AuthorizationPolicy","microsoft.entra.ConditionalAccessPolicy","microsoft.entra.CrossTenantAccessPolicy","microsoft.entra.CrossTenantAccessPolicyConfigurationDefault","microsoft.entra.CrossTenantAccessPolicyConfigurationPartner","microsoft.entra.EntitlementManagementAccessPackage","microsoft.entra.EntitlementManagementAccessPackageAssignmentPolicy","microsoft.entra.EntitlementManagementAccessPackageCatalog","microsoft.entra.EntitlementManagementConnectedOrganization","microsoft.entra.ExternalIdentityPolicy","microsoft.entra.Group","microsoft.entra.GroupLifecyclePolicy","microsoft.entra.NamedLocationPolicy","microsoft.entra.RoleDefinition","microsoft.entra.RoleEligibilityScheduleRequest","microsoft.entra.RoleSetting","microsoft.entra.SecurityDefaults","microsoft.entra.ServicePrincipal","microsoft.entra.SocialIdentityProvider","microsoft.entra.TenantDetails","microsoft.entra.TokenLifetimePolicy","microsoft.entra.User","microsoft.exchange.AcceptedDomain","microsoft.exchange.ActiveSyncDeviceAccessRule","microsoft.exchange.AddressBookPolicy","microsoft.exchange.AntiPhishPolicy","microsoft.exchange.AntiPhishRule","microsoft.exchange.AtpPolicyForO365","microsoft.exchange.AuthenticationPolicy","microsoft.exchange.AuthenticationPolicyAssignment","microsoft.exchange.AvailabilityAddressSpace","microsoft.exchange.AvailabilityConfig","microsoft.exchange.CalendarProcessing","microsoft.exchange.CASMailboxPlan","microsoft.exchange.CASMailboxSettings","microsoft.exchange.DataClassification","microsoft.exchange.DataEncryptionPolicy","microsoft.exchange.DistributionGroup","microsoft.exchange.DkimSigningConfig","microsoft.exchange.EmailAddressPolicy","microsoft.exchange.GroupSettings","microsoft.exchange.HostedConnectionFilterPolicy","microsoft.exchange.HostedContentFilterPolicy","microsoft.exchange.HostedContentFilterRule","microsoft.exchange.HostedOutboundSpamFilterRule","microsoft.exchange.InboundConnector","microsoft.exchange.IntraOrganizationConnector","microsoft.exchange.IRMConfiguration","microsoft.exchange.JournalRule","microsoft.exchange.MailboxAutoReplyConfiguration","microsoft.exchange.MailboxCalendarFolder","microsoft.exchange.MailboxPermission","microsoft.exchange.MailboxPlan","microsoft.exchange.MailboxSettings","microsoft.exchange.MailContact","microsoft.exchange.MalwareFilterPolicy","microsoft.exchange.MalwareFilterRule","microsoft.exchange.ManagementRole","microsoft.exchange.ManagementRoleAssignment","microsoft.exchange.MessageClassification","microsoft.exchange.MobileDeviceMailboxPolicy","microsoft.exchange.OMEConfiguration","microsoft.exchange.OnPremisesOrganization","microsoft.exchange.OrganizationConfig","microsoft.exchange.OrganizationRelationship","microsoft.exchange.HostedOutboundSpamFilterPolicy","microsoft.exchange.OutboundConnector","microsoft.exchange.OwaMailboxPolicy","microsoft.exchange.PartnerApplication","microsoft.exchange.PerimeterConfiguration","microsoft.exchange.Place","microsoft.exchange.PolicyTipConfig","microsoft.exchange.QuarantinePolicy","microsoft.exchange.RecipientPermission","microsoft.exchange.RemoteDomain","microsoft.exchange.ReportSubmissionPolicy","microsoft.exchange.ReportSubmissionRule","microsoft.exchange.ResourceConfiguration","microsoft.exchange.RoleAssignmentPolicy","microsoft.exchange.RoleGroup","microsoft.exchange.SafeAttachmentPolicy","microsoft.exchange.SafeAttachmentRule","microsoft.exchange.SafeLinksPolicy","microsoft.exchange.SafeLinksRule","microsoft.exchange.SharedMailbox","microsoft.exchange.SharingPolicy","microsoft.exchange.TransportConfig","microsoft.exchange.TransportRule","microsoft.intune.AccountProtectionLocalUserGroupMembershipPolicy","microsoft.intune.AccountProtectionPolicy","microsoft.intune.ApplicationControlPolicyWindows10","microsoft.intune.AppProtectionPolicyiOS","microsoft.intune.DeviceAndAppManagementAssignmentFilter","microsoft.intune.DeviceCategory","microsoft.intune.DeviceCleanupRule","microsoft.intune.DeviceCompliancePolicyAndroid","microsoft.intune.DeviceCompliancePolicyAndroidDeviceOwner","microsoft.intune.DeviceCompliancePolicyAndroidWorkProfile","microsoft.intune.DeviceCompliancePolicyiOs","microsoft.intune.DeviceCompliancePolicyMacOS","microsoft.intune.DeviceCompliancePolicyWindows10","microsoft.intune.DeviceConfigurationDefenderForEndpointOnboardingPolicyWindows10","microsoft.intune.DeviceConfigurationDomainJoinPolicyWindows10","microsoft.intune.DeviceConfigurationEmailProfilePolicyWindows10","microsoft.intune.DeviceConfigurationEndpointProtectionPolicyWindows10","microsoft.intune.DeviceConfigurationFirmwareInterfacePolicyWindows10","microsoft.intune.DeviceConfigurationHealthMonitoringConfigurationPolicyWindows10","microsoft.intune.DeviceConfigurationIdentityProtectionPolicyWindows10","microsoft.intune.DeviceConfigurationImportedPfxCertificatePolicyWindows10","microsoft.intune.DeviceConfigurationNetworkBoundaryPolicyWindows10","microsoft.intune.DeviceConfigurationPkcsCertificatePolicyWindows10","microsoft.intune.DeviceConfigurationPolicyAndroidDeviceOwner","microsoft.intune.DeviceConfigurationPolicyAndroidOpenSourceProject","microsoft.intune.DeviceConfigurationPolicyAndroidWorkProfile","microsoft.intune.DeviceConfigurationPolicyWindows10","microsoft.intune.DeviceConfigurationSharedMultiDevicePolicyWindows10","microsoft.intune.DeviceConfigurationTrustedCertificatePolicyWindows10","microsoft.intune.DeviceConfigurationVpnPolicyWindows10","microsoft.intune.DeviceConfigurationWindowsTeamPolicyWindows10","microsoft.intune.DeviceEnrollmentLimitRestriction","microsoft.intune.RoleDefinition","microsoft.intune.WifiConfigurationPolicyAndroidOpenSourceProject","microsoft.intune.WifiConfigurationPolicyIOS","microsoft.intune.WifiConfigurationPolicyWindows10","microsoft.intune.WindowsInformationProtectionPolicyWindows10MdmEnrolled","microsoft.intune.WindowsUpdateForBusinessRingUpdateProfileWindows10","microsoft.teams.AppPermissionPolicy","microsoft.teams.AppSetupPolicy","microsoft.teams.CallHoldPolicy","microsoft.teams.CallingPolicy","microsoft.teams.CallParkPolicy","microsoft.teams.ChannelsPolicy","microsoft.teams.ClientConfiguration","microsoft.teams.CortanaPolicy","microsoft.teams.DialInConferencingTenantSettings","microsoft.teams.EnhancedEncryptionPolicy","microsoft.teams.EventsPolicy","microsoft.teams.FederationConfiguration","microsoft.teams.FeedbackPolicy","microsoft.teams.FilesPolicy","microsoft.teams.GuestCallingConfiguration","microsoft.teams.GuestMeetingConfiguration","microsoft.teams.GuestMessagingConfiguration","microsoft.teams.IPPhonePolicy","microsoft.teams.MeetingBroadcastConfiguration","microsoft.teams.MeetingBroadcastPolicy","microsoft.teams.MeetingConfiguration","microsoft.teams.MeetingPolicy","microsoft.teams.MessagingPolicy","microsoft.teams.MobilityPolicy","microsoft.teams.NetworkRoamingPolicy","microsoft.teams.OnlineVoicemailUserSettings","microsoft.teams.OnlineVoicemailPolicy","microsoft.teams.PstnUsage","microsoft.teams.ShiftsPolicy","microsoft.teams.TemplatesPolicy","microsoft.teams.TenantDialPlan","microsoft.teams.TenantNetworkRegion","microsoft.teams.TenantNetworkSite","microsoft.teams.TenantNetworkSubnet","microsoft.teams.TenantTrustedIPAddress","microsoft.teams.TranslationRule","microsoft.teams.UnassignedNumberTreatment","microsoft.teams.UpdateManagementPolicy","microsoft.teams.UpgradeConfiguration","microsoft.teams.UpgradePolicy","microsoft.teams.VdiPolicy","microsoft.teams.VoiceRoute","microsoft.teams.VoiceRoutingPolicy","microsoft.teams.WorkloadPolicy"];
  let resourceDDL = createElement('select');
  resourceDDL.setAttribute('multiple', "true");
  resourceDDL.setAttribute('id', 'mon-resources');
  resourceDDL.setAttribute('name', 'mon-resources');
  resourceDDL.style.height = "200px";
  resourceDDL.style.width = "100%";
  for(const resource of allResources)
  {
    resChoice = createElement('option', null, resource);
    resChoice.setAttribute('value', resource);
    resourceDDL.appendChild(resChoice);
  }
  form.appendChild(resourceDDL);

  let createButton = createElement('button', 'btn btn-primary me-2', 'Create');
  createButton.setAttribute('type', 'button');
  createButton.setAttribute('onclick', 'createNewSnapshot();');
  form.appendChild(createButton);

  let cancelButton = createElement('button', 'btn btn-secondary', 'Cancel');
  cancelButton.setAttribute('type', 'button');
  cancelButton.setAttribute('onclick', 'getSnapshotJobs();');
  form.appendChild(cancelButton);

  mainContainer.innerHTML = '';
  mainContainer.appendChild(form);
}

function showNewMonitorForm() {
  let form = document.createElement('form');

  let displayNameGroup = createElement('div', 'form-group mb-2');
  form.appendChild(displayNameGroup);

  displayNameGroup.appendChild(createElement('label', '', 'Display Name'));

  let displayNameInput = createElement('input', 'form-control');
  displayNameInput.setAttribute('id', 'mon-displayName');
  displayNameInput.setAttribute('type', 'text');
  displayNameGroup.appendChild(displayNameInput);

  let descriptionGroup = createElement('div', 'form-group mb-2');
  form.appendChild(descriptionGroup);

  descriptionGroup.appendChild(createElement('label', '', 'Description'));

  let descriptionInput = createElement('input', 'form-control');
  descriptionInput.setAttribute('id', 'mon-description');
  descriptionInput.setAttribute('type', 'text');
  descriptionGroup.appendChild(descriptionInput);

  let baselineGroup = createElement('div', 'form-group mb-2');
  form.appendChild(baselineGroup);

  baselineGroup.appendChild(createElement('label', '', 'Baseline'));
  let baselineInput = createElement('textarea', 'form-control');
  baselineInput.setAttribute('id', 'mon-baseline');
  baselineInput.setAttribute('type', 'text');
  baselineInput.setAttribute('rows', '20');
  baselineGroup.appendChild(baselineInput);

  let parametersGroup = createElement('div', 'form-group mb-2');
  form.appendChild(parametersGroup);

  parametersGroup.appendChild(createElement('label', '', 'Parameters'));
  let parametersInput = createElement('textarea', 'form-control');
  parametersInput.setAttribute('id', 'mon-parameters');
  parametersInput.setAttribute('type', 'text');
  parametersInput.setAttribute('rows', '5');
  parametersGroup.appendChild(parametersInput);

  let createButton = createElement('button', 'btn btn-primary me-2', 'Create');
  createButton.setAttribute('type', 'button');
  createButton.setAttribute('onclick', 'createNewMonitor();');
  form.appendChild(createButton);

  let cancelButton = createElement('button', 'btn btn-secondary', 'Cancel');
  cancelButton.setAttribute('type', 'button');
  cancelButton.setAttribute('onclick', 'getMonitors();');
  form.appendChild(cancelButton);

  mainContainer.innerHTML = '';
  mainContainer.appendChild(form);
}

function showSnapshot(data) {

  let divCountResources = document.createElement('div')
  divCountResources.innerHTML = "<strong>This snapshot contains:</strong> " + data.resources.length + " resources";
  let form = document.createElement('form');

  let contentGroup = createElement('div', 'form-group mb-2');
  contentGroup.appendChild(divCountResources);
  form.appendChild(contentGroup);

  contentGroup.appendChild(createElement('label', '', 'Snapshot Content:'));

  let contentInput = createElement('textarea', 'form-control');
  contentInput.setAttribute('id', 'snap-content');
  contentInput.setAttribute('type', 'text');
  contentInput.setAttribute('rows', '30');
  contentInput.innerHTML = JSON.stringify(data, null, 4);
  contentGroup.appendChild(contentInput);

  mainContainer.innerHTML = '';
  mainContainer.appendChild(form);
}

function showSnapshotErrors(snapshotErrors) {

  let div = document.createElement('div');
  div.appendChild(createElement('h1', 'mb-3', 'Snapshot Errors'));

  let tableErrors = createElement('table', 'table');
  div.appendChild(tableErrors);

  let thead = document.createElement('thead');
  tableErrors.appendChild(thead);

  let headerrow = document.createElement('tr');
  thead.appendChild(headerrow);

  let nameCell = createElement('th', null, 'Details');
  nameCell.setAttribute('scope', 'col');
  headerrow.appendChild(nameCell);

  let tbody = document.createElement('tbody');
  tableErrors.appendChild(tbody);

  var i = 1;
  for (const error of snapshotErrors.errorDetails)
  {
    let errorRow = document.createElement('tr');
    errorRow.setAttribute('key', i);
    tbody.appendChild(errorRow);

    let cell1 = createElement('td', null, error.slice(0, -1));
    errorRow.appendChild(cell1);
    i++;
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(div);
}

function showSnapshotJobs(snapshotJobs) {

  let div = document.createElement('div');

  div.appendChild(createElement('h1', 'mb-3', 'Snapshot Jobs'));

  let newEventButton = createElement('button', 'btn btn-light btn-sm mb-3 btn-create', 'Create Snapshot Job');
  newEventButton.setAttribute('onclick', 'showNewSnapshotForm();');
  div.appendChild(newEventButton);

  let refreshIcon = createElement('span');
  refreshIcon.innerHTML = "&nbsp;&nbsp;<img src='refresh.jpg' alt='Refresh' onclick='getSnapshotJobs();' width='25' style='cursor:pointer;float:right;margin-top:-10px;' />";
  div.appendChild(refreshIcon);

  let tableJobs = createElement('table', 'table');
  div.appendChild(tableJobs);

  let thead = document.createElement('thead');
  tableJobs.appendChild(thead);

  let headerrow = document.createElement('tr');
  thead.appendChild(headerrow);

  let nameCell = createElement('th', null, 'Display Name');
  nameCell.setAttribute('scope', 'col');
  headerrow.appendChild(nameCell);

  let statusCell = createElement('th', null, 'Status');
  statusCell.setAttribute('scope', 'col');
  headerrow.appendChild(statusCell);

  let timeCell = createElement('th', null, 'Created');
  timeCell.setAttribute('scope', 'col');
  headerrow.appendChild(timeCell);

  let timetakenCell = createElement('th', null, 'Time Taken');
  timetakenCell.setAttribute('scope', 'col');
  headerrow.appendChild(timetakenCell);

  let resCell = createElement('th', null, 'Resources');
  resCell.setAttribute('scope', 'col');
  headerrow.appendChild(resCell);

  let instanceCell = createElement('th', null, 'Snapshot');
  instanceCell.setAttribute('scope', 'col');
  headerrow.appendChild(instanceCell);

  let errorCell = createElement('th', null, 'Error(s)');
  errorCell.setAttribute('scope', 'col');
  headerrow.appendChild(errorCell);

  let delCell = createElement("th", null, 'Delete')
  delCell.setAttribute('scope', 'col');
  headerrow.appendChild(delCell);

  let tbody = document.createElement('tbody');
  tableJobs.appendChild(tbody);

  for (const job of snapshotJobs)
  {
    let jobRow = document.createElement('tr');
    jobRow.setAttribute('key', job.id);
    tbody.appendChild(jobRow);

    let cell1 = createElement('td', null, job.displayName);
    jobRow.appendChild(cell1);

    var statusContent = job.status;

    if (statusContent == 'running')
    {
      statusContent = "running<br/><img src='running.gif' width='25' alt='running' />"
    }
    let cell3 = createElement('td', null, null);
    cell3.innerHTML = statusContent;
    jobRow.appendChild(cell3);

    const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    const date = new Date(job.createdDateTime); // UTC time
    const convertedDate = convertTimeZone(date, timezone);
    var timeParts = convertedDate.toLocaleString();
    let cell5 = createElement('td', null, timeParts);
    jobRow.appendChild(cell5);

    // If the job is completed, calculate the time it took to complete in seconds.
    var timeDiff = 'N/A';
    if (job.completedDateTime != '0001-01-01T00:00:00Z')
    {
      var created = new Date(job.createdDateTime);
      var completed = new Date(job.completedDateTime);

      var diff = completed.getTime() - created.getTime();
      var seconds = Math.floor(diff / (1000));
      timeDiff = seconds + "s";
    }
    let cell6 = createElement('td', null, timeDiff);
    jobRow.appendChild(cell6);

    var resourceList = "<ol>"
    for(const resource of job.resources)
    {
      resourceList += "<li>" + resource + "</li>";
    }
    resourceList += "</ol>";
    let cell7 = createElement('td', null, null);
    let span = createElement('span');
    span.innerHTML = resourceList;
    cell7.appendChild(span);
    jobRow.appendChild(cell7);

    let cell8 = createElement('td', null, null);
    if ('' != job.resourceLocation)
    {
      let file = createElement('span');
      var resourceLocationId = job.resourceLocation.split('(')[1];
      resourceLocationId = resourceLocationId.split(')')[0].replace("'", "").replace("'", "").replace(" ","");
      file.innerHTML = '<a href="#" onclick="getSnapshot(\'' + resourceLocationId + '\');"><img src="json.png" alt="View Snapshot" width="25" /></a>';
      cell8.appendChild(file);
    }
    jobRow.appendChild(cell8);

    let cellError = createElement('td', null, null);
    if (null != job.errorDetails && job.errorDetails.length != 0)
    {
      let errorIcon = createElement('span');
      errorIcon.innerHTML = '<a href="#" onclick="getSnapshotErrors(\'' + job.id + '\');"><img src="error.png" alt="View Errors" width="25" /></a>';
      cellError.appendChild(errorIcon);
    }
    else
    {
      cellError.innerHTML = "N/A"
    }
    jobRow.appendChild(cellError);

    // Only show the delete button if a monitor was created using credentials.
    let deletecell = createElement('td', null, null);
    if (null == job.createdBy.application.displayName && job.status != 'notStarted' && job.status != 'running')
    {
      let deleteSpan = createElement('span');
      deleteSpan.innerHTML = '<a href"#" onclick="deleteSnapshotJob(\'' + job.id + '\');"><img src="delete.png" alt="Delete Snapshot" width="25" /></a>';
      deletecell.appendChild(deleteSpan);
    }
    jobRow.appendChild(deletecell);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(div);
}

function showDrifts(drifts) {
  let div = document.createElement('div');
  let tableDrifts = createElement('table', 'table');
  div.appendChild(tableDrifts);

  let thead = document.createElement('thead');
  tableDrifts.appendChild(thead);

  let headerrow = document.createElement('tr');
  thead.appendChild(headerrow);

  let statusCell = createElement('th', null, 'Status');
  statusCell.setAttribute('scope', 'col');
  headerrow.appendChild(statusCell);

  let idCell = createElement('th', null, 'Id');
  idCell.setAttribute('scope', 'col');
  headerrow.appendChild(idCell);

  let timeCell = createElement('th', null, 'Time Reported');
  timeCell.setAttribute('scope', 'col');
  headerrow.appendChild(timeCell);

  let restypeCell = createElement('th', null, 'Resource Type');
  restypeCell.setAttribute('scope', 'col');
  headerrow.appendChild(restypeCell);

  let instanceCell = createElement('th', null, 'Instance Name');
  instanceCell.setAttribute('scope', 'col');
  headerrow.appendChild(instanceCell);

  let driftedPropCell = createElement('th', null, 'Drifted Properties');
  driftedPropCell.setAttribute('scope', 'col');
  headerrow.appendChild(driftedPropCell);

  let tbody = document.createElement('tbody');
  tableDrifts.appendChild(tbody);

  for (const drift of drifts)
  {
    let driftRow = document.createElement('tr');
    driftRow.setAttribute('key', drift.id);
    tbody.appendChild(driftRow);

    let cell1 = createElement('td', null, drift.status);
    driftRow.appendChild(cell1);

    let cell2 = createElement('td', null, drift.id);
    driftRow.appendChild(cell2);

    var timeParts = drift.firstReportedDateTime.split("T");
    var timeSubParts = timeParts[1].split(":")
    var runTimeValue = timeParts[0] + " " + timeSubParts[0] + ":" + timeSubParts[1];
    let cell3 = createElement('td', null, runTimeValue);
    driftRow.appendChild(cell3);

    let cell4 = createElement('td', null, drift.resourceType);
    driftRow.appendChild(cell4);

    let cell5 = createElement('td', null, JSON.stringify(drift.resourceInstanceIdentifier));
    driftRow.appendChild(cell5);

    var propertiesContent = "<ol>";
    for(const prop of drift.driftedProperties)
    {
      propertiesContent += "<li><strong>" + prop.propertyName + "</strong><br />Current Value: " + prop.currentValue + "<br />Desired Value:" + prop.desiredValue + "</li>";
    }
    propertiesContent += "</ul>";

    let cell6 = createElement('td', null, null);
    let contentHtml = createElement('span');
    contentHtml.innerHTML = propertiesContent;
    cell6.appendChild(contentHtml);
    driftRow.appendChild(cell6);
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(div);
}

function showMonitors(monitors, runs) {
  let div = document.createElement('div');

  div.appendChild(createElement('h1', 'mb-3', 'Monitors'));

  let newEventButton = createElement('button', 'btn btn-light btn-sm mb-3 btn-create', 'Create Monitor');
  newEventButton.setAttribute('onclick', 'showNewMonitorForm();');
  div.appendChild(newEventButton);

  let refreshIcon = createElement('span');
  refreshIcon.innerHTML = "&nbsp;&nbsp;<img src='refresh.jpg' alt='Refresh' onclick='getMonitors();' width='25' style='cursor:pointer;float:right;margin-top:-10px;' />";
  div.appendChild(refreshIcon);

  let table = createElement('table', 'table');
  div.appendChild(table);

  let thead = document.createElement('thead');
  table.appendChild(thead);

  let headerrow = document.createElement('tr');
  thead.appendChild(headerrow);

  let organizer = createElement('th', null, 'DisplayName');
  organizer.setAttribute('scope', 'col');
  headerrow.appendChild(organizer);

  let subject = createElement('th', null, 'Id');
  subject.setAttribute('scope', 'col');
  headerrow.appendChild(subject);

  let status = createElement('th', null, 'Status');
  status.setAttribute('scope', 'col');
  headerrow.appendChild(status);

  let deleteMonitor = createElement('th', null, 'Delete');
  deleteMonitor.setAttribute('scope', 'col');
  headerrow.appendChild(deleteMonitor);

  let tbody = document.createElement('tbody');
  table.appendChild(tbody);

  for (const monitor of monitors) {
    let monitorrow = document.createElement('tr');
    monitorrow.setAttribute('key', monitor.displayname);
    tbody.appendChild(monitorrow);

    let namecell = createElement('td', 'boldheader', monitor.displayName);
    monitorrow.appendChild(namecell);

    let idcell = createElement('td', 'boldheader', monitor.id);
    monitorrow.appendChild(idcell);

    let statuscell = createElement('td', 'boldheader', monitor.status);
    monitorrow.appendChild(statuscell);

    let deletecell = createElement('td', 'boldheader', null);

    // Only show the delete button if a monitor was created using credentials.
    if (null == monitor.createdBy.application.displayName)
    {
      let deleteSpan = createElement('span');
      deleteSpan.innerHTML = '<a href"#" onclick="deleteMonitor(\'' + monitor.id + '\');"><img src="delete.png" alt="Delete Monitor" width="25" /></a>';
      deletecell.appendChild(deleteSpan);
    }
    monitorrow.appendChild(deletecell);

    for (const run of runs)
    {
      if (run.monitorId == monitor.id)
      {
        try
        {
          let runrow = document.createElement('tr');
          runrow.setAttribute('key', run.id);
          tbody.appendChild(runrow);

          // Usage
          const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
          const date = new Date(run.runInitiationDateTime); // UTC time
          const convertedDate = convertTimeZone(date, timezone);
          var timeParts = convertedDate.toLocaleString();
          let runtimecell = createElement('td', 'subValue', timeParts);
          runrow.appendChild(runtimecell);

          let runstatuscell = createElement('td', 'subValue', run.runStatus);
          runrow.appendChild(runstatuscell);

          var timeDiff = 'N/A';
          if (monitor.runCompletionDateTime != '0001-01-01T00:00:00Z')
          {
            var created = new Date(run.runInitiationDateTime);
            var completed = new Date(run.runCompletionDateTime);

            var diff = completed.getTime() - created.getTime();
            var seconds = Math.floor(diff / (1000));
            diff -= seconds * (1000);
            timeDiff = diff + "s";
          }

          let execTimeCell = createElement('td', null, timeDiff);
          runrow.appendChild(execTimeCell);

          let driftcell
          if (run.driftsCount > 0)
          {
            driftcell = createElement('td', 'subValueRed', null);
            let driftLink = createElement('a', null, run.driftsCount + " Drift(s) Detected")
            driftLink.setAttribute('onclick', 'getDrifts("' + monitor.id + '");');
            driftcell.appendChild(driftLink);
          }
          else
          {
            driftcell = createElement('td', 'subValueGreen', run.driftsCount + " Drift(s) Detected");
          }
          runrow.appendChild(driftcell);

          if (null != run.errorDetails && run.errorDetails.length != 0)
          {
            let cellError = createElement('td', null, null);
            let errorIcon = createElement('span');
            errorIcon.innerHTML = '<a href="#" onclick="getMonitorRunErrors(\'' + job.id + '\');"><img src="error.png" alt="View Errors" width="25" /></a>';
            cellError.appendChild(errorIcon);
            runrow.appendChild(cellError);
          }
        }
        catch{}
      }
    }
  }

  mainContainer.innerHTML = '';
  mainContainer.appendChild(div);
}

function convertTimeZone(date, timeZone) {
  return new Date(date.toLocaleString('en-US', { timeZone: timeZone }));
}

updatePage(Views.home);
