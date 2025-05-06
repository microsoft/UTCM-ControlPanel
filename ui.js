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
  var termsLink = createElement('a', 'btn btn-link nav-link', 'Terms & Conditions');
  termsLink.setAttribute('href','https://aka.ms/M365CCPTandCs');
  accountNav.appendChild(termsLink);
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

function updatePage(view, data, data2, graphURI) {
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
      showMonitors(data, data2, graphURI);
      break;
    case Views.drifts:
      showDrifts(data, graphURI);
      break;
    case Views.snapshots:
      showSnapshotJobs(data, graphURI);
      break;
    case Views.snapshotInfo:
      showSnapshot(data, graphURI);
      break;
    case Views.snapshotErrors:
      showSnapshotErrors(data, graphURI);
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

  var allResources = ["microsoft.exchange.accepteddomain",
        "microsoft.exchange.activesyncdeviceaccessrule",
        "microsoft.exchange.addressbookpolicy",
        "microsoft.exchange.antiphishpolicy",
        "microsoft.exchange.antiphishrule",
        "microsoft.exchange.applicationaccesspolicy",
        "microsoft.exchange.atppolicyforo365",
        "microsoft.exchange.authenticationpolicy",
        "microsoft.exchange.authenticationpolicyassignment",
        "microsoft.exchange.availabilityaddressspace",
        "microsoft.exchange.availabilityconfig",
        "microsoft.exchange.calendarprocessing",
        "microsoft.exchange.casmailboxplan",
        "microsoft.exchange.casmailboxsettings",
        "microsoft.exchange.dataclassification",
        "microsoft.exchange.dataencryptionpolicy",
        "microsoft.exchange.distributiongroup",
        "microsoft.exchange.dkimsigningconfig",
        "microsoft.exchange.emailaddresspolicy",
        "microsoft.exchange.groupsettings",
        "microsoft.exchange.hostedconnectionfilterpolicy",
        "microsoft.exchange.hostedcontentfilterpolicy",
        "microsoft.exchange.hostedcontentfilterrule",
        "microsoft.exchange.hostedoutboundspamfilterpolicy",
        "microsoft.exchange.hostedoutboundspamfilterrule",
        "microsoft.exchange.inboundconnector",
        "microsoft.exchange.intraorganizationconnector",
        "microsoft.exchange.irmconfiguration",
        "microsoft.exchange.journalrule",
        "microsoft.exchange.mailboxautoreplyconfiguration",
        "microsoft.exchange.mailboxcalendarfolder",
        "microsoft.exchange.mailboxpermission",
        "microsoft.exchange.mailboxplan",
        "microsoft.exchange.mailboxsettings",
        "microsoft.exchange.mailcontact",
        "microsoft.exchange.mailtips",
        "microsoft.exchange.malwarefilterpolicy",
        "microsoft.exchange.malwarefilterrule",
        "microsoft.exchange.managementrole",
        "microsoft.exchange.managementroleassignment",
        "microsoft.exchange.messageclassification",
        "microsoft.exchange.mobiledevicemailboxpolicy",
        "microsoft.exchange.omeconfiguration",
        "microsoft.exchange.onpremisesorganization",
        "microsoft.exchange.organizationconfig",
        "microsoft.exchange.organizationrelationship",
        "microsoft.exchange.outboundconnector",
        "microsoft.exchange.owamailboxpolicy",
        "microsoft.exchange.partnerapplication",
        "microsoft.exchange.perimeterconfiguration",
        "microsoft.exchange.place",
        "microsoft.exchange.policytipconfig",
        "microsoft.exchange.quarantinepolicy",
        "microsoft.exchange.recipientpermission",
        "microsoft.exchange.remotedomain",
        "microsoft.exchange.reportsubmissionpolicy",
        "microsoft.exchange.reportsubmissionrule",
        "microsoft.exchange.resourceconfiguration",
        "microsoft.exchange.roleassignmentpolicy",
        "microsoft.exchange.rolegroup",
        "microsoft.exchange.safeattachmentpolicy",
        "microsoft.exchange.safeattachmentrule",
        "microsoft.exchange.safelinkspolicy",
        "microsoft.exchange.safelinksrule",
        "microsoft.exchange.sharedmailbox",
        "microsoft.exchange.sharingpolicy",
        "microsoft.exchange.transportconfig",
        "microsoft.exchange.transportrule",
        "microsoft.entra.administrativeunit",
        "microsoft.entra.application",
        "microsoft.entra.authenticationcontextclassreference",
        "microsoft.entra.authenticationmethodpolicy",
        "microsoft.entra.authenticationmethodpolicyauthenticator",
        "microsoft.entra.authenticationmethodpolicyemail",
        "microsoft.entra.authenticationmethodpolicyfido2",
        "microsoft.entra.authenticationmethodpolicysms",
        "microsoft.entra.authenticationmethodpolicysoftware",
        "microsoft.entra.authenticationmethodpolicytemporary",
        "microsoft.entra.authenticationmethodpolicyvoice",
        "microsoft.entra.authenticationmethodpolicyx509",
        "microsoft.entra.authenticationstrengthpolicy",
        "microsoft.entra.authorizationpolicy",
        "microsoft.entra.conditionalaccesspolicy",
        "microsoft.entra.crosstenantaccesspolicy",
        "microsoft.entra.crosstenantaccesspolicyconfigurationdefault",
        "microsoft.entra.crosstenantaccesspolicyconfigurationpartner",
        "microsoft.entra.entitlementmanagementaccesspackage",
        "microsoft.entra.entitlementmanagementaccesspackageassignmentpolicy",
        "microsoft.entra.entitlementmanagementaccesspackagecatalog",
        "microsoft.entra.entitlementmanagementaccesspackagecatalogresource",
        "microsoft.entra.entitlementmanagementconnectedorganization",
        "microsoft.entra.externalidentitypolicy",
        "microsoft.entra.group",
        "microsoft.entra.grouplifecyclepolicy",
        "microsoft.entra.namedlocationpolicy",
        "microsoft.entra.roledefinition",
        "microsoft.entra.roleeligibilityschedulerequest",
        "microsoft.entra.rolesetting",
        "microsoft.entra.securitydefaults",
        "microsoft.entra.serviceprincipal",
        "microsoft.entra.socialidentityprovider",
        "microsoft.entra.tenantdetails",
        "microsoft.entra.tokenlifetimepolicy",
        "microsoft.entra.user",
        "microsoft.teams.apppermissionpolicy",
        "microsoft.teams.appsetuppolicy",
        "microsoft.teams.callingpolicy",
        "microsoft.teams.callparkpolicy",
        "microsoft.teams.channelspolicy",
        "microsoft.teams.clientconfiguration",
        "microsoft.teams.cortanapolicy",
        "microsoft.teams.dialinconferencingtenantsettings",
        "microsoft.teams.enhancedencryptionpolicy",
        "microsoft.teams.eventspolicy",
        "microsoft.teams.federationconfiguration",
        "microsoft.teams.feedbackpolicy",
        "microsoft.teams.filespolicy",
        "microsoft.teams.guestcallingconfiguration",
        "microsoft.teams.guestmeetingconfiguration",
        "microsoft.teams.guestmessagingconfiguration",
        "microsoft.teams.ipphonepolicy",
        "microsoft.teams.meetingbroadcastpolicy",
        "microsoft.teams.meetingconfiguration",
        "microsoft.teams.meetingpolicy",
        "microsoft.teams.messagingpolicy",
        "microsoft.teams.mobilitypolicy",
        "microsoft.teams.networkroamingpolicy",
        "microsoft.teams.onlinevoicemailpolicy",
        "microsoft.teams.shiftspolicy",
        "microsoft.teams.templatespolicy",
        "microsoft.teams.tenantnetworkregion",
        "microsoft.teams.tenantnetworksite",
        "microsoft.teams.tenantnetworksubnet",
        "microsoft.teams.tenanttrustedipaddress",
        "microsoft.teams.translationrule",
        "microsoft.teams.unassignednumbertreatment",
        "microsoft.teams.upgradeconfiguration",
        "microsoft.teams.vdipolicy",
        "microsoft.teams.voiceroute",
        "microsoft.teams.voiceroutingpolicy",
        "microsoft.teams.workloadpolicy",
        "microsoft.teams.callholdpolicy",
        "microsoft.teams.updatemanagementpolicy",
        "microsoft.teams.pstnusage",
        "microsoft.teams.upgradepolicy",
        "microsoft.teams.tenantdialplan",
        "microsoft.teams.meetingbroadcastconfiguration",
        "microsoft.intune.deviceconfigurationidentityprotectionpolicywindows10",
        "microsoft.intune.devicecompliancepolicyandroid",
        "microsoft.intune.deviceconfigurationendpointprotectionpolicywindows10",
        "microsoft.intune.deviceenrollmentlimitrestriction",
        "microsoft.intune.accountprotectionpolicy",
        "microsoft.intune.deviceconfigurationemailprofilepolicywindows10",
        "microsoft.intune.windowsupdateforbusinessringupdateprofilewindows10",
        "microsoft.intune.applicationcontrolpolicywindows10",
        "microsoft.intune.deviceandappmanagementassignmentfilter",
        "microsoft.intune.deviceconfigurationfirmwareinterfacepolicywindows10",
        "microsoft.intune.deviceconfigurationsharedmultidevicepolicywindows10",
        "microsoft.intune.accountprotectionlocalusergroupmembershippolicy",
        "microsoft.intune.deviceconfigurationwindowsteampolicywindows10",
        "microsoft.intune.deviceconfigurationtrustedcertificatepolicywindows10",
        "microsoft.intune.devicecleanuprule",
        "microsoft.intune.wificonfigurationpolicyandroidopensourceproject",
        "microsoft.intune.deviceconfigurationpolicyandroidworkprofile",
        "microsoft.intune.windowsinformationprotectionpolicywindows10mdmenrolled",
        "microsoft.intune.deviceconfigurationhealthmonitoringconfigurationpolicywindows10",
        "microsoft.intune.deviceconfigurationvpnpolicywindows10",
        "microsoft.intune.appprotectionpolicyios",
        "microsoft.intune.roledefinition",
        "microsoft.intune.deviceconfigurationsecureassessmentpolicywindows10",
        "microsoft.intune.devicecompliancepolicymacos",
        "microsoft.intune.deviceconfigurationpolicymacos",
        "microsoft.intune.deviceconfigurationdomainjoinpolicywindows10",
        "microsoft.intune.deviceconfigurationdefenderforendpointonboardingpolicywindows10",
        "microsoft.intune.deviceconfigurationpkcscertificatepolicywindows10",
        "microsoft.intune.deviceconfigurationimportedpfxcertificatepolicywindows10",
        "microsoft.intune.devicecompliancepolicyandroiddeviceowner",
        "microsoft.intune.wificonfigurationpolicywindows10",
        "microsoft.intune.devicecompliancepolicyandroidworkprofile",
        "microsoft.intune.deviceconfigurationnetworkboundarypolicywindows10",
        "microsoft.intune.devicecompliancepolicyios",
        "microsoft.intune.wificonfigurationpolicymacos",
        "microsoft.intune.devicecompliancepolicywindows10",
        "microsoft.intune.deviceconfigurationpolicywindows10",
        "microsoft.intune.devicecategory",
        "microsoft.intune.deviceconfigurationpolicyandroiddeviceowner",
        "microsoft.intune.deviceconfigurationpolicyandroidopensourceproject",
        "microsoft.intune.deviceconfigurationpolicyios",
        "microsoft.intune.wificonfigurationpolicyios",
        "microsoft.intune.antiviruspolicywindows10settingcatalog",
        "microsoft.intune.appprotectionpolicyandroid",
        "microsoft.intune.attacksurfacereductionrulespolicywindows10configmanager",
        "microsoft.intune.deviceconfigurationadministrativetemplatepolicywindows10",
        "microsoft.intune.deviceconfigurationcustompolicywindows10",
        "microsoft.intune.deviceconfigurationdeliveryoptimizationpolicywindows10",
        "microsoft.intune.deviceconfigurationkioskpolicywindows10",
        "microsoft.intune.deviceconfigurationpolicyandroiddeviceadministrator",
        "microsoft.intune.deviceconfigurationscepcertificatepolicywindows10",
        "microsoft.intune.deviceconfigurationwirednetworkpolicywindows10",
        "microsoft.intune.deviceenrollmentplatformrestriction",
        "microsoft.intune.deviceenrollmentstatuspagewindows10",
        "microsoft.intune.endpointdetectionandresponsepolicywindows10",
        "microsoft.intune.exploitprotectionpolicywindows10settingcatalog",
        "microsoft.intune.policysets",
        "microsoft.intune.roleassignment",
        "microsoft.intune.settingcatalogcustompolicywindows10",
        "microsoft.intune.wificonfigurationpolicyandroiddeviceadministrator",
        "microsoft.intune.wificonfigurationpolicyandroidenterprisedeviceowner",
        "microsoft.intune.wificonfigurationpolicyandroidforwork",
        "microsoft.intune.windowsautopilotdeploymentprofileazureadhybridjoined",
        "microsoft.intune.windowsautopilotdeploymentprofileazureadjoined",
        "microsoft.intune.windowsupdateforbusinessfeatureupdateprofilewindows10",
        "microsoft.securityandcompliance.auditconfigurationpolicy",
        "microsoft.securityandcompliance.autosensitivitylabelpolicy",
        "microsoft.securityandcompliance.caseholdpolicy",
        "microsoft.securityandcompliance.caseholdrule",
        "microsoft.securityandcompliance.compliancecase",
        "microsoft.securityandcompliance.compliancesearch",
        "microsoft.securityandcompliance.compliancesearchaction",
        "microsoft.securityandcompliance.compliancetag",
        "microsoft.securityandcompliance.deviceconditionalaccesspolicy",
        "microsoft.securityandcompliance.deviceconfigurationpolicy",
        "microsoft.securityandcompliance.dlpcompliancepolicy",
        "microsoft.securityandcompliance.fileplanpropertyauthority",
        "microsoft.securityandcompliance.fileplanpropertycategory",
        "microsoft.securityandcompliance.fileplanpropertycitation",
        "microsoft.securityandcompliance.fileplanpropertydepartment",
        "microsoft.securityandcompliance.fileplanpropertyreferenceid",
        "microsoft.securityandcompliance.fileplanpropertysubcategory",
        "microsoft.securityandcompliance.protectionalert",
        "microsoft.securityandcompliance.retentioncompliancepolicy",
        "microsoft.securityandcompliance.retentioncompliancerule",
        "microsoft.securityandcompliance.retentioneventtype",
        "microsoft.securityandcompliance.securityfilter",
        "microsoft.securityandcompliance.supervisoryreviewpolicy",
        "microsoft.securityandcompliance.supervisoryreviewrule"];
  let resourceDDL = createElement('select');
  resourceDDL.setAttribute('multiple', "true");
  resourceDDL.setAttribute('id', 'mon-resources');
  resourceDDL.setAttribute('name', 'mon-resources');
  resourceDDL.style.height = "500px";
  resourceDDL.style.width = "100%";
  for(const resource of allResources.sort())
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

function showGraphBanner(uri)
{
  let divGraphBanner = document.createElement('div', 'graph-banner', uri);
  divGraphBanner.id = "divGraphBanner";
  divGraphBanner.innerHTML = "<strong>Associated Graph call:</strong> " + uri;
  mainContainer.appendChild(divGraphBanner);
}

function showSnapshot(data, graphURI) {

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

  delete data.id;
  delete data['@odata.context'];
  var snapshotContent = JSON.stringify(data, null, 4);
  contentInput.innerHTML = snapshotContent;
  contentGroup.appendChild(contentInput);

  mainContainer.innerHTML = '';
  mainContainer.appendChild(form);
  showGraphBanner(graphURI);
}

function showSnapshotErrors(snapshotErrors, graphURI) {

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
  showGraphBanner(graphURI);
}

function showSnapshotJobs(snapshotJobs, graphURI) {

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
  showGraphBanner(graphURI);
}

function showDrifts(drifts, graphURI) {
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
  showGraphBanner(graphURI);
}

function showMonitors(monitors, runs, graphURI) {
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
  showGraphBanner(graphURI);
}

function convertTimeZone(date, timeZone) {
  return new Date(date.toLocaleString('en-US', { timeZone: timeZone }));
}

updatePage(Views.home);
