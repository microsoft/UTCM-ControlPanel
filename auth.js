
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <authInitSnippet>
// Create the main MSAL instance
// configuration parameters are located in config.js
const msalClient = new msal.PublicClientApplication(msalConfig);
// </authInitSnippet>

// <checkAuthSnippet>
// Handle the auth response when the page loads inside a popup or redirect.
// Without this, the popup doesn't close after login and triggers block_nested_popups.
msalClient.handleRedirectPromise()
  .then(function (response) {
    if (response) {
      msalClient.setActiveAccount(response.account);
      initializeGraphClient(msalClient, response.account, msalRequest.scopes);
    } else {
      // No redirect response — check for an already logged-in user
      var account = msalClient.getActiveAccount();
      if (account) {
        initializeGraphClient(msalClient, account, msalRequest.scopes);
      }
    }
  })
  .catch(function (error) {
    console.error('handleRedirectPromise error:', error);
  });
// </checkAuthSnippet>

// <signInSnippet>
async function signIn() {
    // Login
    try {
      // Use MSAL to login
      const authResult = await msalClient.loginPopup(msalRequest);
      console.log('id_token acquired at: ' + new Date().toString());
  
      msalClient.setActiveAccount(authResult.account);
  
      // Initialize the Graph client
      initializeGraphClient(msalClient, authResult.account, msalRequest.scopes);
  
      // Get the user's profile from Graph
      const user = await getUser();
      const drifts = await getAllDrifts();
      // Save the profile in session
      sessionStorage.setItem('graphUser', JSON.stringify(user));

      try
      {
        const photo = await getPhoto();
        var urlCreator = window.URL || window.webkitURL;
        var photoUrl = urlCreator.createObjectURL(photo);
        sessionStorage.setItem('graphPhoto', photoUrl);
      }
      catch{}
      sessionStorage.setItem('drifts', JSON.stringify(drifts));
      updatePage(Views.home);
    } catch (error) {
      console.log(error);
      updatePage(Views.error, {
        message: 'Error logging in',
        debug: error
      });
    }
  }
  // </signInSnippet>
  
  // <signOutSnippet>
function signOut() {
    sessionStorage.removeItem('graphUser');
    msalClient.logout();
  }
  // </signOutSnippet>
