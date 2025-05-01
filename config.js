// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const msalConfig = {
  auth: {
    clientId: '006ad752-04b5-4292-9a9e-8d08668edf31',
    redirectUri: 'https://microsoft.github.io/UTCM-ControlPanel/'
  }
};

const msalRequest = {
  scopes: [
    'user.read',
    'ConfigurationMonitoring.ReadWrite.All'
  ]
}
