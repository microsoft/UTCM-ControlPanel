// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const msalConfig = {
  auth: {
    clientId: 'c6336297-ebbf-4cb2-bf22-5304363a3618',//'006ad752-04b5-4292-9a9e-8d08668edf31',
    redirectUri: 'http://localhost:8080'
  }
};

const msalRequest = {
  scopes: [
    'user.read',
    'ConfigurationMonitoring.ReadWrite.All'
  ]
}
