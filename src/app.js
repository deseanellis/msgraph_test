import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/lib/src/authentication/azureTokenCredentials/index.js";

import * as changeKeys from "change-case/keys"; //package: to change key case to camel
import { Keys } from "./variables.js";

const LIST_NAME = "LoanRequestApplications";

const credential = new ClientSecretCredential(
  Keys.TenantID,
  Keys.ClientID,
  Keys.Secret
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ["https://graph.microsoft.com/.default"],
});

const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

graphClient
  .api(`sites/${Keys.SiteID}/lists/${LIST_NAME}/items`)
  .filter("fields/LoanType/Title eq 'Business'")
  .expand(
    "fields($select=ApplicationID,AccountNumber,Name,Branch,Status,Created,Modified)"
  )
  .get()
  .then((response) =>
    console.log(response.value.map((item) => changeKeys.camelCase(item.fields)))
  )
  .catch((e) => console.error(e));
