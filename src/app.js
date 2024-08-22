import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/lib/src/authentication/azureTokenCredentials/index.js";

import * as changeKeys from "change-case/keys"; //package: to change key case to camel
import { Keys } from "./variables.js";

//List Names
const LOAN_REQUEST_LIST = "LoanRequestApplications";
const LOAN_FINALISATION_LIST_NAME = "LoanRequestUnderwritingFinalisation";

const credential = new ClientSecretCredential(
  Keys.TenantID,
  Keys.ClientID,
  Keys.SecretValue
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ["https://graph.microsoft.com/.default"],
});

const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

//REQUEST 1: Get data from the "Loan Request Applications" list
graphClient
  .api(`sites/${Keys.SiteID}/lists/${LOAN_REQUEST_LIST}/items`)
  .filter(
    "fields/AccountNumber eq '25485' AND fields/Branch/Title eq 'Port of Spain'"
  )
  .expand(
    "fields($select=ApplicationID,AccountNumber,Branch,Name,ApplicationStatus,CashNowRequired,Modified)"
  )
  .get()
  .then((response) =>
    console.log(response.value.map((item) => changeKeys.camelCase(item.fields)))
  )
  .catch((e) => console.log(e));

//REQUEST 2: Get data from the "Loan Request Underwriting Finalisation" list
graphClient
  .api(`sites/${Keys.SiteID}/lists/${LOAN_FINALISATION_LIST_NAME}/items`)
  .filter("fields/ApplicationID eq '1100'")
  .expand(
    "fields($select=WithinEligibility,WithinCreditPolicyDSR,WithinCreditPolicyLoanAmount,FullySecured,LoanAmount,LoanType,LoanTerm,Modified)"
  )
  .get()
  .then((response) =>
    console.log(response.value.map((item) => changeKeys.camelCase(item.fields)))
  )
  .catch((e) => console.log(e));
