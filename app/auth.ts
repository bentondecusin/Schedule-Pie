/**
 * The following codes are adapted from Microsoft
 * See more at https://github.com/MicrosoftDocs/mslearn-retrieve-m365-data-with-msgraph-quickstart and adapted
 */

import "isomorphic-fetch";
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  SilentRequest,
} from "@azure/msal-browser";

//MSAL configuration
export const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID || "",
    // authority: authority ID
    // redirectUri: "https://localhost:8080",
  },
};

export const msalRequest = {
  scopes: ["Calendars.read", "user.read"],
};

export function ensureScope(scope: string) {
  if (
    !msalRequest.scopes.some(
      (s: string) => s.toLowerCase() === scope.toLowerCase()
    )
  ) {
    msalRequest.scopes.push(scope);
  }
}

//Initialize MSAL client
export const msalClient: PublicClientApplication = new PublicClientApplication(
  msalConfig
);

// Log the user in
export async function signIn() {
  const authResult = await msalClient
    .initialize()
    .then(() => msalClient.loginPopup(msalRequest));
  sessionStorage.setItem("msalAccount", authResult.account.username);
}

//Get token from Graph
export async function getToken() {
  let account = sessionStorage.getItem("msalAccount");
  if (!account) {
    throw new Error(
      "User info cleared from session. Please sign out and sign in again."
    );
  }
  try {
    // First, attempt to get the token silently
    const msalAccount = msalClient.getAccountByUsername(account);
    // If msal account does not exist
    if (!msalAccount) throw Error("account DNE");

    const silentRequest: SilentRequest = {
      scopes: msalRequest.scopes,
      account: msalAccount,
    };
    const silentResult = await msalClient.acquireTokenSilent(silentRequest);

    return silentResult.accessToken;
  } catch (silentError) {
    // If silent requests fails with InteractionRequiredAuthError,
    // attempt to get the token interactively
    if (silentError instanceof InteractionRequiredAuthError) {
      const interactiveResult = await msalClient.acquireTokenPopup(msalRequest);
      return interactiveResult.accessToken;
    } else {
      throw silentError;
    }
  }
}
