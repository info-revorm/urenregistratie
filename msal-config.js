/* ============================================================
   revorm urenregistratie — Microsoft-login configuratie
   ------------------------------------------------------------
   Vul hieronder de twee waarden in die je krijgt bij de
   Entra ID app-registratie (zie SETUP.md, stap A):

   1) CLIENT_ID  = "Application (client) ID"
   2) TENANT_ID  = "Directory (tenant) ID"

   Verder hoef je niets aan te passen.
   ============================================================ */

const CLIENT_ID = "fbc6e099-71ff-4e20-a904-7e092eb794d2";
const TENANT_ID = "8b834b88-bc34-4126-8d94-cfd33d6db6bc";

/* Vaste instellingen — niet wijzigen tenzij je weet wat je doet */
const MSAL_CONFIG = {
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/" + TENANT_ID,
    // Moet exact overeenkomen met de Redirect-URI in Entra (incl. afsluitende /)
    redirectUri: window.location.origin + window.location.pathname,
    postLogoutRedirectUri: window.location.origin + window.location.pathname,
    navigateToLoginRequestUrl: true
  },
  cache: {
    // localStorage is nodig zodat de login bewaard blijft op iPhone/Safari
    cacheLocation: "localStorage",
    temporaryCacheLocation: "localStorage"
  },
  system: {
    allowRedirectInIframe: false
  }
};

/* Minimale rechten: alleen de eigen app-map in OneDrive + wie ben ik */
const GRAPH_SCOPES = ["Files.ReadWrite.AppFolder", "User.Read"];

/* Bestandsnaam van de uren-data in OneDrive (app-map /Apps/<app naam>/) */
const GRAPH_FILE = "revorm_uren_v1.json";

/* Snelle check of de configuratie is ingevuld */
const MSAL_CONFIGURED =
  CLIENT_ID.indexOf("PLAK-HIER") === -1 && TENANT_ID.indexOf("PLAK-HIER") === -1;
