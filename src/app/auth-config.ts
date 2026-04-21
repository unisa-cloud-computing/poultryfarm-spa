import {InteractionType, type IPublicClientApplication, LogLevel, PublicClientApplication} from '@azure/msal-browser';
import {MsalGuardConfiguration, MsalInterceptorConfiguration} from '@azure/msal-angular';


// L'APP_ID della app registration "poultryfarm-api" che creerai in Entra
// Sostituisci con il valore reale dopo averla creata
const API_APP_ID = '19c1ec73-39ac-4e18-aead-31ee6f3d3230';

// L'URL base di APIM (lo conoscerai dopo aver creato l'istanza APIM)
const APIM_BASE_URL = 'https://apim-poultryfarm.azure-api.net';


export const msalConfig = {
  auth: {
    // ID dell'applicazione registrata in Azure AD e dice ad Entra ID qual è il client che sta chiedendo i token.
    clientId: '3341bcc2-589a-4911-bbea-c32fa387491d',
    // endpoint dell'external tenant e serve a MSAL per indirizzare l'utente alla login e da dove accettare i token.
    authority: 'https://cloudsaasmanagement.ciamlogin.com/c7f158c6-9d4c-4789-ae4e-664fbdb8f405',
    // url di DEV utile per redirect dove Entra ID rimanda l'utente dopo la login. Questo URL deve essere registrato anche nell'APP Registration.
    redirectUri: 'https://witty-rock-04434f010.7.azurestaticapps.net/',
    postLogoutRedirectUri: 'https://witty-rock-04434f010.7.azurestaticapps.net/logout-success', // oppure una pagina libera da MsalGuard /welcome
  },
  cache: {
    // dove MSAL memorizza i token. localStorage persiste anche dopo la chiusura del browser, sessionStorage invece no.
    cacheLocation: 'localStorage' as const,
    // se true MSAL memorizza anche lo stato dell'autenticazione nei cookie, utile per browser che non supportano localStorage o sessionStorage.
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      // callback chiamata da MSAL per loggare eventi
      loggerCallback: (level: LogLevel, message: string) => {
        console.log(message);
      },
      logLevel: LogLevel.Info,
      // se true, MSAL logga anche informazioni personali come username o token, utile per debug ma da disabilitare in produzione.
      piiLoggingEnabled: false,
    },
  },
};

// loginRequest è l'oggeto che descrive cosa chiedere a Entra ID durante la login
// oggetto usato da MSALGuardConfigFactory
// openid, profile, email → servono per l'ID token
// api://<APP_ID>/user_impersonation → dice a Entra di emettere anche
// un access token per poultryfarm-api, con aud = api://<APP_ID>

export const loginRequest = {
  scopes: ['openid', 'profile', 'email', `api://${API_APP_ID}/user_impersonation`]
};

// MSAL instance
export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication(msalConfig);
}

// come MSAL si comporta quando blocca una rotta
export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return {
    // se utente non autenticato, allora MSAL deve fare redirect per la login
    interactionType: InteractionType.Redirect,
    authRequest: loginRequest
  };
}

// come MSAL si comporta quando intercetta una chiamata HTTP verso un API protetta
export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  return {
    // Se per ottenere un token per una certa API manca il consenso, allora MSAL deve fare redirect per chiedere il consenso all'utente
    interactionType: InteractionType.Redirect,
    // Dice a MsalInterceptor: "ogni volta che Angular fa una chiamata HTTP
    // verso un URL che inizia con APIM_BASE_URL, allega automaticamente
    // l'access token con lo scope user_impersonation nell'header Authorization"
    protectedResourceMap: new Map<string, string[]>([
      [APIM_BASE_URL, [`api://${API_APP_ID}/user_impersonation`]]
    ])
  };
}
