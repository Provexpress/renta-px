(function () {
  let msalInstance = null;

  function getScopes() {
    return APP_CONFIG.graph.scopes || ["User.Read"];
  }

  function getAuthority() {
    return `https://login.microsoftonline.com/${APP_CONFIG.msal.tenantId}`;
  }

  async function initializeMsal() {
    if (!window.msal) {
      throw new Error("No se pudo cargar MSAL.js desde el CDN.");
    }

    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: APP_CONFIG.msal.clientId,
        authority: getAuthority(),
        redirectUri: APP_CONFIG.msal.redirectUri
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
      }
    });

    if (typeof msalInstance.initialize === "function") {
      await msalInstance.initialize();
    }

    const response = await msalInstance.handleRedirectPromise();
    if (response && response.account) {
      msalInstance.setActiveAccount(response.account);
    }

    const accounts = msalInstance.getAllAccounts();
    if (!msalInstance.getActiveAccount() && accounts.length) {
      msalInstance.setActiveAccount(accounts[0]);
    }

    return getActiveAccount();
  }

  async function login() {
    const request = { scopes: getScopes(), prompt: "select_account" };

    try {
      const response = await msalInstance.loginPopup(request);
      msalInstance.setActiveAccount(response.account);
      return response.account;
    } catch (error) {
      if (error && error.errorCode === "popup_window_error") {
        await msalInstance.loginRedirect(request);
        return null;
      }
      throw error;
    }
  }

  function logout() {
    const account = getActiveAccount();
    return msalInstance.logoutPopup({
      account,
      postLogoutRedirectUri: APP_CONFIG.msal.redirectUri
    });
  }

  function getActiveAccount() {
    if (!msalInstance) return null;
    return msalInstance.getActiveAccount();
  }

  async function getAccessToken() {
    const account = getActiveAccount();
    if (!account) {
      throw new Error("No hay una cuenta activa.");
    }

    const request = {
      account,
      scopes: getScopes()
    };

    try {
      const response = await msalInstance.acquireTokenSilent(request);
      return response.accessToken;
    } catch (error) {
      const response = await msalInstance.acquireTokenPopup(request);
      return response.accessToken;
    }
  }

  async function getUserProfile() {
    const token = await getAccessToken();
    const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName", {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    if (!response.ok) {
      throw new Error("No se pudo obtener el perfil del usuario.");
    }

    return response.json();
  }

  window.AuthService = {
    initializeMsal,
    login,
    logout,
    getActiveAccount,
    getAccessToken,
    getUserProfile
  };
})();
