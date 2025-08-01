// Authentication SSO

// TODO Change parameters below to your client id, tenant id and token endpoint.
// See documentation (https://learn.microsoft.com/en-us/microsoft-copilot-studio/configure-sso?tabs=webApp)
const clientId = "310f8b9b-5434-4416-a44e-4f9e7499cfe7";
const tenantId = "5db30893-fb80-480e-be22-911156d0e5e0";
//Staging Production DirectLine token
// const tokenEndpoint =
//   "https://8e0a9f6142454f9b8a1a3247bca2d2.35.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cr2a3_salesAgent/directline/token?api-version=2022-03-01-preview"; // you can find the token URL via the Mobile app channel configuration
//Production Token Directline
const tokenEndpoint =
  "https://c81dcd50c0fd49b2b9b4661ef9fc17.ed.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cr262_agent/directline/token?api-version=2022-03-01-preview"; // you can find the token URL via the Mobile app channel configuration

// Config object to be passed to Msal on creation
const msalConfig = {
  auth: {
    clientId: clientId,
    authority: "https://login.microsoftonline.com/" + tenantId,
    // redirectUri: "https://grazitti936.crm8.dynamics.com/", //Staging
    redirectUri: "https://grazitti934.crm8.dynamics.com/", //Production
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: true, // Set this to 'true' if you are having issues on IE11 or Edge
  },
};
// Handle login request after user clicks on login button
const msalInstance = new msal.PublicClientApplication(msalConfig);

async function onSignInClick() {
  // Add here scopes for id token to be used at MS Identity Platform endpoints.
  const loginRequest = {
    scopes: ["User.Read", "openid", "profile"],
  };
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      console.log("Account found no login initiated", accounts[0]);
      msalInstance.setActiveAccount(accounts[0]);
      user = accounts[0];
      await renderChatWidget();
    } else {
      try {
        const response = await msalInstance.ssoSilent({
          scopes: ["User.Read", "openid", "profile"], // or your API scope
        });
        user = response.account;

        msalInstance.setActiveAccount(user);
        console.log("Account added after successfull silent SSO", user);
        await renderChatWidget();
      } catch (err) {
        console.log("❌ Silent SSO failed: Initiating Login Popup Login",err);
        try {
          // No account found, fallback to loginPopup
          console.log("Account not found login initiated");
          const loginResponse = await msalInstance.loginPopup(loginRequest);
          // const loginResponse = await msalInstance.acquireTokenSilent(loginRequest);

          user = loginResponse.account;
          msalInstance.setActiveAccount(user);
          await renderChatWidget();
        } catch (err) {
          console.log(err);
          return;
        }
      }
    }
  } catch (err) {
    console.log(err);
  }
}

// // Handle sign out request and refresh page
async function onSignOutClick() {
  result = await msalInstance.logoutPopup({
    account: user,
  });
  location.reload();
}

window.onload = function () {
  onSignInClick();
};

/**
 * Retrieve tokenExchangeResource from OAuth card provided by the bot
 * This tokenExchangeResource will later be used to request an accessToken with the right scope.
 */
function getOAuthCardResourceUri(activity) {
  if (
    activity &&
    activity.attachments &&
    activity.attachments[0] &&
    activity.attachments[0].contentType ===
      "application/vnd.microsoft.card.oauth" &&
    activity.attachments[0].content.tokenExchangeResource
  ) {
    // asking for token exchange with AAD
    return activity.attachments[0].content.tokenExchangeResource.uri;
  }
}

/**
 * Retrieve a new access token from the user for the PVA scope based on the tokenExchangeResource
 */
async function exchangeTokenAsync(resourceUri) {
  let user = msalInstance.getAllAccounts();

  if (user.length <= 0) {
    return null;
  }

  const tokenRequest = {
    scopes: [resourceUri],
  };

  try {
    const tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
    return tokenResponse.accessToken;
  } catch (err) {
    console.log(err);
    return null;
  }

  return null;
}

/**
 * Helper function to fetch a JSON API
 */
async function fetchJSON(url, options = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      ...options.headers,
      accept: "application/json",
    },
  });

  if (!res.ok) {
    throw new Error(`Failed to fetch JSON due to ${res.status}`);
  }

  return await res.json();
}

async function renderChatWidget() {
  const userID =
    user?.localAccountId != null
      ? user.localAccountId.substr(0, 36)
      : (Math.random().toString() + Date.now().toString()).substr(0, 64);

  const { token } = await fetchJSON(tokenEndpoint);
  const directLine = window.WebChat.createDirectLine({ token });

  const store = WebChat.createStore(
    {},
    ({ dispatch }) =>
      (next) =>
      (action) => {
        const { type } = action;

        // Block user input until token exchange completes
        if (type === "DIRECT_LINE/CONNECT_FULFILLED") {
          dispatch({ type: "WEB_CHAT/DISABLE_SEND_BOX" });

          dispatch({
            meta: { method: "keyboard" },
            payload: {
              activity: {
                channelData: { postBack: true },
                from: {
                  id: userID,
                  name: user.name,
                  role: "user",
                },
                name: "startConversation",
                type: "event",
              },
            },
            type: "DIRECT_LINE/POST_ACTIVITY",
          });
        }

        if (type === "DIRECT_LINE/INCOMING_ACTIVITY") {
          const activity = action.payload.activity;
          let resourceUri;

          if (
            activity.from &&
            activity.from.role === "bot" &&
            (resourceUri = getOAuthCardResourceUri(activity))
          ) {
            exchangeTokenAsync(resourceUri).then((token) => {
              if (token) {
                directLine
                  .postActivity({
                    type: "invoke",
                    name: "signin/tokenExchange",
                    value: {
                      id: activity.attachments[0].content.tokenExchangeResource
                        .id,
                      connectionName:
                        activity.attachments[0].content.connectionName,
                      token,
                    },
                    from: {
                      id: userID,
                      name: user.name,
                      role: "user",
                    },
                  })
                  .subscribe(
                    (id) => {
                      if (id === "retry") {
                        // Failed to process token, show OAuth card
                        dispatch({ type: "WEB_CHAT/ENABLE_SEND_BOX" });
                        return next(action);
                      }

                      // ✅ Token exchange succeeded — allow user input
                      dispatch({ type: "WEB_CHAT/ENABLE_SEND_BOX" });
                    },
                    (error) => {
                      // ❌ Error during token exchange — optionally log or show fallback
                      dispatch({ type: "WEB_CHAT/ENABLE_SEND_BOX" });
                      return next(action);
                    }
                  );
              } else {
                dispatch({ type: "WEB_CHAT/ENABLE_SEND_BOX" });
                return next(action);
              }
            });

            return; // prevent further processing until token handled
          }
        }

        return next(action);
      }
  );

  const styleOptions = {
    hideUploadButton: true,
    sendBoxTextWrap: true,
    backgroundColor: "#f2f3f5",
    bubbleFromUserBackground: "#133b6d",
    bubbleFromUserTextColor: "#ffffff",
    bubbleFromUserBorderRadius: 18,
    bubbleBackground: "#e9eff7",
    bubbleTextColor: "#000000",
    bubbleBorderRadius: 18,
    botAvatarInitials: "",
    botAvatarImage: "ser_Insight365_BotImage",
    botAvatarBackgroundColor: "#021838",
    userAvatarInitials: "",
    userAvatarBackgroundColor: "#021838",
    userAvatarImage: "ser_Insight365_UserImage",
    fontSize: "15px",
    fontFamily: `'Segoe UI', sans-serif`,
    sendBoxBackground: "#ffffff",
    sendBoxTextColor: "#000",
    sendBoxButtonColor: "#021838",
    sendBoxButtonColorOnFocus: "#0057b8",
  };

  window.WebChat.renderWebChat(
    {
      directLine,
      store,
      userID,
      styleOptions,
    },
    document.getElementById("webchat")
  );
}
