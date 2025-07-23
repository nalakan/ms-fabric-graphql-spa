const msalConfig = {
    auth: {
        clientId: "4c072a54-b964-4f2e-a8cf-d571df4c58aa",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://graphql.bidiaries.com/" // Updated for custom domain
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
    scopes: [
        "https://analysis.windows.net/powerbi/api/Dataset.Read.All",
        "https://analysis.windows.net/powerbi/api/Workspace.Read.All",
        "https://analysis.windows.net/powerbi/api/Capacity.Read.All"
    ]
};

const loginScreen = document.getElementById("loginScreen");
const mainContent = document.getElementById("mainContent");
const loginBtn = document.getElementById("loginBtn");
const logoutBtn = document.getElementById("logoutBtn");

const GRAPHQL_ENDPOINT = 'https://bb4b4fcd2a8943f0b63391db3f3c4f9e.zbb.graphql.fabric.microsoft.com/v1/workspaces/bb4b4fcd-2a89-43f0-b633-91db3f3c4f9e/graphqlapis/69ea77b8-daf1-45b5-9200-69e4826a1a5a/graphql';

function showMainContent() {
    loginScreen.style.display = "none";
    mainContent.style.display = "block";
    logoutBtn.style.display = "block";
    initializeGraphQLPlayground();
}

function showLoginScreen() {
    loginScreen.style.display = "block";
    mainContent.style.display = "none";
    logoutBtn.style.display = "none";
}

async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
        return null;
    }
    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        });
        return tokenResponse.accessToken;
    } catch (error) {
        console.error("Silent token acquisition failed, acquiring token interactively:", error);
        msalInstance.loginRedirect(loginRequest);
        return null; // Will redirect, so no token immediately
    }
}

async function initializeGraphQLPlayground() {
    const accessToken = await getAccessToken();
    if (!accessToken) {
        console.error("No access token available for GraphQL Playground.");
        return;
    }

    // Ensure the GraphQLPlayground object is available
    if (typeof GraphQLPlayground === 'undefined' || !GraphQLPlayground.init) {
        console.error("GraphQLPlayground not loaded. Retrying...");
        // Simple retry mechanism, or more robust loading check
        setTimeout(initializeGraphQLPlayground, 500);
        return;
    }

    GraphQLPlayground.init(document.getElementById('graphql-playground'), {
        endpoint: GRAPHQL_ENDPOINT,
        settings: {
            'request.credentials': 'omit',
            'editor.theme': 'dark',
            'editor.reuseHeaders': true,
            'editor.fontFamily': `'Source Code Pro', 'Consolas', 'Inconsolata', 'Droid Sans Mono', 'Monaco', monospace`,
            'editor.fontSize': 14,
            'tracing.hideTracingByDefault': true,
            'queryPlan.hideQueryPlanByDefault': true,
        },
        headers: {
            'Authorization': `Bearer ${accessToken}`,
        },
        tab: {
            endpoint: GRAPHQL_ENDPOINT,
            query: `query {
  trips(first: 5) {
    items {
      PaymentType
      medallion(first: 1) {
        items {
          MedallionCode
        }
      }
    }
  }
}`,
            variables: `{}`,
        },
    });
}

loginBtn.addEventListener("click", () => {
    msalInstance.loginRedirect(loginRequest);
});

logoutBtn.addEventListener("click", () => {
    msalInstance.logoutRedirect({
        postLogoutRedirectUri: msalConfig.auth.redirectUri
    });
});

// Handle redirect callback
msalInstance.handleRedirectPromise().then(response => {
    if (response && response.accessToken) {
        showMainContent();
    } else {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            showMainContent();
        } else {
            showLoginScreen();
        }
    }
}).catch(error => {
    console.error(error);
    showLoginScreen();
});
