const msalConfig = {
    auth: {
        clientId: "4c072a54-b964-4f2e-a8cf-d571df4c58aa",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://nalakan.github.io/-fabric-graphql-spa/"
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
const themeToggleBtn = document.getElementById("themeToggleBtn");
const loadingOverlay = document.getElementById("loadingOverlay");
const errorMessageDiv = document.getElementById("errorMessage");

const GRAPHQL_ENDPOINT = 'https://bb4b4fcd2a8943f0b63391db3f3c4f9e.zbb.graphql.fabric.microsoft.com/v1/workspaces/bb4b4fcd-2a89-43f0-b633-91db3f3c4f9e/graphqlapis/69ea77b8-daf1-45b5-9200-69e4826a1a5a/graphql';

function showLoading() {
    loadingOverlay.classList.add("show");
}

function hideLoading() {
    loadingOverlay.classList.remove("show");
}

function showErrorMessage(message) {
    errorMessageDiv.textContent = message;
    errorMessageDiv.classList.remove("d-none");
}

function hideErrorMessage() {
    errorMessageDiv.classList.add("d-none");
    errorMessageDiv.textContent = "";
}

function showMainContent() {
    hideLoading();
    hideErrorMessage();
    loginScreen.style.display = "none";
    mainContent.style.display = "block";
    logoutBtn.style.display = "block";
    themeToggleBtn.style.display = "block"; // Show theme toggle
    initializeGraphQLPlayground();
}

function showLoginScreen() {
    hideLoading();
    hideErrorMessage();
    loginScreen.style.display = "flex"; // Use flex for centering
    mainContent.style.display = "none";
    logoutBtn.style.display = "none";
    themeToggleBtn.style.display = "none"; // Hide theme toggle
}

async function getAccessToken() {
    hideErrorMessage();
    showLoading();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
        hideLoading();
        return null;
    }
    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        });
        hideLoading();
        return tokenResponse.accessToken;
    } catch (error) {
        hideLoading();
        console.error("Silent token acquisition failed, acquiring token interactively:", error);
        showErrorMessage("Authentication failed. Please try logging in again.");
        msalInstance.loginRedirect(loginRequest);
        return null; // Will redirect, so no token immediately
    }
}

async function initializeGraphQLPlayground() {
    const accessToken = await getAccessToken();
    if (!accessToken) {
        console.error("No access token available for GraphQL Playground.");
        showErrorMessage("Could not get access token for GraphQL API. Please log in.");
        return;
    }

    if (typeof GraphQLPlayground === 'undefined' || !GraphQLPlayground.init) {
        console.error("GraphQLPlayground not loaded. Retrying...");
        setTimeout(initializeGraphQLPlayground, 500);
        return;
    }

    // Determine initial theme for Playground based on body class
    const initialTheme = document.body.classList.contains('dark-theme') ? 'dark' : 'light';

    GraphQLPlayground.init(document.getElementById('graphql-playground'), {
        endpoint: GRAPHQL_ENDPOINT,
        settings: {
            'request.credentials': 'omit',
            'editor.theme': initialTheme, // Set Playground theme based on app theme
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

// Theme Toggle Logic
function toggleTheme() {
    document.body.classList.toggle('dark-theme');
    const currentTheme = document.body.classList.contains('dark-theme') ? 'dark' : 'light';
    // Attempt to update Playground theme if it's initialized
    if (typeof GraphQLPlayground !== 'undefined' && GraphQLPlayground.setSettings) {
        GraphQLPlayground.setSettings({
            'editor.theme': currentTheme
        });
    }
}

loginBtn.addEventListener("click", () => {
    hideErrorMessage();
    showLoading();
    msalInstance.loginRedirect(loginRequest);
});

logoutBtn.addEventListener("click", () => {
    hideErrorMessage();
    showLoading();
    msalInstance.logoutRedirect({
        postLogoutRedirectUri: msalConfig.auth.redirectUri
    });
});

themeToggleBtn.addEventListener("click", toggleTheme);

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
    hideLoading();
    console.error(error);
    showErrorMessage("An error occurred during authentication. Please try again.");
    showLoginScreen();
});

// Initial check on page load
showLoading();
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
    hideLoading();
    console.error(error);
    showErrorMessage("An error occurred during initial load. Please try logging in.");
    showLoginScreen();
});
