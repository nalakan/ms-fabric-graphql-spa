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
const schemaDiagramBtn = document.getElementById("schemaDiagramBtn");
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
    console.log("Showing main content.");
    hideLoading();
    hideErrorMessage();
    loginScreen.style.display = "none";
    mainContent.style.display = "block";
    logoutBtn.style.display = "block";
    themeToggleBtn.style.display = "block"; // Show theme toggle
    schemaDiagramBtn.style.display = "block"; // Show schema diagram button
    initializeGraphQLPlayground();
}

function showLoginScreen() {
    console.log("Showing login screen.");
    hideLoading();
    hideErrorMessage();
    loginScreen.style.display = "flex"; // Use flex for centering
    mainContent.style.display = "none";
    logoutBtn.style.display = "none";
    themeToggleBtn.style.display = "none"; // Hide theme toggle
    schemaDiagramBtn.style.display = "none"; // Hide schema diagram button
}

async function getAccessToken() {
    console.log("Attempting to get access token.");
    hideErrorMessage();
    showLoading();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
        console.log("No accounts found, cannot get silent token.");
        hideLoading();
        return null;
    }
    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        });
        console.log("Access token acquired silently.", tokenResponse);
        hideLoading();
        return tokenResponse.accessToken;
    } catch (error) {
        hideLoading();
        console.error("Silent token acquisition failed:", error);
        showErrorMessage("Authentication failed. Please try logging in again.");
        msalInstance.loginRedirect(loginRequest);
        return null; // Will redirect, so no token immediately
    }
}

async function initializeGraphQLPlayground() {
    console.log("Initializing GraphQL Playground.");
    const accessToken = await getAccessToken();
    if (!accessToken) {
        console.error("No access token available for GraphQL Playground. Login required.");
        showErrorMessage("Could not get access token for GraphQL API. Please log in.");
        return;
    }

    if (typeof GraphQLPlayground === 'undefined' || !GraphQLPlayground.init) {
        console.error("GraphQLPlayground not loaded. Retrying...");
        setTimeout(initializeGraphQLPlayground, 500);
        return;
    }

    const initialTheme = document.body.classList.contains('dark-theme') ? 'dark' : 'light';

    GraphQLPlayground.init(document.getElementById('graphql-playground'), {
        endpoint: GRAPHQL_ENDPOINT,
        settings: {
            'request.credentials': 'omit',
            'editor.theme': initialTheme,
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
    console.log("GraphQL Playground initialized.");
}

// Theme Toggle Logic
function toggleTheme() {
    console.log("Theme toggle button clicked.");
    document.body.classList.toggle('dark-theme');
    const currentTheme = document.body.classList.contains('dark-theme') ? 'dark' : 'light';
    if (typeof GraphQLPlayground !== 'undefined' && GraphQLPlayground.setSettings) {
        GraphQLPlayground.setSettings({
            'editor.theme': currentTheme
        });
    }
}

// Event Listeners
loginBtn.addEventListener("click", () => {
    console.log("Login button clicked.");
    hideErrorMessage();
    showLoading();
    msalInstance.loginRedirect(loginRequest);
});

logoutBtn.addEventListener("click", () => {
    console.log("Logout button clicked.");
    hideErrorMessage();
    showLoading();
    msalInstance.logoutRedirect({
        postLogoutRedirectUri: msalConfig.auth.redirectUri
    });
});

themeToggleBtn.addEventListener("click", toggleTheme);

schemaDiagramBtn.addEventListener("click", () => {
    console.log("Schema Diagram button clicked.");
    const voyagerUrl = `https://graphql-voyager.com/?${encodeURIComponent(GRAPHQL_ENDPOINT)}`;
    window.open(voyagerUrl, '_blank');
});

// Main application logic on page load
console.log("Page loaded. Handling redirect promise...");
showLoading();
msalInstance.handleRedirectPromise().then(response => {
    if (response && response.accessToken) {
        console.log("Redirect promise resolved with access token.", response);
        showMainContent();
    } else {
        console.log("Redirect promise resolved without access token. Checking for existing accounts...");
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            console.log("Existing accounts found. Showing main content.");
            showMainContent();
        } else {
            console.log("No existing accounts. Showing login screen.");
            showLoginScreen();
        }
    }
}).catch(error => {
    hideLoading();
    console.error("Error handling redirect promise:", error);
    showErrorMessage("An error occurred during authentication. Please try again.");
    showLoginScreen();
});
