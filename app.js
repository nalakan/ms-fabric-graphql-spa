// This is a test comment to trigger GitHub Pages build
const msalConfig = {
    auth: {
        clientId: "4c072a54-b964-4f2e-a8cf-d571df4c58aa",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "http://localhost:5500"
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
const runQueryBtn = document.getElementById("runQueryBtn");
const queryInput = document.getElementById("queryInput");
const output = document.getElementById("output");

function setOutput(message) {
    output.textContent = message;
}

function showMainContent() {
    loginScreen.style.display = "none";
    mainContent.style.display = "block";
    logoutBtn.style.display = "block";
}

function showLoginScreen() {
    loginScreen.style.display = "block";
    mainContent.style.display = "none";
    logoutBtn.style.display = "none";
}

async function callApi(accessToken, query) {
    setOutput("Loading data...");
    const endpoint = 'https://bb4b4fcd2a8943f0b63391db3f3c4f9e.zbb.graphql.fabric.microsoft.com/v1/workspaces/bb4b4fcd-2a89-43f0-b633-91db3f3c4f9e/graphqlapis/69ea77b8-daf1-45b5-9200-69e4826a1a5a/graphql';

    try {
        const response = await fetch(endpoint, {
            method: "POST",
            headers: {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ query: query, variables: {} })
        });

        const data = await response.json();
        setOutput(JSON.stringify(data, null, 4));
    } catch (error) {
        setOutput("Error loading data: " + error);
    }
}

async function runQuery() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const tokenResponse = await msalInstance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });
            const query = queryInput.value;
            await callApi(tokenResponse.accessToken, query);
        } catch (error) {
            msalInstance.loginRedirect(loginRequest);
        }
    }
}

loginBtn.addEventListener("click", () => {
    msalInstance.loginRedirect(loginRequest);
});

logoutBtn.addEventListener("click", () => {
    msalInstance.logoutRedirect({
        postLogoutRedirectUri: "http://localhost:5500"
    });
});

runQueryBtn.addEventListener("click", runQuery);

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