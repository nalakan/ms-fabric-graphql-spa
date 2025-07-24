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

// DOM Elements
const loginScreen = document.getElementById("loginScreen");
const mainContent = document.getElementById("mainContent");
const loginBtn = document.getElementById("loginBtn");
const logoutBtn = document.getElementById("logoutBtn");
const loadingOverlay = document.getElementById("loadingOverlay");
const errorMessageDiv = document.getElementById("errorMessage");
const playgroundBtn = document.getElementById("playgroundBtn");
const voyagerBtn = document.getElementById("voyagerBtn");
const viewToggler = document.getElementById("viewToggler");
const playgroundContainer = document.getElementById("graphql-playground");
const voyagerContainer = document.getElementById("voyager-container");
const zoomSlider = document.getElementById("zoom-slider");
const skipRelayCheckbox = document.getElementById("skip-relay-checkbox");
const skipDeprecatedCheckbox = document.getElementById("skip-deprecated-checkbox");

const contentWrapper = document.querySelector('.content-wrapper');

const GRAPHQL_ENDPOINT = 'https://bb4b4fcd2a8943f0b63391db3f3c4f9e.zbb.graphql.fabric.microsoft.com/v1/workspaces/bb4b4fcd-2a89-43f0-b633-91db3f3c4f9e/graphqlapis/69ea77b8-daf1-45b5-9200-69e4826a1a5a/graphql';

// --- Core UI Functions ---

function showLoading() { loadingOverlay.classList.add("show"); }
function hideLoading() { loadingOverlay.classList.remove("show"); }

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
    mainContent.style.display = "flex";
    logoutBtn.style.display = "block";
    viewToggler.style.display = "block";
    
    initializeGraphQLPlayground();
}

function showLoginScreen() {
    hideLoading();
    hideErrorMessage();
    loginScreen.style.display = "block";
    mainContent.style.display = "none";
    logoutBtn.style.display = "none";
    viewToggler.style.display = "none";
}

// --- MSAL Authentication ---

async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return null;
    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
        return tokenResponse.accessToken;
    } catch (error) {
        console.error("Silent token acquisition failed, acquiring token interactively:", error);
        msalInstance.loginRedirect(loginRequest);
        return null;
    }
}

loginBtn.addEventListener("click", () => {
    hideErrorMessage();
    showLoading();
    msalInstance.loginRedirect(loginRequest);
});

logoutBtn.addEventListener("click", () => {
    msalInstance.logoutRedirect({ postLogoutRedirectUri: msalConfig.auth.redirectUri });
});

msalInstance.handleRedirectPromise().then(response => {
    if (response && response.accessToken) {
        showMainContent();
    } else if (msalInstance.getAllAccounts().length > 0) {
        showMainContent();
    } else {
        showLoginScreen();
    }
}).catch(error => {
    console.error(error);
    showErrorMessage("An error occurred during authentication. Please try again.");
    showLoginScreen();
});

// --- GraphQL Tools Initialization ---



async function initializeGraphQLPlayground() {
    const accessToken = await getAccessToken();
    if (!accessToken) {
        showErrorMessage("Could not get access token. Please log in.");
        return;
    }
    GraphQLPlayground.init(playgroundContainer, {
        endpoint: GRAPHQL_ENDPOINT,
        settings: { 'editor.theme': 'dark', 'editor.reuseHeaders': true },
        headers: { 'Authorization': `Bearer ${accessToken}` },
    });

    
}

async function introspectionProvider(query) {
    const accessToken = await getAccessToken();
    if (!accessToken) throw new Error("No access token for introspection.");
    const response = await fetch(GRAPHQL_ENDPOINT, {
        method: 'post',
        headers: { 'Accept': 'application/json', 'Content-Type': 'application/json', 'Authorization': `Bearer ${accessToken}` },
        body: JSON.stringify({ query: query }),
    });
    return response.json();
}

let voyagerInitialized = false;

async function initializeVoyager(options = {}) {
    try {
        if (!voyagerInitialized) {
            GraphQLVoyager.init(voyagerContainer, { introspection: await introspectionProvider, ...options });
            voyagerInitialized = true;
        } else {
            // Assuming there's a way to update options after initialization
            // This might require re-initializing or calling a specific update method
            // For now, we'll re-initialize if options change, which might not be ideal for performance
            // A better approach would be to find a Voyager API for updating settings.
            GraphQLVoyager.init(voyagerContainer, { introspection: await introspectionProvider, ...options });
        }
    } catch (error) {
        console.error("Failed to initialize Voyager:", error);
        showErrorMessage("Could not initialize Voyager.");
    }
}

playgroundBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "block";
    voyagerContainer.style.display = "none";
    playgroundBtn.classList.add("active");
    voyagerBtn.classList.remove("active");
    contentWrapper.classList.remove('voyager-active');
});

voyagerBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "none";
    voyagerContainer.style.display = "block";
    voyagerBtn.classList.add("active");
    playgroundBtn.classList.remove("active");
    contentWrapper.classList.add('voyager-active');
    
    const voyagerOptions = {
        skipRelay: skipRelayCheckbox.checked,
        skipDeprecated: skipDeprecatedCheckbox.checked,
        // Add zoom level if Voyager supports it directly in init, otherwise handle separately
    };
    initializeVoyager(voyagerOptions);
});

zoomSlider.addEventListener("input", (event) => {
    // This assumes Voyager has a direct way to set zoom or we need to apply CSS transform
    // For now, we'll just log the value. Actual implementation depends on Voyager API.
    console.log("Voyager Zoom Level:", event.target.value);
    // If Voyager doesn't have a direct zoom API, we might need to apply CSS transform
    // voyagerContainer.style.transform = `scale(${event.target.value})`;
    // voyagerContainer.style.transformOrigin = `top left`;
});

skipRelayCheckbox.addEventListener("change", () => {
    initializeVoyager({
        skipRelay: skipRelayCheckbox.checked,
        skipDeprecated: skipDeprecatedCheckbox.checked,
    });
});

skipDeprecatedCheckbox.addEventListener("change", () => {
    initializeVoyager({
        skipRelay: skipRelayCheckbox.checked,
        skipDeprecated: skipDeprecatedCheckbox.checked,
    });
});



// Initial Load
showLoading();
msalInstance.handleRedirectPromise().catch(err => console.error(err));