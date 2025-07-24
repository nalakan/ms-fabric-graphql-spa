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
const tableResultContainer = document.getElementById('table-result-container');
const generateTableBtn = document.getElementById('generate-table-btn');
const tableContainer = document.getElementById('table-container');
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
    tableResultContainer.style.display = "flex"; // Default to showing table with playground
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

function injectShowInTableButton() {
    const topBar = playgroundContainer.querySelector('.topBar');
    if (topBar) {
        const button = document.createElement('button');
        button.innerText = 'Show in Table';
        button.className = 'execute-button'; // Use playground's button style
        button.style.marginLeft = '10px';
        button.onclick = () => {
            try {
                const resultViewer = playgroundContainer.querySelector('.result-window .CodeMirror-code');
                if (resultViewer) {
                    const lines = Array.from(resultViewer.querySelectorAll('.CodeMirror-line')).map(line => line.innerText).join('\n');
                    // Remove leading line numbers if they exist
                    // Remove leading line numbers and any non-JSON prefix/suffix
                    const jsonMatch = lines.match(/\{.*\}/s);
                    const cleanJson = jsonMatch ? jsonMatch[0] : lines.replace(/^\d+\s/gm, '').trim();
                    const jsonData = JSON.parse(cleanJson);
                    const dataArray = findDataArray(jsonData.data);
                    if (dataArray) {
                        createTable(dataArray);
                        tableResultContainer.style.display = 'flex';
                    } else {
                        tableContainer.innerHTML = '<p>Could not find an array of objects to display in the result.</p>';
                    }
                }
            } catch (e) {
                console.error('Error processing playground result:', e);
                tableContainer.innerHTML = '<p class="text-danger">Could not parse JSON from the result pane.</p>';
            }
        };
        topBar.appendChild(button);
    }
}

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

    // Inject the button after a short delay to ensure the playground is rendered
    setTimeout(() => {
        console.log("Attempting to inject 'Show in Table' button...");
        injectShowInTableButton();
    }, 5000);
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

async function initializeVoyager() {
    try {
        GraphQLVoyager.init(voyagerContainer, { introspection: await introspectionProvider });
    } catch (error) {
        console.error("Failed to initialize Voyager:", error);
        showErrorMessage("Could not initialize Voyager.");
    }
}

playgroundBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "block";
    voyagerContainer.style.display = "none";
    tableResultContainer.style.display = "flex"; // Show table view
    playgroundBtn.classList.add("active");
    voyagerBtn.classList.remove("active");
    contentWrapper.classList.remove('voyager-active');
});

voyagerBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "none";
    voyagerContainer.style.display = "block";
    tableResultContainer.style.display = "none"; // Hide table view
    voyagerBtn.classList.add("active");
    playgroundBtn.classList.remove("active");
    contentWrapper.classList.add('voyager-active');
    initializeVoyager();
});

// --- Tabular Result Logic (Clipboard Method) ---

function findDataArray(obj) {
    const queue = [obj];
    while (queue.length > 0) {
        const current = queue.shift();
        if (Array.isArray(current) && current.length > 0 && typeof current[0] === 'object' && current[0] !== null) {
            return current;
        }
        if (current && typeof current === 'object') {
            Object.values(current).forEach(value => queue.push(value));
        }
    }
    return null;
}

function createTable(dataArray) {
    tableContainer.innerHTML = '';
    const table = document.createElement('table');
    table.className = 'table table-bordered table-striped';
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    const headers = Object.keys(dataArray[0]);
    headers.forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    dataArray.forEach(rowData => {
        const row = document.createElement('tr');
        headers.forEach(header => {
            const cell = document.createElement('td');
            const cellValue = rowData[header];
            if (typeof cellValue === 'object' && cellValue !== null) {
                const pre = document.createElement('pre');
                pre.textContent = JSON.stringify(cellValue, null, 2);
                cell.appendChild(pre);
            } else {
                cell.textContent = cellValue;
            }
            row.appendChild(cell);
        });
        tbody.appendChild(row);
    });
    table.appendChild(tbody);
    tableContainer.appendChild(table);
}

generateTableBtn.addEventListener('click', async () => {
    try {
        const clipboardText = await navigator.clipboard.readText();
        if (!clipboardText) {
            tableContainer.innerHTML = '<p>Clipboard is empty.</p>';
            tableResultContainer.style.display = 'flex';
            return;
        }

        const jsonData = JSON.parse(clipboardText);
        const dataArray = findDataArray(jsonData);

        if (dataArray) {
            createTable(dataArray);
        } else {
            tableContainer.innerHTML = '<p>Could not find an array of objects to display in the clipboard text.</p>';
        }
        tableResultContainer.style.display = 'flex';

    } catch (err) {
        console.error('Failed to generate table:', err);
        let errorMessage = 'An error occurred.';
        if (err.name === 'NotAllowedError') {
            errorMessage = 'Permission to read clipboard was denied. Please allow access in your browser settings or manually paste the JSON into the playground and use the "Show in Table" button.';
        } else if (err instanceof SyntaxError) {
            errorMessage = 'The text on the clipboard is not valid JSON.';
        } else if (err.message.includes('readText')) {
            errorMessage = 'Could not read from clipboard. This might be due to browser security restrictions (e.g., not running over HTTPS or missing user gesture). Try manually pasting the JSON into the playground and using the "Show in Table" button.';
        }
        tableContainer.innerHTML = `<p class="text-danger">${errorMessage}</p>`;
        tableResultContainer.style.display = 'flex';
    }
});

// Initial Load
showLoading();
msalInstance.handleRedirectPromise().catch(err => console.error(err));