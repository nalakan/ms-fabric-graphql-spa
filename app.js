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
const loadingOverlay = document.getElementById("loadingOverlay");
const errorMessageDiv = document.getElementById("errorMessage");

const playgroundBtn = document.getElementById("playgroundBtn");
const voyagerBtn = document.getElementById("voyagerBtn");
const viewToggler = document.getElementById("viewToggler");

const playgroundContainer = document.getElementById("graphql-playground");
const voyagerContainer = document.getElementById("voyager-container");
const tableResultContainer = document.getElementById('table-result-container');

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
        return null;
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

    GraphQLPlayground.init(document.getElementById('graphql-playground'), {
        endpoint: GRAPHQL_ENDPOINT,
        settings: {
            'request.credentials': 'omit',
            'editor.theme': 'dark',
            'editor.reuseHeaders': true,
            'editor.fontFamily': `'Source Code Pro', 'Consolas', 'Inconsolata', 'Droid Sans Mono', 'Monaco', monospace`,
            'editor.fontSize': 14,
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

showLoading();

async function introspectionProvider(query) {
    const accessToken = await getAccessToken();
    if (!accessToken) {
        throw new Error("No access token available for introspection.");
    }
    return fetch(GRAPHQL_ENDPOINT, {
        method: 'post',
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${accessToken}`
        },
        body: JSON.stringify({
            query: query
        }),
    }).then(response => response.json());
}

async function initializeVoyager() {
    try {
        const provider = await introspectionProvider;
        GraphQLVoyager.init(voyagerContainer, {
            introspection: provider,
            displayOptions: {
                rootType: "Query",
                sortByAlphabet: true,
                showPanel: true,
            }
        });
    } catch (error) {
        console.error("Failed to initialize Voyager:", error);
        showErrorMessage("Could not initialize the Voyager visualization.");
    }
}

playgroundBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "block";
    voyagerContainer.style.display = "none";
    playgroundBtn.classList.add("active");
    voyagerBtn.classList.remove("active");
});

voyagerBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "none";
    voyagerContainer.style.display = "block";
    voyagerBtn.classList.add("active");
    playgroundBtn.classList.remove("active");
    initializeVoyager();
});

// --- Tabular Result Logic ---

function findDataArray(obj) {
    let result = null;
    let maxDepth = -1;

    function recurse(currentObj, depth) {
        if (!currentObj || typeof currentObj !== 'object') {
            return;
        }

        if (Array.isArray(currentObj)) {
            if (depth > maxDepth) {
                maxDepth = depth;
                result = currentObj;
            }
            return;
        }

        for (const key in currentObj) {
            if (currentObj.hasOwnProperty(key)) {
                recurse(currentObj[key], depth + 1);
            }
        }
    }

    recurse(obj, 0);
    return result;
}

function createTable(data) {
    const tableContainer = document.getElementById('table-container');
    if (!tableContainer || !tableResultContainer) return;

    tableContainer.innerHTML = '';
    const dataArray = findDataArray(data);

    if (!dataArray || dataArray.length === 0 || typeof dataArray[0] !== 'object') {
        playgroundContainer.style.height = '100%';
        tableResultContainer.classList.remove('visible');
        return;
    }

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

    playgroundContainer.style.height = '60%';
    tableResultContainer.classList.add('visible');
}

// --- Fetch Interceptor ---
const originalFetch = window.fetch;
window.fetch = function(...args) {
    const [url] = args;
    const promise = originalFetch.apply(this, args);

    if (url === GRAPHQL_ENDPOINT) {
        promise.then(response => {
            if (response.ok) {
                response.clone().json().then(data => {
                    console.log("GraphQL response intercepted:", data);
                    createTable(data);
                }).catch(err => {
                    console.error('Error parsing JSON for table:', err);
                    playgroundContainer.style.height = '100%';
                    tableResultContainer.classList.remove('visible');
                });
            } else {
                 playgroundContainer.style.height = '100%';
                 tableResultContainer.classList.remove('visible');
            }
        });
    }

    return promise;
};