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
const tableViewBtn = document.getElementById("tableViewBtn");
const viewToggler = document.getElementById("viewToggler");
const playgroundContainer = document.getElementById("graphql-playground");
const voyagerContainer = document.getElementById("voyager-container");
const tableContainer = document.getElementById("table-view");

const jsonResponseInput = document.getElementById("jsonResponseInput"); // MODIFIED: JSON response textarea
const parseAndShowTableBtn = document.getElementById("parseAndShowTableBtn"); // MODIFIED: Parse button
const tableContentDiv = document.getElementById("tableContent"); // Div to render table into

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
    // Default to playground view on login
    playgroundBtn.click();
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
            GraphQLVoyager.init(voyagerContainer, { introspection: introspectionProvider, ...options });
            voyagerInitialized = true;
        } else {
            GraphQLVoyager.init(voyagerContainer, { introspection: introspectionProvider, ...options });
        }
    } catch (error) {
        console.error("Failed to initialize Voyager:", error);
        showErrorMessage("Could not initialize Voyager.");
    }
}

// --- View Toggler Event Listeners ---
playgroundBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "block";
    voyagerContainer.style.display = "none";
    tableContainer.style.display = "none";
    playgroundBtn.classList.add("active");
    voyagerBtn.classList.remove("active");
    tableViewBtn.classList.remove("active");
    contentWrapper.classList.remove('voyager-active', 'table-active');
});

voyagerBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "none";
    voyagerContainer.style.display = "block";
    tableContainer.style.display = "none";
    voyagerBtn.classList.add("active");
    playgroundBtn.classList.remove("active");
    tableViewBtn.classList.remove("active");
    contentWrapper.classList.remove('table-active');
    contentWrapper.classList.add('voyager-active');

    const voyagerOptions = {
        skipRelay: skipRelayCheckbox.checked,
        skipDeprecated: skipDeprecatedCheckbox.checked,
    };
    initializeVoyager(voyagerOptions);
});

tableViewBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "none";
    voyagerContainer.style.display = "none";
    tableContainer.style.display = "block";
    tableViewBtn.classList.add("active");
    playgroundBtn.classList.remove("active");
    voyagerBtn.classList.remove("active");
    contentWrapper.classList.remove('voyager-active');
    contentWrapper.classList.add('table-active');

    // Clear previous table content when switching to this view
    tableContentDiv.innerHTML = '<p class="text-muted">Execute a query in the \'Playground\' tab, copy the JSON response, and paste it above to see data in a table.</p>';
});

// NEW: Event listener for parsing the user-pasted JSON response
parseAndShowTableBtn.addEventListener("click", parseAndRenderTableData);


// --- Voyager Controls ---
zoomSlider.addEventListener("input", (event) => {
    console.log("Voyager Zoom Level:", event.target.value);
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

// --- NEW: Table Data Parsing and Rendering ---

function parseAndRenderTableData() {
    showLoading();
    hideErrorMessage();
    tableContentDiv.innerHTML = ''; // Clear previous table data

    const jsonString = jsonResponseInput.value.trim();

    if (!jsonString) {
        showErrorMessage("Please paste a JSON response in the text area.");
        hideLoading();
        return;
    }

    try {
        const parsedResponse = JSON.parse(jsonString);

        if (parsedResponse.errors) {
            console.error("GraphQL Errors in pasted response:", parsedResponse.errors);
            showErrorMessage(`GraphQL Errors in pasted response: ${parsedResponse.errors.map(e => e.message).join(", ")}`);
            tableContentDiv.innerHTML = `<p class="text-danger">Pasted JSON contains GraphQL errors.</p>`;
            return;
        }

        if (!parsedResponse.data) {
            showErrorMessage("Pasted JSON does not contain a 'data' field. Please ensure it's a valid GraphQL response.");
            tableContentDiv.innerHTML = `<p class="text-danger">Invalid GraphQL response structure.</p>`;
            return;
        }

        // --- Logic to extract data from the 'data' object ---
        let dataToRender = null;
        // Assuming the top-level field in 'data' typically holds the list (e.g., fact_sales, dimension_customers)
        // We'll try to find the first array under 'data' and its 'items' if it's a connection type.
        for (const key in parsedResponse.data) {
            if (parsedResponse.data.hasOwnProperty(key)) {
                const potentialConnection = parsedResponse.data[key];
                if (potentialConnection && Array.isArray(potentialConnection.items)) {
                    dataToRender = potentialConnection.items;
                    break; // Found the data, stop searching
                } else if (Array.isArray(potentialConnection)) { // If it's a direct array (e.g., non-paginated list)
                    dataToRender = potentialConnection;
                    break;
                } else if (typeof potentialConnection === 'object' && potentialConnection !== null) {
                    // Handle cases where 'data' might directly return a single object, e.g., for a single item query
                    // For tabular display, we primarily expect arrays of objects.
                    // If it's a single object, we might wrap it in an array or skip.
                    // For now, we'll just handle connections and direct arrays.
                }
            }
        }

        if (!dataToRender || dataToRender.length === 0) {
            showErrorMessage("Could not find tabular data (e.g., an 'items' array within a connection, or a direct array) in the pasted JSON. Please check your query results.");
            tableContentDiv.innerHTML = `<p class="text-info">No tabular data found in the pasted JSON response.</p>`;
            return;
        }

        renderTable(dataToRender);

    } catch (error) {
        console.error("Error parsing JSON or rendering table:", error);
        showErrorMessage(`Error parsing JSON: ${error.message}. Please ensure it's valid JSON.`);
        tableContentDiv.innerHTML = `<p class="text-danger">An error occurred while parsing the JSON.</p>`;
    } finally {
        hideLoading();
    }
}

// Function to recursively get all headers (keys) from nested objects
function getHeaders(data, prefix = '') {
    const headers = new Set();
    if (!data || data.length === 0) return [];

    data.forEach(item => {
        for (const key in item) {
            if (item.hasOwnProperty(key)) {
                if (typeof item[key] === 'object' && item[key] !== null && !Array.isArray(item[key])) {
                    // Recursively get headers for nested objects
                    getHeaders([item[key]], `${prefix}${key}.`).forEach(nestedHeader => headers.add(nestedHeader));
                } else if (Array.isArray(item[key])) {
                    // For arrays, we might need to decide how to handle them.
                    // For simplicity in a basic table, we'll usually ignore complex arrays
                    // or just represent them as "Array" or "..."
                    headers.add(`${prefix}${key}`); // Just add the array field name
                }
                else {
                    headers.add(`${prefix}${key}`);
                }
            }
        }
    });
    return Array.from(headers);
}

// Function to recursively get a value for a header path
function getNestedValue(obj, path) {
    const parts = path.split('.');
    let current = obj;
    for (let i = 0; i < parts.length; i++) {
        if (current === null || typeof current !== 'object' || !current.hasOwnProperty(parts[i])) {
            return ''; // Return empty string if path doesn't exist
        }
        current = current[parts[i]];
    }
    // Special formatting for DateTimes if they are common
    if (typeof current === 'string' && current.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/)) {
        return new Date(current).toLocaleDateString() + ' ' + new Date(current).toLocaleTimeString();
    }
    // Handle null/undefined display
    if (current === null || current === undefined) {
        return '';
    }
    // Handle numbers with fixed decimals (e.g., currency)
    if (typeof current === 'number' && (path.includes('Price') || path.includes('Tax') || path.includes('Profit'))) {
        return current.toFixed(2);
    }
    return current.toString(); // Convert other values to string
}


function renderTable(data) {
    if (!data || data.length === 0) {
        tableContentDiv.innerHTML = '<p class="text-info">No data items found in the response to render a table.</p>';
        return;
    }

    // Get dynamic headers from the data structure
    const headers = getHeaders(data);
    if (headers.length === 0) {
        tableContentDiv.innerHTML = '<p class="text-warning">Could not determine table headers from the data structure.</p>';
        return;
    }

    let tableHTML = `<table class="table table-striped table-hover data-table"><thead><tr>`;
    headers.forEach(header => {
        // Make headers more readable (e.g., "dimension_customer.Customer" -> "Customer")
        const displayHeader = header.includes('.') ? header.split('.').pop().replace(/([A-Z])/g, ' $1').trim() : header.replace(/([A-Z])/g, ' $1').trim();
        tableHTML += `<th>${displayHeader}</th>`;
    });
    tableHTML += `</tr></thead><tbody>`;

    // Populate table rows
    data.forEach(item => {
        tableHTML += `<tr>`;
        headers.forEach(headerPath => {
            tableHTML += `<td>${getNestedValue(item, headerPath)}</td>`;
        });
        tableHTML += `</tr>`;
    });

    tableHTML += `</tbody></table>`;
    tableContentDiv.innerHTML = tableHTML;
}


// Initial Load
showLoading();
msalInstance.handleRedirectPromise().catch(err => console.error(err));