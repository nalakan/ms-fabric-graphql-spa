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

const jsonResponseInput = document.getElementById("jsonResponseInput");
const parseAndShowTableBtn = document.getElementById("parseAndShowTableBtn");
const tableContentDiv = document.getElementById("tableContent");
const downloadButtonsDiv = document.getElementById("downloadButtons"); // NEW: Download buttons container
const downloadCsvBtn = document.getElementById("downloadCsvBtn");     // NEW: Download CSV button
const downloadExcelBtn = document.getElementById("downloadExcelBtn"); // NEW: Download Excel button

const contentWrapper = document.querySelector('.content-wrapper');

const GRAPHQL_ENDPOINT = 'https://bb4b4fcd2a8943f0b63391db3f3c4f9e.zbb.graphql.fabric.microsoft.com/v1/workspaces/bb4b4fcd-2a89-43f0-b633-91db3f3c4f9e/graphqlapis/69ea77b8-daf1-45b5-9200-69e4826a1a5a/graphql';

// Global variables to store processed data and headers for download
let currentTableData = [];
let currentTableHeaders = []; // Stores the raw header paths like "dimension_customer.Customer"

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
    playgroundBtn.click(); // Default to playground view on login
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
        msalInstance.loginRedirect(loginRequest); // Force interactive login if silent fails
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

// Fixed for Voyager blank page: Ensure introspectionProvider explicitly handles errors
async function introspectionProvider(query) {
    const accessToken = await getAccessToken();
    if (!accessToken) {
        console.error("introspectionProvider: No access token available.");
        // This will be caught by initializeVoyager, but good to log here too
        throw new Error("No access token for introspection.");
    }
    try {
        const response = await fetch(GRAPHQL_ENDPOINT, {
            method: 'post',
            headers: { 'Accept': 'application/json', 'Content-Type': 'application/json', 'Authorization': `Bearer ${accessToken}` },
            body: JSON.stringify({ query: query }),
        });
        if (!response.ok) {
            const errorBody = await response.text();
            console.error(`introspectionProvider: Network response not ok: ${response.status} ${response.statusText}`, errorBody);
            throw new Error(`Failed to fetch schema: ${response.status} ${response.statusText}`);
        }
        const jsonResponse = await response.json();
        if (jsonResponse.errors) {
            console.error("introspectionProvider: GraphQL errors in schema response:", jsonResponse.errors);
            throw new Error(`GraphQL errors fetching schema: ${jsonResponse.errors.map(e => e.message).join(", ")}`);
        }
        return jsonResponse;
    } catch (error) {
        console.error("introspectionProvider: Error during fetch or parsing:", error);
        // Rethrow the error so initializeVoyager can catch and display
        throw error;
    }
}

let voyagerInitialized = false;

async function initializeVoyager(options = {}) {
    try {
        if (!voyagerInitialized) {
            // Ensure voyagerContainer is visible and has dimensions before initializing
            voyagerContainer.style.display = 'block'; // Ensure it's rendered to get dimensions
            GraphQLVoyager.init(voyagerContainer, { introspection: introspectionProvider, ...options });
            voyagerInitialized = true;
        } else {
            // If already initialized, you might need to re-render or update if options change
            // For Voyager, often re-initializing is the simplest path if it doesn't offer update APIs
            GraphQLVoyager.init(voyagerContainer, { introspection: introspectionProvider, ...options });
        }
    } catch (error) {
        console.error("Failed to initialize Voyager:", error);
        showErrorMessage("Could not initialize Voyager. Check console for details.");
    }
}

// --- View Toggler Event Listeners ---
playgroundBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "block";
    voyagerContainer.style.display = "none";
    tableContainer.style.display = "none";
    downloadButtonsDiv.style.display = "none"; // Hide download buttons
    playgroundBtn.classList.add("active");
    voyagerBtn.classList.remove("active");
    tableViewBtn.classList.remove("active");
    contentWrapper.classList.remove('voyager-active', 'table-active');
});

voyagerBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "none";
    voyagerContainer.style.display = "block"; // Make sure it's block for Voyager to draw
    tableContainer.style.display = "none";
    downloadButtonsDiv.style.display = "none"; // Hide download buttons
    voyagerBtn.classList.add("active");
    playgroundBtn.classList.remove("active");
    tableViewBtn.classList.remove("active");
    contentWrapper.classList.remove('table-active');
    contentWrapper.classList.add('voyager-active');

    const voyagerOptions = {
        skipRelay: skipRelayCheckbox.checked,
        skipDeprecated: skipDeprecatedCheckbox.checked,
    };
    initializeVoyager(voyagerOptions); // Initialize/re-initialize Voyager
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

    // Clear previous table content and hide download buttons when switching to this view
    tableContentDiv.innerHTML = '<p class="text-muted">Execute a query in the \'Playground\' tab, copy the JSON response, and paste it above to see data in a table.</p>';
    downloadButtonsDiv.style.display = "none";
    currentTableData = [];
    currentTableHeaders = [];
});

// Event listener for parsing the user-pasted JSON response
parseAndShowTableBtn.addEventListener("click", parseAndRenderTableData);

// NEW: Event listeners for download buttons
downloadCsvBtn.addEventListener("click", () => downloadTableData('csv'));
downloadExcelBtn.addEventListener("click", () => downloadTableData('excel'));


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

// --- Table Data Parsing and Rendering ---

function parseAndRenderTableData() {
    showLoading();
    hideErrorMessage();
    tableContentDiv.innerHTML = ''; // Clear previous table data
    downloadButtonsDiv.style.display = "none"; // Hide download buttons until data is rendered

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
            // Clear global data/headers if errors are present
            currentTableData = [];
            currentTableHeaders = [];
            return;
        }

        if (!parsedResponse.data) {
            showErrorMessage("Pasted JSON does not contain a 'data' field. Please ensure it's a valid GraphQL response.");
            tableContentDiv.innerHTML = `<p class="text-danger">Invalid GraphQL response structure.</p>`;
            currentTableData = [];
            currentTableHeaders = [];
            return;
        }

        let dataToRender = null;
        // Try to find the array of items. Prioritize 'items' within a connection.
        // This makes the table view more robust for various query outputs.
        for (const key in parsedResponse.data) {
            if (parsedResponse.data.hasOwnProperty(key)) {
                const potentialData = parsedResponse.data[key];
                if (potentialData && typeof potentialData === 'object' && potentialData.items && Array.isArray(potentialData.items)) {
                    dataToRender = potentialData.items;
                    break;
                } else if (Array.isArray(potentialData)) { // If it's a direct array (e.g., from a top-level list query without 'items')
                    dataToRender = potentialData;
                    break;
                } else if (typeof potentialData === 'object' && potentialData !== null) {
                    // If it's a single object (e.g., a query for a single item by ID)
                    // We can wrap it in an array to make it tabular, or decide to skip if it's not meant for tabular display
                    // For now, if no array found, and there's one object, we'll try to treat it as a single row.
                    let allAreObjects = true;
                    for (const subKey in potentialData) {
                        if (potentialData.hasOwnProperty(subKey) && ! (typeof potentialData[subKey] === 'object' && potentialData[subKey] !== null)) {
                            allAreObjects = false; // Not a nested object structure, might be a flat object
                            break;
                        }
                    }
                    if (allAreObjects && Object.keys(potentialData).length > 0) { // If it's a single flat object, treat as one row
                         dataToRender = [potentialData];
                         break;
                    }
                }
            }
        }

        if (!dataToRender || dataToRender.length === 0) {
            showErrorMessage("Could not find tabular data (e.g., an 'items' array within a connection, a direct array, or a single top-level object) in the pasted JSON. Please check your query results.");
            tableContentDiv.innerHTML = `<p class="text-info">No tabular data found in the pasted JSON response.</p>`;
            currentTableData = [];
            currentTableHeaders = [];
            return;
        }

        renderTable(dataToRender);
        downloadButtonsDiv.style.display = "block"; // Show download buttons after successful render

    } catch (error) {
        console.error("Error parsing JSON or rendering table:", error);
        showErrorMessage(`Error parsing JSON: ${error.message}. Please ensure it's valid JSON.`);
        tableContentDiv.innerHTML = `<p class="text-danger">An error occurred while parsing the JSON.</p>`;
        currentTableData = [];
        currentTableHeaders = [];
    } finally {
        hideLoading();
    }
}

// Function to recursively get all headers (keys) from nested objects
function getHeaders(data, prefix = '') {
    const headers = new Set();
    if (!data || data.length === 0) return [];

    // If data is a single object (e.g., query for one item), wrap it in an array
    const dataArray = Array.isArray(data) ? data : [data];

    dataArray.forEach(item => {
        for (const key in item) {
            if (item.hasOwnProperty(key)) {
                if (typeof item[key] === 'object' && item[key] !== null && !Array.isArray(item[key])) {
                    // Recursively get headers for nested objects
                    getHeaders([item[key]], `${prefix}${key}.`).forEach(nestedHeader => headers.add(nestedHeader));
                } else if (Array.isArray(item[key])) {
                    // For arrays, just add the field name. Complex array content won't be in simple tabular view.
                    headers.add(`${prefix}${key}`);
                } else {
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
            return ''; // Return empty string if path doesn't exist or is null/not an object
        }
        current = current[parts[i]];
    }
    // Special formatting for DateTimes
    if (typeof current === 'string' && current.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(Z|-\d{2}:\d{2})?$/)) { // Handle Z and timezone offsets
        try {
            const date = new Date(current);
            if (!isNaN(date.getTime())) { // Check if date is valid
                return date.toLocaleDateString('en-US') + ' ' + date.toLocaleTimeString('en-US'); // Specify locale for consistency
            }
        } catch (e) { /* fall through */ }
    }
    // Handle null/undefined display
    if (current === null || current === undefined) {
        return '';
    }
    // Handle numbers with fixed decimals (e.g., currency). Adjust fields as needed.
    if (typeof current === 'number' && (path.includes('Price') || path.includes('Tax') || path.includes('Profit') || path.includes('Rate') || path.includes('Amount'))) {
        return current.toFixed(2);
    }
    // Return arrays as a simple string representation
    if (Array.isArray(current)) {
        return `[${current.length} items]`; // Or current.join(', ') for simple arrays of primitives
    }
    return current.toString(); // Convert other values to string
}


function renderTable(data) {
    currentTableData = data; // Store the raw data for download
    currentTableHeaders = getHeaders(data); // Store the headers for download

    if (!data || data.length === 0) {
        tableContentDiv.innerHTML = '<p class="text-info">No data items found in the response to render a table.</p>';
        return;
    }

    if (currentTableHeaders.length === 0) {
        tableContentDiv.innerHTML = '<p class="text-warning">Could not determine table headers from the data structure.</p>';
        return;
    }

    let tableHTML = `<table class="table table-striped table-hover data-table"><thead><tr>`;
    currentTableHeaders.forEach(header => {
        const displayHeader = header.includes('.') ? header.split('.').pop().replace(/([A-Z])/g, ' $1').trim() : header.replace(/([A-Z])/g, ' $1').trim();
        tableHTML += `<th>${displayHeader}</th>`;
    });
    tableHTML += `</tr></thead><tbody>`;

    data.forEach(item => {
        tableHTML += `<tr>`;
        currentTableHeaders.forEach(headerPath => {
            tableHTML += `<td>${getNestedValue(item, headerPath)}</td>`;
        });
        tableHTML += `</tr>`;
    });

    tableHTML += `</tbody></table>`;
    tableContentDiv.innerHTML = tableHTML;
}

// --- NEW: Download Functions ---

function downloadTableData(format) {
    if (currentTableData.length === 0 || currentTableHeaders.length === 0) {
        showErrorMessage("No data available to download. Please parse a JSON response first.");
        return;
    }

    let fileContent = "";
    let fileName = "data";
    let mimeType = "";

    // Prepare headers for CSV (using the display format)
    const displayHeaders = currentTableHeaders.map(header => {
        return header.includes('.') ? header.split('.').pop().replace(/([A-Z])/g, ' $1').trim() : header.replace(/([A-Z])/g, ' $1').trim();
    });

    if (format === 'csv' || format === 'excel') {
        // CSV Content - ensure proper quoting for CSV
        fileContent += displayHeaders.map(h => `"${h.replace(/"/g, '""')}"`).join(',') + '\n'; // CSV header row

        currentTableData.forEach(item => {
            const row = currentTableHeaders.map(headerPath => {
                let value = getNestedValue(item, headerPath);
                // Convert value to string and escape for CSV
                let stringValue = String(value);
                if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
                    return `"${stringValue.replace(/"/g, '""')}"`;
                }
                return stringValue;
            }).join(',');
            fileContent += row + '\n';
        });

        if (format === 'csv') {
            fileName = "graphql_data.csv";
            mimeType = "text/csv;charset=utf-8;";
        } else { // Excel (using CSV as base for simple Excel)
            fileName = "graphql_data.xls"; // .xls is often recognized by Excel for CSV content
            mimeType = "application/vnd.ms-excel"; // More specific for Excel, but still CSV-like
            // For true XLSX, you'd need a library like SheetJS, which is beyond simple client-side JS
        }
    }

    const blob = new Blob([fileContent], { type: mimeType });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href); // Clean up
}


// Initial Load
showLoading();
msalInstance.handleRedirectPromise().catch(err => console.error(err));