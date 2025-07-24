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
const tableViewBtn = document.getElementById("tableViewBtn"); // NEW: Table View Button
const viewToggler = document.getElementById("viewToggler");
const playgroundContainer = document.getElementById("graphql-playground");
const voyagerContainer = document.getElementById("voyager-container");
const tableContainer = document.getElementById("table-view"); // NEW: Table View Container
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

    // Initialize Playground on login/redirect to ensure it's always ready
    initializeGraphQLPlayground();
    // Default to playground view
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
        // Force interactive login if silent fails (e.g., token expired or no valid session)
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
            // For simplicity, re-initialize if options change.
            // A more advanced approach would update Voyager's settings if its API supports it.
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
    tableContainer.style.display = "none"; // Hide table view
    playgroundBtn.classList.add("active");
    voyagerBtn.classList.remove("active");
    tableViewBtn.classList.remove("active"); // Remove active from table button
    contentWrapper.classList.remove('voyager-active', 'table-active'); // Remove all active classes
});

voyagerBtn.addEventListener("click", () => {
    playgroundContainer.style.display = "none";
    voyagerContainer.style.display = "block";
    tableContainer.style.display = "none"; // Hide table view
    voyagerBtn.classList.add("active");
    playgroundBtn.classList.remove("active");
    tableViewBtn.classList.remove("active"); // Remove active from table button
    contentWrapper.classList.remove('table-active'); // Remove table active class
    contentWrapper.classList.add('voyager-active'); // Add voyager active class

    const voyagerOptions = {
        skipRelay: skipRelayCheckbox.checked,
        skipDeprecated: skipDeprecatedCheckbox.checked,
    };
    initializeVoyager(voyagerOptions);
});

tableViewBtn.addEventListener("click", async () => { // NEW: Table View Button Click
    playgroundContainer.style.display = "none";
    voyagerContainer.style.display = "none";
    tableContainer.style.display = "block"; // Show table view
    tableViewBtn.classList.add("active");
    playgroundBtn.classList.remove("active");
    voyagerBtn.classList.remove("active");
    contentWrapper.classList.remove('voyager-active'); // Remove voyager active class
    contentWrapper.classList.add('table-active'); // Add table active class

    await fetchAndRenderTableData(); // Fetch and render data when switching to table view
});

// --- Voyager Controls ---
zoomSlider.addEventListener("input", (event) => {
    // This part depends on how GraphQL Voyager implements zoom.
    // As per the original code, it just logs. If direct CSS transform is desired:
    // voyagerContainer.style.transform = `scale(${event.target.value})`;
    // voyagerContainer.style.transformOrigin = `top left`;
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

// --- NEW: Table Data Fetching and Rendering ---

async function fetchAndRenderTableData() {
    showLoading();
    hideErrorMessage();
    const accessToken = await getAccessToken();

    if (!accessToken) {
        showErrorMessage("No access token available to fetch table data. Please log in.");
        hideLoading();
        return;
    }

    // Define the GraphQL query to fetch sales and customer data
    const query = `
        query {
            fact_sales(first: 50) { # Limiting to 50 for initial display, adjust as needed
                items {
                    SaleKey
                    InvoiceDateKey
                    Quantity
                    UnitPrice
                    TotalIncludingTax
                    Profit
                    dimension_customer {
                        Customer
                        Category
                        PostalCode
                    }
                    dimension_stock_item { # Also pulling stock item info for more robust table
                        StockItem
                        Brand
                    }
                }
            }
        }
    `;

    try {
        const response = await fetch(GRAPHQL_ENDPOINT, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${accessToken}`
            },
            body: JSON.stringify({ query: query })
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`GraphQL network response was not ok: ${response.status} - ${errorText}`);
        }

        const result = await response.json();

        if (result.errors) {
            console.error("GraphQL Errors:", result.errors);
            showErrorMessage(`GraphQL Errors: ${result.errors.map(e => e.message).join(", ")}`);
            tableContainer.innerHTML = `<p class="text-danger">Failed to load data due to GraphQL errors.</p>`;
            return;
        }

        const sales = result.data.fact_sales.items;
        renderTable(sales);

    } catch (error) {
        console.error("Error fetching table data:", error);
        showErrorMessage(`Error fetching table data: ${error.message}`);
        tableContainer.innerHTML = `<p class="text-danger">An error occurred while fetching data.</p>`;
    } finally {
        hideLoading();
    }
}

function renderTable(data) {
    if (!data || data.length === 0) {
        tableContainer.innerHTML = '<p class="text-info">No sales data available.</p>';
        return;
    }

    // Dynamically create table headers based on the first item's keys, including nested
    let headers = [
        "Sale Key", "Invoice Date", "Quantity", "Unit Price",
        "Total Including Tax", "Profit", "Customer Name", "Customer Category",
        "Customer Postal Code", "Stock Item", "Brand"
    ];

    let tableHTML = `<table class="table table-striped table-hover data-table"><thead><tr>`;
    headers.forEach(header => {
        tableHTML += `<th>${header}</th>`;
    });
    tableHTML += `</tr></thead><tbody>`;

    // Populate table rows
    data.forEach(item => {
        const customer = item.dimension_customer || {}; // Handle potential missing customer
        const stockItem = item.dimension_stock_item || {}; // Handle potential missing stock item

        tableHTML += `<tr>`;
        tableHTML += `<td>${item.SaleKey || ''}</td>`;
        tableHTML += `<td>${item.InvoiceDateKey ? new Date(item.InvoiceDateKey).toLocaleDateString() : ''}</td>`;
        tableHTML += `<td>${item.Quantity || ''}</td>`;
        tableHTML += `<td>${item.UnitPrice !== null ? item.UnitPrice.toFixed(2) : ''}</td>`;
        tableHTML += `<td>${item.TotalIncludingTax !== null ? item.TotalIncludingTax.toFixed(2) : ''}</td>`;
        tableHTML += `<td>${item.Profit !== null ? item.Profit.toFixed(2) : ''}</td>`;
        tableHTML += `<td>${customer.Customer || ''}</td>`;
        tableHTML += `<td>${customer.Category || ''}</td>`;
        tableHTML += `<td>${customer.PostalCode || ''}</td>`;
        tableHTML += `<td>${stockItem.StockItem || ''}</td>`;
        tableHTML += `<td>${stockItem.Brand || ''}</td>`;
        tableHTML += `</tr>`;
    });

    tableHTML += `</tbody></table>`;
    tableContainer.innerHTML = tableHTML;
}


// Initial Load
showLoading();
// The msalInstance.handleRedirectPromise() will eventually call showMainContent() or showLoginScreen()
// It handles the initial authentication flow after redirect.
msalInstance.handleRedirectPromise().catch(err => console.error(err));