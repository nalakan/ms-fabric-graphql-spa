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

// GLOBAL DOM Elements that are always present immediately
const loginScreen = document.getElementById("loginScreen");
const mainContent = document.getElementById("mainContent");
const loginBtn = document.getElementById("loginBtn");
const logoutBtn = document.getElementById("logoutBtn");
const loadingOverlay = document.getElementById("loadingOverlay");
const errorMessageDiv = document.getElementById("errorMessage");
const viewToggler = document.getElementById("viewToggler");
const contentWrapper = document.querySelector('.content-wrapper');


const GRAPHQL_ENDPOINT = 'https://bb4b4fcd2a8943f0b63391db3f3c4f9e.zbb.graphql.fabric.microsoft.com/v1/workspaces/bb4b4fcd-2a89-43f0-b633-91db3f3c4f9e/graphqlapis/69ea77b8-daf1-45b5-9200-69e4826a1a5a/graphql';

// Global variables to store processed data and headers for download
let currentTableData = [];
let currentTableHeaders = [];

// *** Wrap the rest of your app.js logic in DOMContentLoaded ***
document.addEventListener('DOMContentLoaded', () => {

    // DOM Elements that might not be available immediately if their parent is display:none
    // It's safer to define these inside DOMContentLoaded
    const playgroundBtn = document.getElementById("playgroundBtn");
    const voyagerBtn = document.getElementById("voyagerBtn");
    const tableViewBtn = document.getElementById("tableViewBtn");

    const playgroundContainer = document.getElementById("graphql-playground");
    const voyagerContainer = document.getElementById("voyager-container");
    const tableContainer = document.getElementById("table-view");

    const jsonResponseInput = document.getElementById("jsonResponseInput");
    const parseAndShowTableBtn = document.getElementById("parseAndShowTableBtn");
    const tableContentDiv = document.getElementById("tableContent");
    const downloadButtonsDiv = document.getElementById("downloadButtons");
    const downloadCsvBtn = document.getElementById("downloadCsvBtn");
    const downloadExcelBtn = document.getElementById("downloadExcelBtn");

    const zoomSlider = document.getElementById("zoom-slider"); // Now safely inside
    const skipRelayCheckbox = document.getElementById("skip-relay-checkbox"); // Now safely inside
    const skipDeprecatedCheckbox = document.getElementById("skip-deprecated-checkbox"); // Now safely inside


    // --- Core UI Functions --- (No changes to functions themselves)
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

    // --- MSAL Authentication --- (No changes to functions themselves)
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

    // --- GraphQL Tools Initialization --- (No changes to functions themselves)
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
        if (!accessToken) {
            console.error("Voyager Introspection: No access token available. Cannot perform introspection.");
            showErrorMessage("Voyager requires an access token for schema introspection. Please log in.");
            throw new Error("No access token for introspection.");
        }

        try {
            console.log("Voyager Introspection: Sending introspection query to:", GRAPHQL_ENDPOINT);
            console.log("Voyager Introspection: Query being sent:", query);

            const response = await fetch(GRAPHQL_ENDPOINT, {
                method: 'post',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${accessToken}`
                },
                body: JSON.stringify({ query: query }),
            });

            if (!response.ok) {
                const errorBody = await response.text();
                console.error(`Voyager Introspection: Network response not OK: ${response.status} ${response.statusText}`, errorBody);
                showErrorMessage(`Voyager: Failed to fetch schema (${response.status} ${response.statusText}). Check network/console.`);
                throw new Error(`Failed to fetch schema: ${response.status} ${response.statusText}`);
            }

            const jsonResponse = await response.json();
            console.log("Voyager Introspection: Received JSON response:", jsonResponse);

            if (jsonResponse.errors) {
                console.error("Voyager Introspection: GraphQL errors in schema response:", jsonResponse.errors);
                showErrorMessage(`Voyager: GraphQL errors fetching schema: ${jsonResponse.errors.map(e => e.message).join(", ")}. Check console.`);
                throw new Error(`GraphQL errors fetching schema: ${jsonResponse.errors.map(e => e.message).join(", ")}`);
            }

            if (!jsonResponse.data || !jsonResponse.data.__schema) {
                console.error("Voyager Introspection: Response does not contain a valid introspection schema.", jsonResponse);
                showErrorMessage("Voyager: Invalid schema response format. Missing '__schema' field.");
                throw new Error("Invalid schema response format from introspection.");
            }

            console.log("Voyager Introspection: Schema fetched successfully!");
            return jsonResponse;
        } catch (error) {
            console.error("Voyager Introspection: Error during fetch or parsing:", error);
            showErrorMessage(`Voyager: An error occurred during schema introspection: ${error.message}.`);
            throw error;
        }
    }

    let voyagerInitialized = false;

    async function initializeVoyager(options = {}) {
        voyagerContainer.style.display = 'block';

        if (voyagerInitialized) {
            console.log("Voyager: Re-initializing with new options.");
        } else {
            console.log("Voyager: Initializing for the first time.");
        }

        try {
            GraphQLVoyager.init(voyagerContainer, { introspection: introspectionProvider, ...options });
            voyagerInitialized = true;
            console.log("Voyager: Initialization attempt complete.");
        } catch (error) {
            console.error("Voyager: Final initialization failed (caught in initializeVoyager):", error);
            voyagerContainer.innerHTML = `<div style="text-align:center; padding: 20px;">
                                            <h3>Could not render schema.</h3>
                                            <p>Check your network, access token, and browser console for detailed errors.</p>
                                            <p>Error: ${error.message}</p>
                                          </div>`;
        }
    }


    // --- View Toggler Event Listeners ---
    playgroundBtn.addEventListener("click", () => {
        playgroundContainer.style.display = "block";
        voyagerContainer.style.display = "none";
        tableContainer.style.display = "none";
        downloadButtonsDiv.style.display = "none";
        playgroundBtn.classList.add("active");
        voyagerBtn.classList.remove("active");
        tableViewBtn.classList.remove("active");
        contentWrapper.classList.remove('voyager-active', 'table-active');
    });

    voyagerBtn.addEventListener("click", () => {
        playgroundContainer.style.display = "none";
        voyagerContainer.style.display = "block";
        tableContainer.style.display = "none";
        downloadButtonsDiv.style.display = "none";
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

        tableContentDiv.innerHTML = '<p class="text-muted">Execute a query in the \'Playground\' tab, copy the JSON response, and paste it above to see data in a table.</p>';
        downloadButtonsDiv.style.display = "none";
        currentTableData = [];
        currentTableHeaders = [];
    });

    parseAndShowTableBtn.addEventListener("click", parseAndRenderTableData);
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
            skipRelay: skipDeprecatedCheckbox.checked,
            skipDeprecated: skipDeprecatedCheckbox.checked,
        });
    });

    // --- Table Data Parsing and Rendering --- (No changes to functions themselves)
    function parseAndRenderTableData() {
        showLoading();
        hideErrorMessage();
        tableContentDiv.innerHTML = '';
        downloadButtonsDiv.style.display = "none";

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
            for (const key in parsedResponse.data) {
                if (parsedResponse.data.hasOwnProperty(key)) {
                    const potentialData = parsedResponse.data[key];
                    if (potentialData && typeof potentialData === 'object' && potentialData.items && Array.isArray(potentialData.items)) {
                        dataToRender = potentialData.items;
                        break;
                    } else if (Array.isArray(potentialData)) {
                        dataToRender = potentialData;
                        break;
                    } else if (typeof potentialData === 'object' && potentialData !== null && Object.keys(potentialData).length > 0) {
                        dataToRender = [potentialData];
                        break;
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
            downloadButtonsDiv.style.display = "block";

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

    function getHeaders(data, prefix = '') {
        const headers = new Set();
        if (!data || data.length === 0) return [];
        const dataArray = Array.isArray(data) ? data : [data];
        dataArray.forEach(item => {
            for (const key in item) {
                if (item.hasOwnProperty(key)) {
                    if (typeof item[key] === 'object' && item[key] !== null && !Array.isArray(item[key])) {
                        getHeaders([item[key]], `${prefix}${key}.`).forEach(nestedHeader => headers.add(nestedHeader));
                    } else if (Array.isArray(item[key])) {
                        headers.add(`${prefix}${key}`);
                    } else {
                        headers.add(`${prefix}${key}`);
                    }
                }
            }
        });
        return Array.from(headers);
    }

    function getNestedValue(obj, path) {
        const parts = path.split('.');
        let current = obj;
        for (let i = 0; i < parts.length; i++) {
            if (current === null || typeof current !== 'object' || !current.hasOwnProperty(parts[i])) {
                return '';
            }
            current = current[parts[i]];
        }
        if (typeof current === 'string' && current.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(Z|-\d{2}:\d{2})?$/)) {
            try {
                const date = new Date(current);
                if (!isNaN(date.getTime())) {
                    return date.toLocaleDateString('en-US') + ' ' + date.toLocaleTimeString('en-US');
                }
            } catch (e) { /* fall through */ }
        }
        if (current === null || current === undefined) {
            return '';
        }
        if (typeof current === 'number' && (path.includes('Price') || path.includes('Tax') || path.includes('Profit') || path.includes('Rate') || path.includes('Amount'))) {
            return current.toFixed(2);
        }
        if (Array.isArray(current)) {
            return `[${current.length} items]`;
        }
        return current.toString();
    }


    function renderTable(data) {
        currentTableData = data;
        currentTableHeaders = getHeaders(data);

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

    // --- Download Functions --- (No changes to function itself)
    function downloadTableData(format) {
        if (currentTableData.length === 0 || currentTableHeaders.length === 0) {
            showErrorMessage("No data available to download. Please parse a JSON response first.");
            return;
        }

        let fileContent = "";
        let fileName = "data";
        let mimeType = "";

        const displayHeaders = currentTableHeaders.map(header => {
            return header.includes('.') ? header.split('.').pop().replace(/([A-Z])/g, ' $1').trim() : header.replace(/([A-Z])/g, ' $1').trim();
        });

        if (format === 'csv' || format === 'excel') {
            fileContent += displayHeaders.map(h => `"${h.replace(/"/g, '""')}"`).join(',') + '\n';

            currentTableData.forEach(item => {
                const row = currentTableHeaders.map(headerPath => {
                    let value = getNestedValue(item, headerPath);
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
            } else {
                fileName = "graphql_data.xls";
                mimeType = "application/vnd.ms-excel";
            }
        }

        const blob = new Blob([fileContent], { type: mimeType });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
    }

    // Initial Load - This handles MSAL redirect logic
    // It must be outside DOMContentLoaded if handleRedirectPromise() is designed
    // to potentially redirect the page before DOMContentLoaded fires.
    // However, it's fine here as it's the very first thing called *after*
    // the DOM is ready for element access.
    msalInstance.handleRedirectPromise().catch(err => console.error(err));
    showLoading(); // Show loading initially
}); // End of DOMContentLoaded