# Fabric GraphQL SPA with MSAL.js and GraphQL Playground

This project demonstrates a Single-Page Application (SPA) designed to interact with a Microsoft Fabric GraphQL API. It leverages Microsoft Authentication Library for JavaScript (MSAL.js) for authentication against Azure Active Directory (AAD) and integrates GraphQL Playground for an interactive GraphQL query experience.

## Features

-   **Azure AD Authentication:** Securely authenticates users via MSAL.js against Azure Active Directory.
-   **Interactive GraphQL Interface:** Utilizes GraphQL Playground, providing a rich environment for writing, executing, and exploring the schema of the Fabric GraphQL API.
-   **Modern UI:** Built with Bootstrap for a clean and responsive user interface.
-   **GitHub Pages Deployment:** Configured for easy deployment and hosting via GitHub Pages.

## Architecture and Flow

1.  **Client-Side Application:** The entire application runs in the user's browser (HTML, CSS, JavaScript).
2.  **Authentication (MSAL.js & Azure AD):**
    -   When the user clicks "Login", MSAL.js initiates an authentication flow (e.g., redirect flow) with Azure Active Directory.
    -   Upon successful authentication, Azure AD redirects the user back to the application's `redirectUri` (configured in `app.js` and Azure AD).
    -   MSAL.js then acquires an access token for the user, which is necessary to authorize requests to the Fabric GraphQL API.
3.  **GraphQL Playground Integration:**
    -   After successful login, the main content area displays the GraphQL Playground.
    -   The Playground is initialized with the Fabric GraphQL API endpoint.
    -   Crucially, the acquired access token is automatically injected into the `Authorization` header of all requests made by the GraphQL Playground, ensuring secure communication with the API.
4.  **Fabric GraphQL API Interaction:**
    -   Users can write GraphQL queries and mutations within the Playground.
    -   The Playground sends these queries to the specified Fabric GraphQL API endpoint.
    -   The API processes the request, using the provided access token for authorization, and returns the data.

## Technologies Used

-   **Frontend:** HTML, CSS, JavaScript
-   **Authentication:** [MSAL.js (Microsoft Authentication Library for JavaScript)](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)
-   **UI Framework:** [Bootstrap 5](https://getbootstrap.com/)
-   **GraphQL IDE:** [GraphQL Playground](https://github.com/graphql/graphql-playground)
-   **Hosting:** [GitHub Pages](https://pages.github.com/)
-   **Identity Provider:** [Azure Active Directory (AAD)](https://azure.microsoft.com/en-us/services/active-directory/)

## Setup and Deployment

### 1. Local Development

To run this application locally, you need a simple HTTP server. You can use Python's built-in server or Node.js `http-server`.

**Using Python:**

```bash
cd C:\Users\nalak\Documents\RnD\fabric-graphql-spa
python -m http.server 5500
```

Then, open your browser and navigate to `http://localhost:5500`.

### 2. GitHub Pages Deployment

This application is configured for deployment on GitHub Pages.

#### a. Push to GitHub

Ensure your project is a Git repository and pushed to GitHub. If you haven't already:

```bash
cd C:\Users\nalak\Documents\RnD\fabric-graphql-spa
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git
git push -u origin main
```

#### b. Configure GitHub Pages

1.  Go to your GitHub repository on the web.
2.  Navigate to **Settings > Pages**.
3.  Under "Build and deployment", set "Source" to **"Deploy from a branch"**.
4.  For "Branch", select `main` (or your primary branch) and choose `/ (root)` for the folder.
5.  Click **"Save"**.

    GitHub will automatically build and deploy your site. You can monitor the progress in the **"Actions"** tab of your repository. Once complete, the URL will be displayed at the top of the "Pages" settings page.

#### c. Azure AD App Registration Configuration (CRUCIAL!)

This is the most critical step for authentication to work on GitHub Pages.

1.  Go to the [Azure Portal](https://portal.azure.com/) and log in.
2.  Navigate to **Azure Active Directory > App registrations**.
3.  Find and select the application with the **Application (client) ID:** `4c072a54-b964-4f2e-a8cf-d571df4c58aa`.
4.  In the left-hand menu, click on **"Authentication"**.
5.  Under "Platform configurations", ensure you have a **"Single-page application"** platform configured.
6.  **Add the exact GitHub Pages URL as a Redirect URI:**
    `https://nalakan.github.io/-fabric-graphql-spa/`
    *   **Verify:** No typos, no extra spaces, correct `https://` protocol, and include the trailing slash `/`.
7.  Click **"Save"** at the top of the page.

    *Any mismatch here will result in `AADSTS50011` errors during login.*

## Usage

1.  Access the deployed application (e.g., `https://nalakan.github.io/-fabric-graphql-spa/`).
2.  Click the "Login" button to authenticate via Azure AD.
3.  Upon successful login, the GraphQL Playground will load.
4.  Use the Playground to explore the Fabric GraphQL API schema (via the "Schema" tab on the right) and execute queries.

## Troubleshooting

-   **`AADSTS50011: The redirect URI ... does not match...`**: This error indicates a mismatch between the `redirectUri` in your `app.js` and what's registered in your Azure AD App Registration. Double-check the URL in Azure Portal (Step 2.c) for exactness, including the trailing slash, and ensure it's configured under the "Single-page application" platform.
-   **Site not deploying on GitHub Pages**: Check the "Actions" tab in your GitHub repository for any failed "pages build and deployment" workflows. Ensure your GitHub Pages settings (branch, folder) are correct and saved.

```