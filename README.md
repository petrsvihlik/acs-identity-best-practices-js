# Azure Communication Services - Identity API Playground

A node.js project demonstrating various Identity API flows and provides best practices with regards to handling user access tokens (credentials).

## Running locally

1. Configure .env file
> # ACS Config
> COMMUNICATION_SERVICES_CONNECTION_STRING="endpoint=https://<resource-name>.communication.azure.com/;accesskey=<key>"
> AAD_AUTHORITY="https://login.microsoftonline.com/<guid>"
> AAD_CLIENT_ID="<guid>"
> AAD_USER="<name>@<domain>.<tld>"

2. Run `npm install`
3. Run `node .\issue-communication-access-token.js` or "Run Current File" in VS Code to debug
