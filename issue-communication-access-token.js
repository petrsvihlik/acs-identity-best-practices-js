const express = require("express");
const fetch = require('node-fetch');
const dotenv = require('dotenv');

const { PublicClientApplication, CryptoProvider, InteractionRequiredAuthError } = require('@azure/msal-node');
const { CommunicationIdentityClient } = require('@azure/communication-identity');
const { AzureCommunicationTokenCredential } = require('@azure/communication-common');
const { AbortController } = require("@azure/abort-controller");

dotenv.config();

const HOSTNAME = process.env.HOST || 'localhost';
const PORT = process.env.PORT || 80;
const HOST_URI = `http://${HOSTNAME}:${PORT}`;
const COMMUNICATION_SERVICES_CONNECTION_STRING = process.env.COMMUNICATION_SERVICES_CONNECTION_STRING;
const AAD_USER = process.env.AAD_USER;

// msal config
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: process.env.AUTHORITY,
    }
};

const publicClientApplication = new PublicClientApplication(msalConfig);
const provider = new CryptoProvider();

const app = express();
app.use(express.json());
app.use(express.urlencoded());
let pkceVerifier = null;

app.get('/', async (req, res) => {
    res.json({
        standard: `${HOST_URI}/standard`, cte: `${HOST_URI}/cte`,
    });
});

app.get('/cte', async (req, res) => {
    const { verifier, challenge } = await provider.generatePkceCodes();
    pkceVerifier = verifier;
    const authCodeUrlParameters = {
        scopes: ["https://auth.msft.communication.azure.com/Teams.ManageCalls"],
        redirectUri: `${HOST_URI}/redirect`,
        codeChallenge: challenge, // PKCE Code Challenge
        codeChallengeMethod: "S256" // PKCE Code Challenge Method 
    };

    publicClientApplication.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
        res.redirect(response);
    }).catch((error) => console.log(JSON.stringify(error)));
});

app.get('/redirect', async (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["https://auth.msft.communication.azure.com/Teams.ManageCalls"],
        redirectUri: `${HOST_URI}/redirect`,
        codeVerifier: pkceVerifier,
    };

    publicClientApplication.acquireTokenByCode(tokenRequest).then(async (response) => {
        const tokenResponse = await fetch(`${HOST_URI}/getTokenForTeamsUser`,
            {
                method: "POST",
                body: JSON.stringify({ teamsToken: response.accessToken }),
                headers: { 'Content-Type': 'application/json' }
            });
        const initialToken = (await tokenResponse.json()).communicationIdentityToken;

        const tokenCredential = new AzureCommunicationTokenCredential({
            tokenRefresher: async (abortSignal) => fetchTokenFromMyServerForUserCTE(abortSignal, AAD_USER),
            refreshProactively: true,
            // token: initialToken
        });

        const controller = new AbortController();
        let tkn = (await tokenCredential.getToken({ abortSignal: controller.signal }));
        res.send(tkn).sendStatus(200);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});

app.get('/standard', async (req, res) => {
    const tokenCredential = new AzureCommunicationTokenCredential({
        tokenRefresher: async (abortSignal) => fetchTokenFromMyServerForUser(abortSignal),
        refreshProactively: true,
        token: null
    });

    const controller = new AbortController();
    let tkn = (await tokenCredential.getToken({ abortSignal: controller.signal }));
    res.send(tkn).status(200);
});

/*** SERVER */
app.post('/getToken', async (req, res) => {
    // Custom logic to determine the communication user id
    let userId = await getCommunicationUserIdFromDb(req.body.username);
    // Get a fresh token
    const identityClient = new CommunicationIdentityClient(COMMUNICATION_SERVICES_CONNECTION_STRING);
    let communicationIdentityToken = await identityClient.getToken({ communicationUserId: userId }, ["chat", "voip"]);
    res.json({ communicationIdentityToken: communicationIdentityToken.token });
});

app.post('/getTokenForTeamsUser', async (req, res) => {
    const identityClient = new CommunicationIdentityClient(COMMUNICATION_SERVICES_CONNECTION_STRING);
    let communicationIdentityToken = await identityClient.getTokenForTeamsUser(req.body.teamsToken);
    res.json({ communicationIdentityToken: communicationIdentityToken.token });
});
/*** SERVER */

const getCommunicationUserIdFromDb = async function (username) {
    const identityClient = new CommunicationIdentityClient(COMMUNICATION_SERVICES_CONNECTION_STRING);
    const user = await identityClient.createUser();
    console.log(`ID of User ${username}: ${user.id}`);
    return user.communicationUserId;
};

const refreshAadToken = async function (abortSignal, account, forceRefresh) {
    if (abortSignal.aborted === true) throw new Error("Operation canceled");
    const renewRequest = {
        scopes: ["https://auth.msft.communication.azure.com/Teams.ManageCalls"],
        account: account,
        forceRefresh: forceRefresh
    };
    let tokenResponse = null;
    // Try to get the token silently without the user's interaction    
    await publicClientApplication.acquireTokenSilent(renewRequest).then(renewResponse => {
        tokenResponse = renewResponse;
    }).catch(async (error) => {
        // In case of an InteractionRequired error, send the same request in an interactive call
        if (error instanceof InteractionRequiredAuthError) {
            // You can choose the popup or redirect experience (`acquireTokenPopup` or `acquireTokenRedirect` respectively)
            publicClientApplication.acquireTokenPopup(renewRequest).then(function (renewInteractiveResponse) {
                tokenResponse = renewInteractiveResponse;
            }).catch(function (interactiveError) {
                console.log(interactiveError);
            });
        }
    });
    if (tokenResponse.expiresOn < (Date.now() + (10 * 60 * 1000)) && !forceRefresh) {
        // Make sure the token has at least 10-minute lifetime and if not, force-renew it
        tokenResponse = await refreshAadToken(abortSignal, teamsUser, true);
    }
    return tokenResponse;
}

const fetchTokenFromMyServerForUser = async function (abortSignal, username) {
    const response = await fetch(`${HOST_URI}/getToken`,
        {
            method: "POST",
            body: JSON.stringify({ username: username }),
            signal: abortSignal,
            headers: { 'Content-Type': 'application/json' }
        });

    if (response.ok) {
        const data = await response.json();
        return data.communicationIdentityToken;
    }
};

const fetchTokenFromMyServerForUserCTE = async function (abortSignal, username) {
    // MSAL.js v2 exposes several account APIs; the logic to determine which account to use is the responsibility of the developer. 
    // In this case, we'll use an account from the cache.    
    let teamsUser = (await publicClientApplication.getTokenCache().getAllAccounts()).find(u => u.username === username);

    // Get a fresh AAD token first
    let teamsTokenResponse = await refreshAadToken(abortSignal, teamsUser);

    // Use the fresh AAD token to exchange it for a Communication Identity access token
    const response = await fetch(`${HOST_URI}/getTokenForTeamsUser`,
        {
            method: "POST",
            body: JSON.stringify({ teamsToken: teamsTokenResponse.accessToken }),
            signal: abortSignal,
            headers: { 'Content-Type': 'application/json' }
        });

    if (response.ok) {
        const data = await response.json();
        return data.communicationIdentityToken;
    }
}

app.listen(PORT, () => console.log(`Teams token application started on ${PORT}!`))

