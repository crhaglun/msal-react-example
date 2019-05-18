import React from 'react';
import './App.css';
import { AuthError, AuthResponse, UserAgentApplication } from 'msal'
import { Shimmer } from 'office-ui-fabric-react';
import { Profile } from './Profile';
import { UnreadMailCount } from './UnreadMailCount';

function authCallback(authErr: AuthError, response?: AuthResponse) {
    console.log(authErr);
    console.log(response);
}

async function loginAsync(setToken: React.Dispatch<any>) {

    const msalConfig = {
        auth: {
            clientId: /* client ID goes here */,
            authority: "https://login.microsoftonline.com/common",
            redirectUri: 'http://localhost:3000'
        }
    }

    const request = {
        scopes: ["https://graph.microsoft.com/User.Read"]
    }

    const msalInstance = new UserAgentApplication(msalConfig);
    msalInstance.handleRedirectCallback(authCallback)

    try {
        const accessTokenResponse = await msalInstance.acquireTokenSilent(request);
        const token = accessTokenResponse.accessToken;

        setToken(token);
    }
    catch (err) {
        if (err.errorCode === "login_required" || err.errorCode === "token_renewal_error") {
            await msalInstance.loginRedirect(request);
        } else if (err.errorCode === "consent_required") {
            try {
                await msalInstance.acquireTokenRedirect(request);
            } catch (err2) {
                await msalInstance.loginRedirect(request);
            }
        } else {
            msalInstance.logout();
        }
    }
}

const App: React.FC = () => {

    const [token, setToken] = React.useState();

    var headers: Headers;

    if (token) {

        headers = new Headers();
        var bearer = "Bearer " + token;
        headers.append("Authorization", bearer);

        return (
            <>
                <Profile headers={ headers } />
                <UnreadMailCount headers={ headers } />
            </>
        )

    } else {
        loginAsync(setToken);

        return <Shimmer width="50%" />
    }
}

export default App;
