import React, { ReactHTML } from 'react';
import './App.css';
import * as Msal from 'msal'
import { Shimmer } from 'office-ui-fabric-react/lib/Shimmer';
import { Profile } from './Profile';
import { UnreadMailCount } from './UnreadMailCount';
import { MarkEmailAsRead } from './MarkEmailAsRead';

function authCallback(authErr: Msal.AuthError, response?: Msal.AuthResponse) {
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
        scopes: ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Mail.Read"]
    }

    const msalInstance = new Msal.UserAgentApplication(msalConfig);
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
                <MarkEmailAsRead headers={ headers } />
            </>
        )

    } else {
        loginAsync(setToken);

        return <Shimmer width="50%" />
    }
}

export default App;
