import React from 'react';
import './App.css';
import { AuthError, AuthResponse, UserAgentApplication } from 'msal'
import { Profile } from './Profile';
import { ConsentControl } from './ConsentControl'
import { UnreadMailCount } from './UnreadMailCount';
import { Link } from 'office-ui-fabric-react';

const App: React.FC = () => {

    const msalConfig = {
        auth: {
            clientId: /* client ID goes here */,
            authority: "https://login.microsoftonline.com/common",
        }
    }

    const msalInstance = new UserAgentApplication(msalConfig);

    const redirectCallback = (authErr: AuthError, response?: AuthResponse) => {
        console.log(authErr);
        console.log(response);
    }

    msalInstance.handleRedirectCallback(redirectCallback)

    const getAccessToken = (scope: string) => {
        const request = {
            scopes: [scope]
        }
        return msalInstance.acquireTokenSilent(request)
    }

    const fetchWithScope = async (scope: string, query: string) => {
        try {
            const accessTokenResponse = await getAccessToken(scope);
            const token = accessTokenResponse.accessToken;

            const headers = new Headers({ "Authorization": "Bearer " + token });

            const options = {
                method: "GET",
                headers: headers
            };

            return await fetch(query, options);
        }
        catch (err) {
            console.log(err)
        }
    }

    const requestAccessToken = async (scope: string) => {
        const request = {
            scopes: [scope]
        }

        try {
            msalInstance.acquireTokenRedirect(request);
        } catch (err) {
            if (err.errorCode === "user_login_error") {
                msalInstance.loginRedirect(request);
            } else {
                throw err
            }
        }
    }

    return (
        <>
            <ConsentControl
                getAccessToken={getAccessToken}
                requestAccessToken={requestAccessToken}
                description="Read user profile"
                scope="https://graph.microsoft.com/User.Read" />
            <ConsentControl
                getAccessToken={getAccessToken}
                requestAccessToken={requestAccessToken}
                description="Read mail"
                scope="https://graph.microsoft.com/Mail.Read" />

            <Profile fetchWithScope={fetchWithScope} />
            <UnreadMailCount fetchWithScope={fetchWithScope} />

            <Link onClick={() => msalInstance.logout()}>Sign out</Link>
        </>
    )
}

export default App;
