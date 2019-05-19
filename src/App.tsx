import React from 'react';
import './App.css';
import { AuthError, AuthResponse, UserAgentApplication } from 'msal'
import { Profile } from './Profile';
import { UnreadMailCount } from './UnreadMailCount';
import { PrimaryButton, Link } from 'office-ui-fabric-react';

function authCallback(authErr: AuthError, response?: AuthResponse) {
    console.log(authErr);
    console.log(response);
}

const App: React.FC = () => {

    const msalConfig = {
        auth: {
            clientId: /* client ID goes here */,
            authority: "https://login.microsoftonline.com/common",
        }
    }

    const msalInstance = new UserAgentApplication(msalConfig);
    msalInstance.handleRedirectCallback(authCallback)

    const getAccessToken = async (scope: string) => {
        const request = {
            scopes: [scope]
        }

        try {
            return await msalInstance.acquireTokenSilent(request);
        }
        catch (err) {
            console.log(err)

            if (err.errorCode === "login_required" || err.errorCode === "token_renewal_error") {
                await msalInstance.loginRedirect(request);
            } else if (err.errorCode === "consent_required") {
                try {
                    await msalInstance.acquireTokenRedirect(request);
                } catch (err2) {
                    await msalInstance.loginRedirect(request);
                }
            }
        }
    }

    const scopedQuery = async (scope: string, query: string) => {
        const accessTokenResponse = await getAccessToken(scope);

        if (accessTokenResponse) {
            const token = accessTokenResponse.accessToken;

            const headers = new Headers({ "Authorization": "Bearer " + token });

            const options = {
                method: "GET",
                headers: headers
            };

            return await fetch(query, options);
        }
    }

    return (
        <>
            <Profile scopedQuery={scopedQuery} />
            <UnreadMailCount scopedQuery={scopedQuery} />
            <Link onClick={() => msalInstance.logout()}>Sign out</Link>
        </>
    )
}

export default App;
