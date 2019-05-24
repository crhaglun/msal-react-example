import React from 'react';
import './App.css';
import { Profile } from './Profile';
import { ConsentControl } from './ConsentControl'
import { UnreadMailCount } from './UnreadMailCount';
import { Link } from 'office-ui-fabric-react';
import { createUserAgentApplication } from './MsalInstance';
import { useAuthenticationState } from './useAuthenticationStateHook';

const App: React.FC = () => {

    const userAgentApplication = createUserAgentApplication()

    const userAuth = useAuthenticationState(
        userAgentApplication,
        "Read user profile",
        ["https://graph.microsoft.com/User.Read"])

    const mailAuth = useAuthenticationState(
        userAgentApplication,
        "Read mail",
        ["https://graph.microsoft.com/Mail.Read"])

    const combinedAuth = useAuthenticationState(
        userAgentApplication,
        "Read user profile + mail",
        ["https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Mail.Read"])

    return (
        <>
            <ConsentControl
                accessToken={userAuth.authResponse}
                requestAccessToken={userAuth.requestAccessToken}
                description={userAuth.description} />
            <ConsentControl
                accessToken={mailAuth.authResponse}
                requestAccessToken={mailAuth.requestAccessToken}
                description={mailAuth.description} />
            <ConsentControl
                accessToken={combinedAuth.authResponse}
                requestAccessToken={combinedAuth.requestAccessToken}
                description={combinedAuth.description} />
            <br />
            <Profile authenticationHeaders={userAuth.authHeaders} />
            <UnreadMailCount authenticationHeaders={mailAuth.authHeaders} />
            <Link onClick={() => userAgentApplication.logout()}>Sign out</Link>
        </>
    )
}

export default App;
