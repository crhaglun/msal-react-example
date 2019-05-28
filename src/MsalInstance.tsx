import { UserAgentApplication, AuthError, AuthResponse } from "msal";

declare global {
    interface Window { userAgentApplication: UserAgentApplication }
}

export const createUserAgentApplication = () => {

    if (!window.userAgentApplication) {
        const msalConfig = {
            auth: {
                // Application registration ID
                clientId: 'c454a5cb-2667-4927-820a-89a1e25f0f8d',

                // Authority endpoint
                // Azure AD + consumer MSA: https://login.microsoftonline.com/common
                // Azure AD only:           https://login.microsoftonline.com/organizations
                // Specific AAD tenant:     https://login.microsoftonline.com/<tenant ID>  
                authority: "https://login.microsoftonline.com/common",
            }
        }

        window.userAgentApplication = new UserAgentApplication(msalConfig)

        const redirectCallback = (authErr: AuthError, response?: AuthResponse) => {
            console.log(authErr);
            console.log(response);
        }

        window.userAgentApplication.handleRedirectCallback(redirectCallback)
    }

    return window.userAgentApplication
}