# msal-react-example
Rudimentary example of using MSAL and Microsoft Graph in a React app

### Prerequisites

This sample requires the following:  

* [Node.js](https://nodejs.org/). Node is required to run the sample on a development server and to install dependencies.
* A [work or school account](https://dev.office.com/devprogram) or a [personal Microsoft account](https://account.microsoft.com/account)
  
## Register the application

1. Navigate to the Azure portal [App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page.

2. Choose **New registration**.

3. When the **Register an application page** appears, enter your application's registration information:

    * In the **Name** section, enter the application name, for example `MSAL React Example`
    * Change **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook)**.
    * In the Redirect URI (optional) section, select **Web** in the combo-box and enter the following redirect URI: `https://localhost:3000/`.

4. Select **Register** to create the application.

   The registration overview page displays, listing the properties of your app.

5. Copy the **Application (client) ID** and record it. This is the unique identifier for your app. You'll use this value to configure your app.

6. Select the **Authentication** section.
    * In the **Advanced settings** | **Implicit grant** section, check **Access tokens** and **ID tokens** as this sample requires
    the [Implicit grant flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-implicit-grant-flow) to be enabled to
    sign in the user and call an API.

7. Select **Save**.

## Build and run the sample

1. Clone or download the MSAL React Sample

2. Using your favorite IDE, open App.tsx in the *src* directory.

3. Replace the **clientId** placeholder value with the application ID of your registered Azure application.

4. Open a command prompt in the sample's root directory, and run the following command to install project dependencies.

  ```
  npm install
  ```

5. After the dependencies are installed, run the following command to start the development server.

  ```
  npm start
  ```

6. Navigate to *http://localhost:3000* in your web browser.

7. Sign in with a work or school account, or a personal Microsoft account. 
