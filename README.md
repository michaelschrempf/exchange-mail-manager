# exchange-mail-manager

An Angular single-page application that helps a Microsoft 365 user review and delete Exchange messages sent by a specific email address.

## What the app does

1. Prompts the user to sign in with Microsoft.
2. Accepts a sender email address.
3. Retrieves all messages in the signed-in mailbox that were sent by that address.
4. Lets the user review the matching messages and select which ones to delete.
5. Deletes the selected messages only after the user confirms the action.

## Prerequisites

- Node.js 20 or newer
- npm 10 or newer
- An Azure tenant where you can register a single-page application
- A Microsoft 365 account with access to the mailbox you want to clean up

## Azure tenant setup manual

### 1. Register the app

1. Sign in to the Azure portal.
2. Go to **Microsoft Entra ID** → **App registrations** → **New registration**.
3. Name the app `Exchange Mail Manager`.
4. Choose the supported account type that matches your tenant.
5. Under **Redirect URI**, select **Single-page application (SPA)** and enter `http://localhost:4200`.
6. Create the registration.

### 2. Collect tenant values

After the app registration is created, copy these values from the overview page:

- **Application (client) ID**
- **Directory (tenant) ID**

### 3. Configure API permissions

1. Open the app registration.
2. Go to **API permissions**.
3. Add these **delegated Microsoft Graph permissions**:
   - `User.Read`
   - `Mail.ReadWrite`
4. Grant tenant-wide admin consent if your tenant requires it.

### 4. Allow the signed-in user to access the mailbox

The signed-in Microsoft 365 account must have access to the Exchange mailbox that will be cleaned up. For a shared mailbox or another user's mailbox, grant the account the necessary Exchange permissions before using the app.

### 5. Configure the Angular app

Update `/tmp/workspace/michaelschrempf/exchange-mail-manager/src/environments/environment.ts` with your Azure values:

```ts
export const environment = {
  azure: {
    clientId: 'YOUR-CLIENT-ID',
    tenantId: 'YOUR-TENANT-ID',
    redirectUri: 'http://localhost:4200',
  },
};
```

## Local development

Install dependencies:

```bash
npm install
```

Run the Angular development server:

```bash
npm start
```

Then open `http://localhost:4200` in a browser.

## Build the app

```bash
npm run build
```

## Run tests

```bash
npm test -- --watch=false --browsers=ChromeHeadless
```

## How to use the app

1. Click **Sign in with Microsoft**.
2. Authenticate with the Microsoft 365 account that can access the target mailbox.
3. Enter the sender email address that should be cleaned up.
4. Click **Load messages** to retrieve all matching messages from the mailbox.
5. Review the subjects, dates, and previews.
6. Select the messages you want to remove, or use **Select all**.
7. Click **Delete selected messages** and confirm the prompt.

## Notes

- The app searches the signed-in mailbox with Microsoft Graph `/me/messages`.
- Messages are filtered by sender email address.
- Deletion is permanent from the mailbox context exposed by Microsoft Graph, so review the list carefully before confirming.
