# ms-oauth-imap

`ms-oauth-imap` is a Node.js package that integrates Microsoft OAuth authentication with IMAP email access. It provides functions to generate an OAuth URL, retrieve access tokens, refresh tokens, and read emails from specified folders.

## Features

- Generate an OAuth authentication URL
- Retrieve an access token using an authorization code
- Refresh expired tokens
- Read emails from specified mailboxes (INBOX, Sent, Drafts, etc.)

## Installation

Run the following command to install the package:

```sh
npm install ms-oauth-imap
```

## Prerequisites

- A Microsoft account with access to Outlook email.
- A registered application on the [Azure Portal](https://portal.azure.com) to obtain your `clientId` and `clientSecret`.
- A configured redirect URI for your application.

## Usage

### 1. Import the Package

```js
const {
  generateAuthUrl,
  getToken,
  refreshToken,
  readMail,
} = require("ms-oauth-imap");
```

### 2. Generate the OAuth URL

```js
const authUrl = generateAuthUrl({
  clientId: "YOUR_CLIENT_ID",
  clientSecret: "YOUR_CLIENT_SECRET",
  redirectUri: "http://localhost:4000/auth/callback",
  scope:
    "openid offline_access https://outlook.office.com/IMAP.AccessAsUser.All",
  state: "optional-state",
});
```

### 3. Retrieve the Access Token

```js
const token = await getToken({
  clientId: "YOUR_CLIENT_ID",
  clientSecret: "YOUR_CLIENT_SECRET",
  redirectUri: "http://localhost:4000/auth/callback",
  code: "AUTHORIZATION_CODE_FROM_QUERY",
  scope:
    "openid offline_access https://outlook.office.com/IMAP.AccessAsUser.All",
});
```

### 4. Refresh the Token

```js
const refreshedToken = await refreshToken({
  clientId: "YOUR_CLIENT_ID",
  clientSecret: "YOUR_CLIENT_SECRET",
  refreshToken: token.refresh_token,
});
```

### 5. Read Emails from a Folder

```js
const emails = await readMail({
  userEmail: "YOUR_EMAIL_ADDRESS",
  accessToken: token.access_token,
  folder: "INBOX", // Can be 'INBOX', 'Sent', 'Drafts', etc.
});
```

## Sample Integration with Express

```js
const express = require("express");
const session = require("express-session");
const {
  generateAuthUrl,
  getToken,
  refreshToken,
  readMail,
} = require("ms-oauth-imap");
const app = express();

app.use(session({ secret: "secret", resave: false, saveUninitialized: true }));

app.get("/auth", (req, res) => {
  const url = generateAuthUrl({
    clientId: "YOUR_CLIENT_ID",
    clientSecret: "YOUR_CLIENT_SECRET",
    redirectUri: "http://localhost:4000/auth/callback",
    scope:
      "openid offline_access https://outlook.office.com/IMAP.AccessAsUser.All",
  });
  res.redirect(url);
});

app.get("/auth/callback", async (req, res) => {
  const code = req.query.code;
  try {
    const token = await getToken({
      clientId: "YOUR_CLIENT_ID",
      clientSecret: "YOUR_CLIENT_SECRET",
      redirectUri: "http://localhost:4000/auth/callback",
      code,
      scope:
        "openid offline_access https://outlook.office.com/IMAP.AccessAsUser.All",
    });
    req.session.token = token;
    res.send("Authentication successful");
  } catch (error) {
    res.status(500).send(error.message);
  }
});

app.get("/refresh-token", async (req, res) => {
  if (!req.session.token) {
    return res.status(401).send("User not authenticated");
  }
  try {
    const newToken = await refreshToken({
      clientId: "YOUR_CLIENT_ID",
      clientSecret: "YOUR_CLIENT_SECRET",
      refreshToken: req.session.token.refresh_token,
    });
    req.session.token = newToken;
    res.send("Token refreshed successfully");
  } catch (error) {
    res.status(500).send(error.message);
  }
});

app.get("/read-mail", async (req, res) => {
  if (!req.session.token) {
    return res.status(401).send("User not authenticated");
  }
  try {
    const mails = await readMail({
      userEmail: "YOUR_EMAIL_ADDRESS",
      accessToken: req.session.token.access_token,
      folder: "INBOX",
    });
    res.json(mails);
  } catch (error) {
    res.status(500).send(error.message);
  }
});

app.listen(4000, () => {
  console.log("Server running on port 4000");
});
```

## Error Handling

This package validates required parameters and throws descriptive errors if any parameter is missing or invalid.

## License

MIT License

## Contributing

Contributions are welcome. Open an issue or submit a pull request to improve the package.
