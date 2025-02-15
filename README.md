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
const { simpleParser } = require("mailparser");
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
    let accessToken = req.session.token.access_token;

    const expiresAt = req.session.token.expires_at || 0;
    if (Date.now() >= expiresAt) {
      try {
        const newToken = await refreshToken({
          clientId: "YOUR_CLIENT_ID",
          clientSecret: "YOUR_CLIENT_SECRET",
          refreshToken: req.session.token.refresh_token,
        });
        req.session.token = {
          ...newToken,
          expires_at: Date.now() + newToken.expires_in * 1000,
        };
        accessToken = newToken.access_token;
      } catch (refreshError) {
        return res
          .status(401)
          .send('Session expired. Please <a href="/auth">re-authenticate</a>.');
      }
    }

    const mails = await readMail({
      userEmail: "YOUR_EMAIL",
      accessToken,
      folder: "INBOX",
    });

    if (!mails || mails.length === 0) {
      return res.status(404).send("No emails found.");
    }

    let html =
      "<html><head><meta charset='utf-8'><title>Emails</title></head><body>";

    await Promise.all(
      mails.map(async (mail, index) => {
        if (!mail.header || !mail.body) {
          return { error: "Incomplete email data" };
        }

        try {
          const emailContent = `${mail.header}\r\n${mail.body}`;
          const parsed = await simpleParser(emailContent);

          html += `<h2>Email #${index + 1}</h2>`;
          html += `<h3>From:</h3> ${parsed.from?.text || "Unknown Sender"}`;
          html += `<h3>To:</h3> ${parsed.to?.text || "Unknown Recipient"}`;
          html += `<h3>Subject:</h3> ${parsed.subject || "No Subject"}`;
          html += `<h3>Date:</h3> ${parsed.date || "Unknown Date"}`;
          html += `<h3>Body:</h3>${parsed.html || parsed.text || "No Content"}`;
          html += `<hr>`;

          return parsed;
        } catch (parseError) {
          return { error: "Failed to parse email content" };
        }
      })
    );

    html += "</body></html>";
    res.send(html);
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
