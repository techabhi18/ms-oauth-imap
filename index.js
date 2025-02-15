const { AuthorizationCode } = require("simple-oauth2");
const Imap = require("node-imap");
const qp = require("quoted-printable");

function validateParams(params, required) {
  for (const param of required) {
    if (!params[param]) {
      throw new Error(`${param} is required`);
    }
  }
}

function getAuthClient(clientId, clientSecret, redirectUri) {
  if (!clientId || !clientSecret || !redirectUri) {
    throw new Error("clientId, clientSecret, and redirectUri are required");
  }
  const clientConfig = {
    client: { id: clientId, secret: clientSecret },
    auth: {
      tokenHost: "https://login.microsoftonline.com",
      tokenPath: "/common/oauth2/v2.0/token",
      authorizePath: "/common/oauth2/v2.0/authorize",
    },
  };
  return new AuthorizationCode(clientConfig);
}

function generateAuthUrl({
  clientId,
  clientSecret,
  redirectUri,
  scope,
  state,
}) {
  validateParams({ clientId, clientSecret, redirectUri, scope }, [
    "clientId",
    "clientSecret",
    "redirectUri",
    "scope",
  ]);
  const client = getAuthClient(clientId, clientSecret, redirectUri);
  try {
    return client.authorizeURL({
      redirect_uri: redirectUri,
      scope,
      state: state || "random-state",
    });
  } catch (error) {
    throw new Error("Error generating auth URL: " + error.message);
  }
}

async function getToken({ clientId, clientSecret, redirectUri, code, scope }) {
  validateParams({ clientId, clientSecret, redirectUri, code, scope }, [
    "clientId",
    "clientSecret",
    "redirectUri",
    "code",
    "scope",
  ]);
  const client = getAuthClient(clientId, clientSecret, redirectUri);
  const options = { code, redirect_uri: redirectUri, scope };
  try {
    const token = await client.getToken(options);
    return token.token;
  } catch (error) {
    throw new Error("Error getting token: " + error.message);
  }
}

async function refreshToken({ clientId, clientSecret, redirectUri, token }) {
  validateParams({ clientId, clientSecret, redirectUri, token }, [
    "clientId",
    "clientSecret",
    "redirectUri",
    "token",
  ]);
  const client = getAuthClient(clientId, clientSecret, redirectUri);
  try {
    const accessToken = client.createToken(token);
    const refreshedToken = await accessToken.refresh();
    return refreshedToken.token;
  } catch (error) {
    throw new Error("Error refreshing token: " + error.message);
  }
}

function readMail({ userEmail, accessToken, folder = "INBOX" }) {
  return new Promise((resolve, reject) => {
    if (!userEmail || !accessToken) {
      return reject(new Error("userEmail and accessToken are required"));
    }
    const xoauth2Str = Buffer.from(
      `user=${userEmail}\x01auth=Bearer ${accessToken}\x01\x01`
    ).toString("base64");
    const imap = new Imap({
      user: userEmail,
      xoauth2: xoauth2Str,
      host: "outlook.office365.com",
      port: 993,
      tls: true,
      tlsOptions: { rejectUnauthorized: false },
    });
    let messages = [];
    imap.once("ready", () => {
      imap.openBox(folder, true, (err) => {
        if (err) {
          reject(new Error("Error opening mailbox: " + err.message));
          return;
        }
        const f = imap.seq.fetch("1:*", {
          bodies: ["HEADER.FIELDS (FROM TO SUBJECT DATE)", "TEXT"],
          struct: true,
        });
        f.on("message", (msg) => {
          let header = "";
          let body = "";
          msg.on("body", (stream, info) => {
            let buffer = "";
            stream.on("data", (chunk) => {
              buffer += chunk.toString("utf8");
            });
            stream.once("end", () => {
              if (info.which.indexOf("HEADER") !== -1) {
                header = buffer;
              } else if (info.which === "TEXT") {
                body = qp.decode(buffer).toString("utf8");
              }
            });
          });
          msg.once("end", () => {
            messages.push({ header, body });
          });
        });
        f.once("error", (err) => {
          reject(new Error("Fetch error: " + err.message));
        });
        f.once("end", () => {
          imap.end();
          resolve(messages);
        });
      });
    });
    imap.once("error", (err) => {
      reject(new Error("IMAP error: " + err.message));
    });
    try {
      imap.connect();
    } catch (error) {
      reject(new Error("Connection error: " + error.message));
    }
  });
}

module.exports = {
  generateAuthUrl,
  getToken,
  refreshToken,
  readMail,
};
