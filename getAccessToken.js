const axios = require("axios");

const tenant_id = "ae613de3-9060-4327-ba38-f1665bd71a9a";
const client_id = "0f3605d1-2218-485f-aa82-678f14087e2c";
const client_secret = "KDM8Q~GpPQiOQr1OuGhf4jtNnMrmskm-7S5Lpcpu";
const scope = "https://graph.microsoft.com/.default";

// OAuth 2.0 token endpoint
const token_url = `https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`;

// Request payload
const payload = new URLSearchParams({
  grant_type: "client_credentials",
  client_id: client_id,
  client_secret: client_secret,
  scope: scope,
});

// Make the token request
axios
  .post(token_url, payload)
  .then((response) => {
    const access_token = response.data.access_token;
    console.log("Access Token:", access_token);
  })
  .catch((error) => {
    console.error("Error fetching access token:", error);
  });
