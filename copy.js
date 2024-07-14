const axios = require("axios");

const tenant_id = "ae613de3-9060-4327-ba38-f1665bd71a9a";
const client_id = "0f3605d1-2218-485f-aa82-678f14087e2c";
const client_secret = "KDM8Q~GpPQiOQr1OuGhf4jtNnMrmskm-7S5Lpcpu"; // Secret value, not the secret ID
const scope = "https://graph.microsoft.com/.default";

const siteHostname = "ceadll.sharepoint.com"; // Replace with your SharePoint hostname
const sitePath = "/sites/home"; // Replace with your SharePoint site path
const filePath =
  "/SiteAssets/SitePages/Home/CEAD-Employee-Handbook---Canada--January-2023-.pdf?web=1"; // Replace with your file path

async function getAccessToken() {
  const token_url = `https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`;

  const payload = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: client_id,
    client_secret: client_secret,
    scope: scope,
  });

  try {
    const response = await axios.post(token_url, payload);
    console.log("Access Token:", response.data.access_token); // Log the token
    return response.data.access_token;
  } catch (error) {
    console.error("Error fetching access token:", error);
    throw error;
  }
}

async function getSiteId(accessToken) {
  try {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
        },
      }
    );
    console.log("Site ID:", response.data.id);
    return response.data.id;
  } catch (error) {
    console.error(
      "Error fetching site ID:",
      error.response ? error.response.data : error.message
    );
    if (error.response.status === 403) {
      console.log(
        "Forbidden: The request was understood, but it has been refused or access is not allowed."
      );
      console.log("Possible reasons:");
      console.log("- Insufficient permissions");
      console.log("- Expired token");
      console.log("- Incorrect token");
    }
    throw error;
  }
}

async function testGetSiteId() {
  try {
    const accessToken = await getAccessToken();
    if (accessToken) {
      await getSiteId(accessToken);
    }
  } catch (error) {
    console.error("Test failed:", error);
  }
}

// Run the function
testGetSiteId();

/*
async function getDriveId(accessToken, siteId) {
  try {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    console.log("Drive ID:", response.data.value[0].id);
    return response.data.value[0].id;
  } catch (error) {
    console.error("Error fetching drive ID:", error);
    throw error;
  }
}*/
/*
async function getItemId(accessToken, driveId) {
  try {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:${filePath}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    console.log("Item ID:", response.data.id);
    return response.data.id;
  } catch (error) {
    console.error("Error fetching item ID:", error);
    throw error;
  }
}

async function retrieveDocument() {
  const accessToken = await getAccessToken();
  if (!accessToken) {
    console.error("Failed to obtain access token");
    return;
  }

  const siteId = await getSiteId(accessToken);
  const driveId = await getDriveId(accessToken, siteId);
  const itemId = await getItemId(accessToken, driveId);

  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}`;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    console.log("Document:", response.data);
  } catch (error) {
    console.error("Error retrieving document:", error);
  }
}

// Run the function
retrieveDocument();
*/
