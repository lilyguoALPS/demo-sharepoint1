const axios = require("axios");

const tenant_id = "ae613de3-9060-4327-ba38-f1665bd71a9a";
const client_id = "0f3605d1-2218-485f-aa82-678f14087e2c";
const client_secret = "KDM8Q~GpPQiOQr1OuGhf4jtNnMrmskm-7S5Lpcpu"; // Secret value, not the secret ID
const scope = "https://graph.microsoft.com/.default";

//ceadll.sharepoint.com/sites/APIS-Development/Shared Documents/Knowledge Share/Training/READ ME.txt
const siteHostname = "ceadll.sharepoint.com";
const sitePath = "/sites/APIS-Development";
const libraryName = "Documents";
const filePath = "Knowledge Share/Training/READ ME.txt";

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

function extractUrlParts(sharepointUrl) {
  const url = new URL(sharepointUrl);
  const siteHostname = url.hostname;
  const pathParts = url.pathname.split("/");
  const sitePath = `/sites/${pathParts[2]}`;
  const libraryName = pathParts[3]; // This will be 'Shared Documents' or equivalent
  const filePath = url.pathname.split("/").slice(4).join("/");
  console.log(libraryName);
  return { siteHostname, sitePath, libraryName, filePath };
}

//extractUrlParts("https://ceadll.sharepoint.com/sites/APIS-Development/Shared Documents/Knowledge Share/Training/READ ME.txt");

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
//testGetSiteId();

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

    const drives = response.data.value;
    const drive = drives.find((d) => d.name === libraryName);
    if (!drive) {
      throw new Error(`Drive not found for library name: ${libraryName}`);
    }
    return drive.id;
    console.log("Drive ID:", response.data.value);
    //return response.data.value[1].id;
  } catch (error) {
    console.error("Error fetching drive ID:", error);
    throw error;
  }
}

async function getItemId(accessToken, driveId) {
  try {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}`,
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
async function retrieveDocument1() {
  const accessToken = await getAccessToken();
  if (!accessToken) {
    console.error("Failed to obtain access token");
    return;
  }

  const siteId = await getSiteId(accessToken);
  const driveId = await getDriveId(accessToken, siteId);
  const itemId = await getItemId(accessToken, driveId);
  console.log(itemId);
  //const itemId = "01EMSVNK5TUMLIWKZV6VHLZUUATHWNOBKJ";
  //console.log(itemId);
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

async function retrieveDocument() {
  const accessToken = await getAccessToken();
  if (!accessToken) {
    console.error("Failed to obtain access token");
    return;
  }

  const siteId = await getSiteId(accessToken);
  const driveId = await getDriveId(accessToken, siteId);
  //const itemId = await getItemId(accessToken, driveId);
  const itemId = "01EMSVNK5TUMLIWKZV6VHLZUUATHWNOBKJ";
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

async function retrieveDocumentContent() {
  const accessToken = await getAccessToken();
  if (!accessToken) {
    console.error("Failed to obtain access token");
    return;
  }

  const siteId = await getSiteId(accessToken);
  const driveId = await getDriveId(accessToken, siteId);
  const itemId = await getItemId(accessToken, driveId);
  //const itemId = "01EMSVNK5TUMLIWKZV6VHLZUUATHWNOBKJ";
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/content`;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
      responseType: "arraybuffer", // This is important to correctly handle binary data
    });
    console.log("Document content retrieved successfully");

    // If you want to convert the binary content to a string (assuming it's a text file):
    const content = new TextDecoder("utf-8").decode(response.data);
    console.log("Document content:", content);

    // For other types of files (e.g., PDFs, images), you might need to handle them differently
  } catch (error) {
    console.error("Error retrieving document:", error);
  }
}

// Run the function
//retrieveDocumentContent();

async function checkDriveId() {
  const accessToken = await getAccessToken();
  const siteId = await getSiteId(accessToken);
  const driveId = await getDriveId(accessToken, siteId);
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}`;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    console.log("Drive ID is valid:", response.data);
  } catch (error) {
    if (error.response) {
      console.error("Error response status:", error.response.status);
      if (error.response.status === 404) {
        console.error("Drive ID not found.");
      } else {
        console.error("Error response data:", error.response.data);
      }
    } else {
      console.error("Error:", error.message);
    }
  }
}

async function listRootItems() {
  const accessToken = await getAccessToken();
  const siteId = await getSiteId(accessToken);
  const driveId = await getDriveId(accessToken, siteId);
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    console.log("Root items:", response.data);
  } catch (error) {
    if (error.response) {
      console.error("Error response status:", error.response.status);
      console.error("Error response data:", error.response.data);
    } else {
      console.error("Error:", error.message);
    }
  }
}

async function listFolderItems() {
  const accessToken = await getAccessToken();
  const siteId = await getSiteId(accessToken);
  const driveId = await getDriveId(accessToken, siteId);

  //const folderId = "01EMSVNKZFCYEFS6S25NBJX5ZQ4PW23CB5";//const folderId = "01EMSVNK6P7B4USNCLTBA24ZT2GUVHPVIC";/////////////////
  //const folderId = "01EMSVNKYMAOPTFQRGKFGLEENAK6WCKHJR";
  //const folderId = "01EMSVNK6P7B4USNCLTBA24ZT2GUVHPVIC"; //Knowledge Share
  const folderId = "01EMSVNK2WR2AGFZLWG5AJBEIVN3ZTDZRY"; //Knowledge Share/Training
  //const folderId = "'01EMSVNK5TUMLIWKZV6VHLZUUATHWNOBKJ'"; //Knowledge Share/Training/READ ME.txt
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children`;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    console.log("Folder items:", response.data);
  } catch (error) {
    if (error.response) {
      console.error("Error response status:", error.response.status);
      console.error("Error response data:", error.response.data);
    } else {
      console.error("Error:", error.message);
    }
  }
}

//listRootItems();
//listFolderItems();
//checkDriveId();
//retrieveDocument1();
retrieveDocumentContent();
