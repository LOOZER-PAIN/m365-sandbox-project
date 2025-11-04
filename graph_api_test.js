// Microsoft Graph API sample
// This example lists the user's Microsoft 365 emails
// You can use this in your M365 Developer sandbox

const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

// Use your sandbox credentials here
const accessToken = "YOUR_ACCESS_TOKEN";

const client = Client.init({
  authProvider: (done) => {
    done(null, accessToken);
  },
});

async function getEmails() {
  try {
    const response = await client.api("/me/messages").get();
    console.log(response.value);
  } catch (error) {
    console.error(error);
  }
}

getEmails();












Added Microsoft Graph API sample
