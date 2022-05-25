require("dotenv").config();
const express = require("express");
const app = express();
require("isomorphic-fetch");

const { Client } = require("@microsoft/microsoft-graph-client");
const {
  TokenCredentialAuthenticationProvider,
} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const { ClientSecretCredential } = require("@azure/identity");

app.get("/createMeeting", async (req, res) => {
  try {
    const credential = new ClientSecretCredential(
      process.env.TENANT_ID,
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET
    );
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ["https://graph.microsoft.com/.default"],
    });

    const client = Client.initWithMiddleware({
      debugLogging: true,
      authProvider,
      // Use the authProvider object to create the class.
    });

    const onlineMeeting = {
      startDateTime: "2019-07-12T14:30:34.2444915-07:00",
      endDateTime: "2019-07-12T15:00:34.2464912-07:00",
      subject: `Demo Meeting (${Date.now().toString()})`,
      recordAutomatically: true,
    };

    const userId = "e8b297d6-2c29-4c35-a921-f1664be1ded0";

    const data = await client
      .api(`/users/${userId}/onlineMeetings`)
      .post(onlineMeeting);

    return res.send(data);
  } catch (error) {
    console.log(error);
    return res.send(error);
  }
});

app.listen(3000, (err) => {
  if (err) throw new Error(err);
  console.log("Server running on 3000");
});
