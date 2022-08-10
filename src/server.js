// general dependencies

// <getDependencies>
// Express.js app server
import express from 'express';

// decode jwt token
import jwt_decode from 'jwt-decode';

// <getGraph>
// Microsoft Graph dependencies
import graph from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

// Play with Microsoft Graph 
//    https://developer.microsoft.com/en-us/graph/graph-explorer
// Debug JWT token 
//    https://jwt.ms/
function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate requests
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  return client;
}

// Use access token to get user's profile from Graph
async function getUsersProfile(accessToken) {
  try {
    const graphClient = getAuthenticatedClient(accessToken);

    const profile = await graphClient
      .api('/me')
      .get();

    return profile;

  } catch (err) {
    console.log(err);
    throw err;
  }
}
// </getGraph>

export const create = async () => {
  const app = express();

  // <routeHome>
  // Display form and table
  app.get('/', async (req, res) => {
    return res.send(`
    <!DOCTYPE html>
    <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="X-UA-Compatible" content="ie=edge">
        <title>Easy auth - Microsoft Graph Profile</title>
      </head>
      <body>
      <h1>Easy auth - Microsoft Graph Profile</h1>
      <p><a href="/access-token">Access token</a></p>
      <p><a href="/get-profile">Get profile from Microsoft Graph</a></p>
      <hr>
      <h5>Additional resources</h5>
      <p><a href="https://developer.microsoft.com/en-us/graph/graph-explorer">Explore with the Microsoft Graph interactive explorer</a></p>
      <p><a href="https://jwt.ms/">Decode access token with JWT.ms</a></p>
      </body>
    </html>
    `);
  });
  // </routeHome>

  // <routeInjectedToken>
  app.get('/access-token', async (req, res) => {

    const accessToken = req.headers['x-ms-token-aad-access-token'];
    if (!accessToken) return res.send('No access token found');

    const decoded = JSON.stringify(jwt_decode(accessToken));

    const curlCommandHello = `curl https://YOUR-RESOURCE-NAME.azurewebsites.net/hello -H "Accept: application/json" -H "Authorization: Bearer ${accessToken}"`;
    const curlCommandMe = `curl https://YOUR-RESOURCE-NAME.azurewebsites.net/me -H "Accept: application/json" -H "Authorization: Bearer ${accessToken}"`;
    return res.send(`${accessToken}<br><br>${decoded}<br><br>${curlCommandHello}<br><br>${curlCommandMe}`);
  });
  // </routeInjectedToken>

  // <routeGetProfile>
  app.get('/get-profile', async (req, res) => {

    let profile;
    let bearerToken;

    try {
      // should have `x-ms-token-aad-access-token`
      // insert from App Service if
      // MS AD identity provider is configured
      bearerToken = req.headers['x-ms-token-aad-access-token'];
      if (!bearerToken) return res.status(401).send('No access token found');

      profile = await getUsersProfile(bearerToken);

    } catch (err) {
      console.log(err);
      return res.status(500).json(err);
    } finally {
      return res.status(200).json(profile);
    }
  });
  // </routeGetProfile>

  // instead of 404 - just return home page
  app.get('*', (req, res) => {
    res.redirect('/');
  });

  return app;
};
