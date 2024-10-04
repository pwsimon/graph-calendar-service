// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const router = require('express-promise-router')();
const graph = require('../helpers/graphHelper');
const dbHelper = require('../helpers/dbHelper');

// GET /delegated/signin
router.get('/signin', async function (req, res) {
  // Start the authorization code flow by redirecting the
  // browser to Microsoft identity platforms authorization URL
  const urlParameters = {
    scopes: process.env.OAUTH_SCOPES.split(','),
    redirectUri: process.env.OAUTH_REDIRECT_URI,
    prompt: 'select_account',
  };

  try {
    const authUrl =
      await req.app.locals.msalClient.getAuthCodeUrl(urlParameters);
    res.redirect(authUrl);
  } catch (error) {
    console.log(`Error: ${error}`);
    req.flash('error_msg', {
      message: 'Error getting auth URL',
      debug: JSON.stringify(error, Object.getOwnPropertyNames(error)),
    });
    res.redirect('/');
  }
});

// GET /delegated/callback
router.get('/callback', async function (req, res) {
  // Microsoft identity platform redirects the browser here with the
  // authorization result
  const tokenRequest = {
    code: req.query.code,
    scopes: process.env.OAUTH_SCOPES.split(','),
    redirectUri: process.env.OAUTH_REDIRECT_URI,
  };

  try {
    const response =
      await req.app.locals.msalClient.acquireTokenByCode(tokenRequest);

    // Save the user's homeAccountId in their session
    req.session.userAccountId = response.account.homeAccountId;

    const client = graph.getGraphClientForUser(
      req.app.locals.msalClient,
      req.session.userAccountId,
    );

    // Get the user's profile from Microsoft Graph
    const user = await client.api('/me').select('displayName, mail').get();

    // Save user's name and email address in the session
    req.session.user = {
      name: user.displayName,
      email: user.mail,
    };

    console.log(`Logged in as ${user.displayName}`);

    // If in production, use the current host to receive notifications
    // In development, must use an ngrok proxy
    const notificationHost =
      process.env.NODE_ENV === 'production'
        ? `${req.protocol}://${req.hostname}`
        : process.env.NGROK_PROXY;

    /*
    * Create the subscription
    * in unserem fall brauchen wir hier die folgende anpassung:
    * resource: 'me/mailFolders/inbox/messages' => 'me/events' // calender anstelle von mail
    */
    const subscription = await client.api('/subscriptions').create({
      changeType: 'created,updated,deleted', // [changeType](https://learn.microsoft.com/en-us/graph/api/resources/subscription?view=graph-rest-1.0#properties)
      notificationUrl: `${notificationHost}/listen`,
      lifecycleNotificationUrl: `${notificationHost}/lifecycle`,
      resource: 'me/events',
      clientState: process.env.SUBSCRIPTION_CLIENT_STATE,
      includeResourceData: false,
      expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
    });

    // Save the subscription ID in the session
    req.session.subscriptionId = subscription.id;
    console.log(`Subscribed to: ${subscription.resource}, subscription ID: ${subscription.id}`);

    // Add the subscription to the database
    await dbHelper.addSubscription(subscription.id, req.session.userAccountId);

    // Redirect to subscription page
    res.redirect('/watch');
  } catch (error) {
    req.flash('error_msg', {
      message: 'Error completing authentication',
      debug: JSON.stringify(error, Object.getOwnPropertyNames(error)),
    });

    res.redirect('/');
  }
});

// GET /delegated/signout
router.get('/signout', async function (req, res) {
  // Delete the subscription from database and Graph
  const subscriptionId = req.session.subscriptionId;
  const msalClient = req.app.locals.msalClient;

  await dbHelper.deleteSubscription(subscriptionId);

  const client = graph.getGraphClientForUser(
    msalClient,
    req.session.userAccountId,
  );

  try {
    await client.api(`/subscriptions/${subscriptionId}`).delete();

    req.session.subscriptionId = null;
  } catch (graphErr) {
    console.log(`Error deleting subscription from Graph: ${graphErr.message}`);
  }

  try {
    // Remove user's account from MSAL cache
    const userAccount = await msalClient
      .getTokenCache()
      .getAccountByHomeId(req.session.userAccountId);

    await msalClient.getTokenCache().removeAccount(userAccount);

    req.session.userAccountId = null;
  } catch (msalErr) {
    console.log(`Error removing user from MSAL cache: ${msalErr.message}`);
  }

  res.redirect('/');
});

module.exports = router;
