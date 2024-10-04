// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const router = require('express-promise-router')();
const graph = require('../helpers/graphHelper');
const dbHelper = require('../helpers/dbHelper');
const certHelper = require('../helpers/certHelper');

// GET /apponly/subscribe
router.get('/subscribe', async function (req, res) {
  const client = graph.getGraphClientForApp(req.app.locals.msalClient);

  // Ensure a certificate exists
  await certHelper.createSelfSignedCertificateIfNotExists(
    process.env.CERTIFICATE_PATH,
    process.env.PRIVATE_KEY_PATH,
    process.env.PRIVATE_KEY_PASSWORD,
  );

  // If in production, use the current host to receive notifications
  // In development, must use an ngrok proxy
  const notificationHost =
    process.env.NODE_ENV === 'production'
      ? `${req.protocol}://${req.hostname}`
      : process.env.NGROK_PROXY;

  try {
    const existingSubscriptions =
      await dbHelper.getSubscriptionsByUserAccountId('APP-ONLY');

    // Apps are only allowed one subscription to the /teams/getAllMessages resource
    // If we already had one, delete it so we can create a new one
    if (existingSubscriptions) {
      for (var existingSub of existingSubscriptions) {
        try {
          await client
            .api(`/subscriptions/${existingSub.subscriptionId}`)
            .delete();
        } catch (err) {
          console.error(err);
        }

        await dbHelper.deleteSubscription(existingSub.subscriptionId);
      }
    }

    /*
    * Create the subscription
    * in unserem fall brauchen wir hier die folgende anpassung:
    * resource: '/teams/getAllMessages' => '/users/{id}/events'
    * ich will eig. nicht eine subscription pro user sondern fuer einen Tenant
    * resource: '/users/events' oder so ..., weil einen User hab ich hier schon gar nicht.
    */
    const subscription = await client.api('/subscriptions').create({
      changeType: 'created',
      notificationUrl: `${notificationHost}/listen`,
      lifecycleNotificationUrl: `${notificationHost}/lifecycle`,
      resource: '/teams/getAllMessages',
      clientState: process.env.SUBSCRIPTION_CLIENT_STATE,
      includeResourceData: true,
      // To get resource data, we must provide a public key that
      // Microsoft Graph will use to encrypt their key
      // See https://learn.microsoft.com/graph/webhooks-with-resource-data#creating-a-subscription
      encryptionCertificate: certHelper.getSerializedCertificate(
        process.env.CERTIFICATE_PATH,
      ),
      encryptionCertificateId: process.env.CERTIFICATE_ID,
      expirationDateTime: new Date(Date.now() + 3600000).toISOString(),
    });

    // Save the subscription ID in the session
    req.session.subscriptionId = subscription.id;
    console.log(
      `Subscribed to Teams channel messages, subscription ID: ${subscription.id}`,
    );

    // Add subscription to the database
    await dbHelper.addSubscription(subscription.id, 'APP-ONLY');

    // Redirect to subscription page
    res.redirect('/watch');
  } catch (error) {
    req.flash('error_msg', {
      message: 'Error subscribing for Teams channel message notifications',
      debug: JSON.stringify(error, Object.getOwnPropertyNames(error)),
    });

    res.redirect('/');
  }
});

// GET /apponly/signout
router.get('/signout', async function (req, res) {
  // Delete the subscription from database and Graph
  const subscriptionId = req.session.subscriptionId;
  const msalClient = req.app.locals.msalClient;

  await dbHelper.deleteSubscription(subscriptionId);

  const client = graph.getGraphClientForApp(msalClient);

  try {
    await client.api(`/subscriptions/${subscriptionId}`).delete();

    req.session.subscriptionId = null;
  } catch (graphErr) {
    console.log(`Error deleting subscription from Graph: ${graphErr.message}`);
  }

  res.redirect('/');
});

module.exports = router;
