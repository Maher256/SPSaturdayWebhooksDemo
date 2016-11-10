# SharePoint Saturday Sacramento 2016 code sample

This repo contains a sample NodeJS project that receives SharePoint Webhook Notifications.

# Additional resources 

* [SharePoint Development Center](http://dev.office.com/sharepoint)
* [SharePoint Patterns and Practices](http://aka.ms/sppnp)
* [Classic PnP repository for add-in samples](http://github.com/OfficeDev/PnP)

# Using the samples

To build and start using these projects, you'll need to clone and build the projects. 

Clone this repo by executing the following command in your console:

```
git clone https://github.com/SharePoint/sp-dev-samples.git
```

Navigate to the cloned repo folder which should be the same as the repo name:

```
cd sp-dev-samples
```

Update the config.js file with your Tenant name and U/P

run npm install

run nodemon index.js

Do a HTTP post to the server http://localhost/Webhook?validationtoken=IUGHUYTFTYWD from Postmon to test the subscription.

Do a HTTP post to the server http://localhost/Webhook and send it the following mock data to test a notification:

{
   "value":[
      {
         "subscriptionId":"d45d021d-8c53-490c-8dec-0882ed74b03f",
         "clientState":"00000000-0000-0000-0000-000000000000",
         "expirationDateTime":"2016-11-04T00:00:00.0000000Z",
         "resource":"0d5048e4-3539-425c-a83c-eef68c55fd40",
         "tenantId":"00000000-0000-0000-0000-000000000000",
         "siteUrl":"/sites/spssacramento",
         "webId":"974a1d3a-0d99-4538-a402-04385097e4d7"
      }
   ]
}

If both tests pass go ahead and subscribe to SharePoint's webhooks at /_api/web/Lists/getbytitle('Documents')/subscription

## Contributions

* [csom-node] (https://www.npmjs.com/package/csom-node)