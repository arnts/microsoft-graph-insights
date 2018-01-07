# Example app to demonstrate the Microsoft Graph insights APIs

Insights are relationships calculated using advanced analytics and machine learning techniques. Through the Microsoft Graph API (and the underlying Office Graph) you can get different types of insights: Trending, Used, Shared. This example app uses the me/insights/used API, which returns documents viewed and modified by a user. More details can be found [here](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/insights).

To run this example app, you need to: 1) register the app with the Azure AD identity service, 2) configure the app environment, 3) use an Office 365 account to login (and upload content to the cloud so you get data returned through the Graph APIs). A simple and fun little excercise. 

## 1. Register the app

Head over to the [Application Registration Portal](https://apps.dev.microsoft.com) to quickly get an application ID and secret.

1. Using the Sign in link, sign in with your work or school account (Office 365).
2. Click the Add an app button. Enter microsoft-graph-demo for the name and click Create application.
3. Locate the Application Secrets section, and click the Generate New Password button. Copy the password now and save it to a safe place. Once you've copied the password, click Ok.
4. Locate the Platforms section, and click Add Platform. Choose Web, then enter http://localhost:8000/authorize under Redirect URIs.
5. Click Save to complete the registration. Copy the Application Id and save it along with the password you copied earlier. You'll need those values soon.

Note: In order to use the Application Registration Portal, you need either an Office 365 work or school account, or a Microsoft account. If you don't have either of these, you have a number of options:

1. Sign up for a new Microsoft account [here](https://www.outlook.com).
2. You can obtain an Office 365 subscription in a couple of different ways: You can get a free one-year Office 365 Developer subscription by signing up for the [Office Developer program](http://dev.office.com/devprogram), or you can signup for [a 25-user free trial](https://portal.office.com/Signup/Signup.aspx?OfferId=467eab54-127b-42d3-b046-3844b860bebf&dl=O365_BUSINESS_PREMIUM&alo=1&lc=1033&ali=1#0) of the Office 365 Business subscription.

## 2. Configure the app env

Set environment variables using set or export (replace with values from the Application Registration Portal:

```bash
export APP_ID='YOUR APP ID HERE'
export APP_SECRET='YOUR APP PASSWORD HERE'
```

## 3. Run the app

Run from the command line

```bash
$ npm install
``` 

```bash
$ npm start
``` 

Use your browser, and login using your Office 365 account at http://localhost:8000. Remember to [upload your files to Office 365](https://www.office.com).

# Resources

1. https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/insights 
2. https://github.com/microsoftgraph/msgraph-sdk-javascript
3. https://developer.microsoft.com/en-us/graph

