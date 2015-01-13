# Office 365 REST API Explorer for Sites #

## Overview ##

This app is an example that shows SharePoint developers how to perform HTTP requests to some of the REST endpoints in SharePoint sites hosted on Office 365.

The key goal of this project is to show SharePoint developers how they can use the tokens obtained with the Office 365 APIs to access the REST API in SharePoint.  Specifically, the app shows how to:

- Request an access token
- Get a new access token with a refresh token
- Clear the tokens in the cache
- Sign out the user
- Manage exceptions related to tokens


Visualizing and inspecting how the Office 365 APIs behave at the REST protocol level can be really helpful for learning and understanding them. This project uses the Office 365 APIs to get an access token that can be used with the REST endpoints in SharePoint. You can construct basic CRUD operations on lists, list items, and files viewing both the HTTP request and response. You can try out the REST endpoints that you need for your own apps and websites. The app has controls that let you modify the default CRUD operations.

You can perform CRUD operations on the following items in a SharePoint site:

- Lists
- List items
- Files

## Prerequisites and Configuration ##

This sample requires the following:

  - Visual Studio 2013 with Update 3.
  - [Office 365 API Tools version 1.3.41104.1](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155).
  - An [Office 365 developer site](https://portal.office.com/Signup/Signup.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK&ali=1).

###Configure the sample

Follow these steps to configure the sample.

   1. Open the O365-REST-API-Explorer.sln file using Visual Studio 2013.
   2. Register and configure the app to consume Office 365 services (detailed below).
   3. In the Solution Explorer window, choose Office365RESTAPIforSites project -> Add -> Connected Service. You must install the Office 365 API tools to see this menu option.
   4. In the **Services Manager** dialog box, choose **Office 365** and Register your app.
   5. On the sign-in dialog box, enter the username and password for your Office 365 tenant. 
	   - We recommend that you use your Office 365 Developer Site. Often, this user name will follow the pattern &lt;username&gt;@*&lt;tenant&gt;*.onmicrosoft.com. If you do not have a developer site, you can get a free developer site as part of your MSDN benefits or sign up for a free trial. Be aware that the user must be an tenant admin user—but for tenants created as part of an Office 365 developer site, this is likely to be the case already. Also developer accounts are usually limited to one sign-in.
	   - After you're signed in, you will see a list of all the services. Initially, no permissions are selected. 
   6. To register for the services used in this sample, choose the following permissions:
	- Sites 
		- Create or delete items and lists in all site collections
		- Edit or delete items in all site collections
		- Read items in all site collections
   7. After clicking OK in the Services Manager dialog box, assemblies for connecting to the Office 365 REST API will be added to your project.

## Build ##

1. Open the solution in Visual Studio and press F5.
2. In the initial screen, provide the URL of your SharePoint site.
3. Sign in with your organizational account to Office 365.

## Project Components of Interest ##

**Pages**

- ConfigureAppPage
- ItemsPage
- SplitPage

**Data Access Classes and Files**
   
- DataSource
- DataGroup
- DataItem
- RequestItem
- ResponseItem
- InitialData.json

## Troubleshooting ##

You may run into an authentication error after deploying and running if apps do not have the ability to access account information in the [Windows Privacy Settings](http://www.microsoft.com/security/online-privacy/windows.aspx) menu. Set **Let my apps access my name, picture, and other account info** to **On**.

Known issues

  - You need to provide a site that is at the root of the web application. For example: the https://*&lt;tenant&gt;*.sharepoint.com site. The sign in process fails with non-root site collections.

Copyright (c) Microsoft. All rights reserved.