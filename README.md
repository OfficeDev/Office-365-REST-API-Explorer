# Office 365 REST API Explorer for Sites #

## Overview ##

Visualizing and inspecting how the Office 365 APIs behave at the REST protocol level can be really helpful for learning and understanding them. This project uses the Office 365 APIs to get an access token that can be used with the REST endpoints in SharePoint. You can construct basic CRUD operations on lists, list items, and files viewing both the HTTP request and response. You can try out the REST endpoints that you need for your own apps and websites. The app has controls that let you modify the default CRUD operations.

You can perform CRUD operations on the following items in a SharePoint site:

- Lists
- List items
- Files

## Prerequisites and Configuration ##

This sample requires the following:

  - Visual Studio 2013 with Update 3.
  - [Office 365 API Tools version 1.1.728](http://visualstudiogallery.msdn.microsoft.com/7e947621-ef93-4de7-93d3-d796c43ba34f).
  - An [Office 365 developer site](https://portal.office.com/Signup/Signup.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK&ali=1).

###Configure the sample

Follow these steps to configure the sample.

   1. Open the Office365RESTAPIExplorerforSites.sln file using Visual Studio 2013.
   2. Register and configure the app to consume Office 365 services (detailed below).
   3. In the Solution Explorer window, choose Office365RESTAPIforSites project -> Add -> Connected Service. You must install the Office 365 API tools to see this menu option.
   4. A **Services Manager** dialog box will appear. Choose **Office 365** and Register your app.
   5. On the sign-in dialog box, enter the username and password for your Office 365 tenant. We recommend that you use your Office 365 Developer Site. Often, this user name will follow the pattern &lt;username&gt;@*&lt;tenant&gt;*.onmicrosoft.com. If you do not have a developer site, you can get a free developer site as part of your MSDN benefits or sign up for a free trial. Be aware that the user must be an tenant admin user—but for tenants created as part of an Office 365 developer site, this is likely to be the case already. Also developer accounts are usually limited to one sign-in.
   6. After you're signed in, you will see a list of all the services. Initially, no permissions are selected. 
   7. To register for the services used in this sample, choose the following permissions:
	- Sites 
		- Create or delete items and lists in all site collections (preview)
		- Edit or delete items in all site collections (preview)
		- Read items in all site collections (preview)
   8. After clicking OK in the Services Manager dialog box, assemblies for connecting to the Office 365 REST API will be added to your project.
   
Note: After adding the connected service, a sample file is added to the solution: SitesApiSample.cs. You may delete this file from the solution since the app has no dependencies on this code file.

## Build ##

1. After you've loaded the solution in Visual Studio, press F5 to build and debug.
2. Provide the URL of your SharePoint site.
3. Sign in with your organizational account to Office 365.

## Project Components of Interest ##

**Pages**

- ConfigureAppPage
- ItemsPage
- SplitPage

**Data Access Classes**
   
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
  - You get the following error message: "Invalid JWT token. The token is expired." This means that the access token has expired. Go to the start page and provide your credentials to renew the access token.

Copyright (c) Microsoft. All rights reserved.