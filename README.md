---
topic: sample
products:
- Office 365
languages:
- C#
extensions:
  contentType: samples
  createdDate: 10/7/2014 10:32:11 AM
---
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

  - Windows 8.1
  - Visual Studio 2013 with Update 3.
  - [Office 365 API Tools version 1.3.41104.1](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155).
  - An Office 365 developer site. [Join the Office 365 Developer Program and get a free 1 year subscription to Office 365](https://aka.ms/devprogramsignup).

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

## Questions and comments

We'd love to get your feedback on the Office 365 REST API Explorer project. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/Office-365-REST-API-Explorer/issues) section of this repository.

Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [Office365] and [API].
  
## Additional resources ##

- [Office 365 APIs platform overview](https://msdn.microsoft.com/office/office365/howto/platform-development-overview)
- [Office 365 API code samples and videos](https://msdn.microsoft.com/office/office365/howto/starter-projects-and-code-samples)
- [Office developer code samples](http://dev.office.com/code-samples)
- [Office dev center](http://dev.office.com/)
- [Connecting to Office 365 in Windows Store, Phone, and universal apps](https://github.com/OfficeDev/O365-Win-Connect)
- [Office 365 Code Snippets for Windows](https://github.com/OfficeDev/O365-Win-Snippets)
- [Office 365 Starter Project for Windows Store App](https://github.com/OfficeDev/O365-Windows-Start)
- [Office 365 REST API Explorer for Sites](https://github.com/OfficeDev/Office-365-REST-API-Explorer)
- [Office 365 Profile sample for Windows](https://github.com/OfficeDev/O365-Win-Profile)

## Copyright ##

Copyright (c) 2014 Microsoft. All rights reserved.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
