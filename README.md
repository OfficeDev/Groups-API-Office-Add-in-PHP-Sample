# Groups-API-Office-Add-in-PHP-Sample
This repository contains a PHP code sample that connects to the Office 365 Groups API as both a stand-alone web application and from within an Office add-in.

## Environment Setup ##
As a PHP sample, you need a local server environment to test and debug the solution. I used [MAMP](https://www.mamp.info/en/ "MAMP") on a Mac with [Atom](https://atom.io/ "Atom") as my code editor. Given the solution passes access tokens around when communicating with Office 365, it is advised to leverage SSL/HTTPS on the web server. This documentation will not cover SSL/HTTPS setup as that will vary greatly based on the local server and configuration.

Office 365 applications are secured by Azure Active Directory, which comes as part of an Office 365 subscription. If you do not have an Office 365 Subscription or associated it with Azure AD, [join the Office 365 Developer Program and get a free 1 year subscription to Office 365](https://aka.ms/devprogramsignup), then follow the steps to [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/HowTo/setup-development-environment "Set up your Office 365 development environment") from MSDN.

## Registering the App ##
The first step in developing an application that connects to Office 365 is registering an application with Azure AD.

1. Sign-in to the [Azure Management Portal](https://manage.windowsazure.com "Azure Management Portal") using an account that has administrator access to the Azure AD Directory for Office 365.
2. Locate and select the **Active Directory** option towards the bottom of the left navigation (if you don't have a full Azure subscription, it might be your only option).
3. In the directory listing, select the directory associated with the Office 365 subscription you want to work with.
4. Next, select the **Applications** tab in the top navigation:
![Applications tab](http://i.imgur.com/nv168lw.png)
5. Click on the **ADD** button at the bottom center of the footer to launch the add application dialog:
![Add application in Azure AD](http://i.imgur.com/GbyS3u4.png)
6. Select **Add an application my organization is developing**
7. On the next screen, provide a **Name** for the application, keep the **Type** set to **Web Application and/or Web API** and click the next arrow.
8. Next, provide a **Sign-On URL** (points to where you want tokens returned https://localhost/mygroups/login.php) and **App ID URI** (any globally unique URI such as https://tenant.onmicrosoft.com/MyPHPGroupsApp). Don't worry, these values can be changed later:
![Applicationi Sign-on URL and URI](http://i.imgur.com/ZwnTyP5.png)
9. Click the check button to provision the new application in Azure AD.
10. When the application finishes provisioning, click on the **Configure** tab in the top navigation.
11. Locate the **CLIENT ID** and copy it to notepad...we will configure this as a setting in PHP later.
12. Locate the **Application is multi-tenant** field and toggle it to **YES** (this allows any organization to use the application). 
13. Locate the **Keys** section and use the drop-down to generate a new 2 year key. Please note that the key (also referred to as a password or secret) will only display after clicking the Save button in the footer. This is the only time the key can ever be displayed, so make sure to copy it to notepad so we can use it later:
![Application keys](http://i.imgur.com/ScmVcDU.png)
13.  After saving/retrieving the key, locate the **permissions to other applications** section at the bottom of the screen.
14.  Use the **Add Application** button to launch the Permissions to other applications dialog.
15.  Locate the **Office 365 unified API (preview)** application, select it, then click the check button to close the dialog:
![Permissions to other applications](http://i.imgur.com/16yCo3A.png)
16.  Back on the main configuration screen, locate the **Office 365 unified API (preview)** application you added in the **permissions to other applications** section, click on the Delegated Permissions dropdown and add permissions for **Access directory as the signed in user** and **Read all groups (preview)**:
![Permissions](http://i.imgur.com/61a6wP2.png)
17.   Click the **Save** button in the footer one last time to save the changes you made to permissions.

## Updating Settings.php ##
The solution contains a Settings.php file, which contains all the settings specific to the application. It needs to be updated with many of the values captured from the application registration process we just finished in Azure Active Directory. Specifically, values for $clientId, $password, and $redirectURI should be updated to reflect the values from your application registration in Azure AD.

1. $clientId should be set to the value from **Step 11** above
2. $password should be set to the value from **Step 13** above
3. $redirectURI should be set to the **Sign-on URL** value from **Step 8** above

Here is the complete Settings.php file: 

	<?php
    class Settings
    {
        public static $clientId = '04c16f20-845f-4307-94e8-753afe140bcd';
        public static $password = 'D0yRy92NcmNYAZK0wuvONmY90Sth4Mh8n2wpFWtJUdg=';
        public static $authority = 'https://login.microsoftonline.com/common/';
        public static $redirectURI = 'https://localhost/mygroups/Login.php';
        public static $unifiedAPIResource = 'https://graph.microsoft.com';
        public static $unifiedAPIEndpoint = 'https://graph.microsoft.com/beta/';
        public static $tokenCache = 'TOKEN_CACHE';
        public static $isAddin = 'IS_ADDIN';
        public static $apiRoot = 'API_ROOT';
    }
	?>

## Running the Stand-alone Web Application ##
There isn't anything tricky about the stand-alone web application. When it launches, it will look for a cached refresh token. If one exists, it will use it to get a new access token and query Office 365 Groups. If it doesn't have a cached token, it will redirect the user to a login page that will ultimately acquire a token and put it into cache.
## Running the Office Add-in ##
The Office add-in is a little more tricky in PHP, at least compared to building add-ins using Visual Studio which automatically deploys the XML Manifest files for add-ins into Office. Instead, we will deploy this manifest manually, either to a network share or add-in catalog site collection in SharePoint Online:

> NOTE: at the time of this publication, Office Add-ins were not yet supported in Office for Mac. If you are developing PHP on a Mac, you might need to use the add-in catalog deployment approach and test the add-in in Excel Online (ie - in a browser). 

1. Open the OfficeManifest.xml file and update the DefaultValue attribute of the SourceLocation element to the location you are running the PHP website. You MUST keep the ?addin=1 URL parameter on the updated value. This URL parameter tells the PHP application to include the office.js scripts/components:

    	<SourceLocation DefaultValue="https://localhost/mygroups/index.php?addin=1" />

2. Use exactly the same value for the DefaultValue attribute of the "<bt:Url id="Contoso.Taskpane.Url" element further down in the file.

    	<bt:Url id="Contoso.Taskpane.Url DefaultValue="https://localhost/mygroups/index.php?addin=1" />

2. Save the changes to the OfficeManifest.xml file and deploy it using the steps outlined in [Create a network shared folder catalog for task pane and content add-ins](https://msdn.microsoft.com/EN-US/library/office/fp123503.aspx "Create a network shared folder catalog for task pane and content add-ins") OR [Publish task pane and content add-in to an add-in catalog on SharePoint](https://msdn.microsoft.com/EN-US/library/office/fp123517.aspx "Publish task pane and content add-in to an add-in catalog on SharePoint").
3. Sideload the add-in as described in [Sideload your add-in](https://dev.office.com/docs/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins#sideload-your-add-in). When this is done, there will be a new tab called **Groups API** on the **Home** ribbon in Excel (not shown in the screenshot below). 
4. Click the **Open** button in the **Groups API** tab. When the task pane opens, it should look like the screenshot below:

![Add-in successful load](http://i.imgur.com/PFNfSIJ.png)

> NOTE: Office add-ins must register any domain they will display in AppDomains section of the add-in manifest. This application has registered login.microsoftonline.com, which is the normal login page for Office 365. If you use a federated login, the add-in will not function as the federated login screen will get kicked out into a popup. It is possible to build a functional add-in with federation/popups, but was not the focus of this app.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
