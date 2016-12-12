# Outlook add-in: A mail add-in for a read scenario that checks whether the user is mentioned on the To line, cc line or body of an email.

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Contributing](#contribute)
* [Additional resources](#additional-resources)

<a name="summary"></a>
##Summary

In this sample we show you how to use the [JavaScript API for Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) to create an Outlook add-in that parses the body of an email looking for hyperlinks. The following is a  picture of the scenario in question.

 ![](/readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##Prerequisites
This sample requires the following:  

  - Visual Studio 2015.  
  - A computer running Exchange 2013 with at least one email account, or an Office 365 account. You can sign up for [an Office 365 Developer subscription](https://aka.ms/devprogramsignup) and get an Office 365 account through it.
  - Internet Explorer 9 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.
  - One of the following as the default browser: Edge, Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.
  - Familiarity with JavaScript programming and web services.

<a name="components"></a>
##Key components

This solution was created in [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). It consists of two projects - ScanForMe and ScanForMeWeb. Here's a list of the key files within those projects. 
#### ScanForMe project

* [```ScanForMe.xml```](/ScanForMe/ScanForMeManifest/ScanForMe.xml) The [manifest file](https://dev.office.com/docs/add-ins/outlook/manifests/manifests) for the Outlook add-in.

#### ScanForMeWeb project

* [```ItemRead.html```](/ScanForMeWeb/ItemRead.html) The HTML user interface for the Outlook add-in.
* [```ItemRead.js```](/ScanForMeWeb/ItemRead.js) The JavaScript code used by Home.html to interact with Word using the JavaScript for Office API. 


<a name="codedescription"></a>
##Description of the code

The core logic of this sample is in the [```ItemRead.js```](/ScanForMeWeb/ItemRead.js)  file in the ScanForMeWeb project. 

Once the add-in is initialized, the `item.to` and `item.cc` properties are scanned for the presence of the user's email address. The user email address is retrieved from the [```Office.context.mailbox.userProfile```](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.userProfile) property. If the user was found on the to or cc lines of this email, that fact is registered on the UI of the add-in. 

The [```getAsync()```](http://dev.office.com/reference/add-ins/outlook/Body) method of the Body object is then used to retrieve the body of the email in Text format. When this asynchronous operation is completed, our inline callback function is invoked. This function uses a regular expression to scan the text of the email body for occurrences of the user's first name. If one or more occurrences are found, the UI of the add-in notes that the user was mentioned in the body of the email. 

>Note: For an example of using getAsync to retrieve the body of an email in HTML format, see the [Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer) sample. 


<a name="build"></a>
##Build and debug
1. Open the [```ScanForMe.sln```](ScanForMe.sln) file in Visual Studio.
2. Press F5 to build and deploy the sample add-in 
3. When Outlook launches, select an email from your inbox
4. Launch the add-in by selecting it from the add-in app bar

<a name="questions"></a>
## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues).
- Questions about Office Add-ins development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="contribute"></a>
## Contributing ##
We encourage you to contribute to our samples. For guidelines on how to proceed, see our [contribution guide](./Contributing.md)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.


<a name="additional-resources"></a>
## Additional resources ##

- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Office Add-ins](https://dev.office.com/reference/add-ins)
- [Anatomy of an Add-in](https://dev.office.com/docs/add-ins/overview/office-add-ins#StartBuildingApps_AnatomyofApp)
- [Creating an Office add-in with Visual Studio](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio)


## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
