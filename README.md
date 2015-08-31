# Outlook add-in: A mail add-in for a read scenario which checks whether the user is mentioned on the To line, cc line or body of an email.

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="summary"></a>
##Summary

In this sample we show you how to use the [JavaScript API for Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) to create an Outlook add-in that parses the body of an email looking for hyperlinks. The following is a  picture of the scenario in question.

 - screenshot needed

<a name="prerequisites"></a>
##Prerequisites
This sample requires the following:  

  - Visual Studio 2013 with Update 5 or Visual Studio 2015.  
  - A computer running Exchange 2013 with at least one email account, or an Office 365 account. You can sign up for [an Office 365 Developer subscription](http://aka.ms/ro9c62) and get an Office 365 account through it.
  - Internet Explorer 9 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.
  - One of the following as the default browser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.
  - Familiarity with JavaScript programming and web services.

<a name="components"></a>
##Key components

This solution was created in [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). It consists of two projects - LinkRevealer and LinkRevealerWeb. Here's a list of the key files within those projects. 
#### ActionDetector project

* [```ActionDetector.xml```](https://github.com/OfficeDev/Outlook-Add-in-ActionDetector/blob/master/ActionDetector/ActionDetectorManifest/ActionDetector.xml) The [manifest file](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp) for the Word add-in.

#### ActionDetectorWeb project

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-ActionDetector/blob/master/ActionDetectorWeb/AppRead/Home/Home.html) The HTML user interface for the Word add-in.
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ActionDetector/blob/master/ActionDetectorWeb/AppRead/Home/Home.js) The JavaScript code used by Home.html to interact with Word using the JavaScript for Office API. 


<a name="codedescription"></a>
##Description of the code

<a name="build"></a>
##Build and debug
1. Open the [```ActionDetector.sln```](ActionDetector.sln) file in Visual Studio.
2. Press F5 to build and deploy the sample add-in 
3. When Outlook launches, select an email from your inbox
4. Launch the add-in by selecting it from the add-in app bar

 - need screenshot


5. [Describe what happens next]


<a name="troubleshooting"></a>
## Troubleshooting

- If the add-in does not appear in the task pane, Choose **Insert > My Add-ins >  Action Detector**.

<a name="questions"></a>
## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Outlook-Add-in-ActionDetector/issues).
- Questions about Office Add-ins development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Office Add-ins](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Anatomy of an Add-in](https://msdn.microsoft.com/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Creating an Office add-in with Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS)


## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
