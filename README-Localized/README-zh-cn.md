# <a name="outlook-add-in-a-mail-add-in-for-a-read-scenario-that-checks-whether-the-user-is-mentioned-on-the-to-line-cc-line-or-body-of-an-email"></a>Outlook 外接程序：读取方案的邮件外接程序检查用户是否在“收件人”行、“抄送”行或电子邮件的正文中被提及。

**目录**

* [摘要](#summary)
* [先决条件](#prerequisites)
* [示例主要组件](#components)
* [代码说明](#codedescription)
* [构建和调试](#build)
* [疑难解答](#troubleshooting)
* [问题和意见](#questions)
* [参与](#contribute)
* [其他资源](#additional-resources)

<a name="summary"></a>
##<a name="summary"></a>摘要

本示例介绍如何使用 [适用于 Office 的 JavaScript API](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) 来创建可分析查找超链接的电子邮件正文的 Outlook 外接程序。以下是要讨论的方案的图片。

 ![](../readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##<a name="prerequisites"></a>先决条件
此示例要求如下：  

  - Visual Studio 2015。  
  - 运行至少具有一个电子邮件帐户或 Office 365 帐户的 Exchange 2013 的计算机。你可以注册 [Office 365 开发人员订阅](https://aka.ms/devprogramsignup)，并通过该订阅获得 Office 365 帐户。
  - 必须安装 Internet Explorer 9 或更高版本，但不一定作为默认浏览器。为了支持 Office 外接程序，作为主机的 Office 客户端使用属于 Internet Explorer 9 或更高版本的一部分的浏览器组件。
  - 使用以下任一浏览器作为默认浏览器：Edge、Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13 或这些浏览器的更高版本。
  - 熟悉 JavaScript 编程和 Web 服务。

<a name="components"></a>
##<a name="key-components"></a>主要组件

此解决方案在 [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS) 中创建。它包含两个项目 - ScanForMe 和 ScanForMeWeb。以下是这些项目中关键文件的列表。 
#### <a name="scanforme-project"></a>ScanForMe 项目

* [```ScanForMe.xml```](/ScanForMe/ScanForMeManifest/ScanForMe.xml)：Outlook 外接程序的[清单文件](https://dev.office.com/docs/add-ins/outlook/manifests/manifests)。

#### <a name="scanformeweb-project"></a>ScanForMeWeb 项目

* [```ItemRead.html```](/ScanForMeWeb/ItemRead.html)：Outlook 外接程序的 HTML 用户界面。
* [```ItemRead.js```](/ScanForMeWeb/ItemRead.js)：Home.html 使用的 JavaScript 代码，以使用适用于 Office 的 JavaScript API 与 Word 进行交互。 


<a name="codedescription"></a>
##<a name="description-of-the-code"></a>代码说明

此示例的核心逻辑在于 ScanForMeWeb 项目中的 [```ItemRead.js```](/ScanForMeWeb/ItemRead.js) 文件。 

外接程序进行初始化后，扫描`item.to` 和 `item.cc` 属性以检查用户的电子邮件地址是否存在。从 [```Office.context.mailbox.userProfile```](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.userProfile) 属性检索用户电子邮件地址。如果在此电子邮件的“收件人”行或“抄送”行找到用户，则该事实将注册到外接程序的 UI 上。 

Body 对象的 [```getAsync()```](http://dev.office.com/reference/add-ins/outlook/Body) 方法稍后用于检索文本格式的电子邮件正文。在此异步操作完成时，我们的内联回叫函数会得到调用。此函数使用正则表达式对电子邮件正文的文本进行扫描，以搜索用户名字。如果找到一处或多处，外接程序的 UI 会注明电子邮件正文中提及了用户。 

>注意：有关使用 getAsync 检索 HTML 格式电子邮件正文的示例，请参阅 [Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer) 示例。 


<a name="build"></a>
##<a name="build-and-debug"></a>生成和调试
1. 打开 Visual Studio 中的 [```ScanForMe.sln```](ScanForMe.sln) 文件。
2. 按 F5 构建并部署示例外接程序 
3. 当 Outlook 启动时，从收件箱中选择一封电子邮件
4. 通过从外接程序应用栏选择外接程序，启动外接程序

<a name="questions"></a>
## <a name="questions-and-comments"></a>问题和意见

- 如果你在运行此示例时遇到任何问题，请[记录问题](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues)。
- 与 Office 外接程序开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)。确保你的问题或意见使用了 [office-addins] 标记。


<a name="contribute"></a>
## <a name="contributing"></a>参与 ##
我们鼓励你参与我们的示例。有关如何继续的指南，请参阅我们的[参与指南](./Contributing.md)

此项目采用 [Microsoft 开源行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅 [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/)（行为准则常见问题解答），有任何其他问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。


<a name="additional-resources"></a>
## <a name="additional-resources"></a>其他资源 ##

- [更多外接程序示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Office 外接程序](https://dev.office.com/reference/add-ins)
- [外接程序解析](https://dev.office.com/docs/add-ins/overview/office-add-ins#StartBuildingApps_AnatomyofApp)
- [使用 Visual Studio 创建 Office 外接程序](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio)


## <a name="copyright"></a>版权
版权所有 (c) 2015 Microsoft。保留所有权利。