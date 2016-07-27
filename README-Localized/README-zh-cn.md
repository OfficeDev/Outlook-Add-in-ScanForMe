# Outlook 外接程序：读取方案的邮件外接程序检查用户是否在“收件人”行、“抄送”行或电子邮件的正文中被提及。

**目录**

* [摘要](#summary)
* [先决条件](#prerequisites)
* [示例的主要组件](#components)
* [代码说明](#codedescription)
* [构建和调试](#build)
* [疑难解答](#troubleshooting)
* [问题和意见](#questions)
* [供稿](#contribute)
* [其他资源](#additional-resources)

<a name="summary"></a>
##摘要

在此示例中，我们将向你介绍如何使用[适用于 Office 的 JavaScript API](https://msdn.microsoft.com/zh-cn/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15) 来创建解析查找超链接的电子邮件正文的 Outlook 外接程序。以下是要讨论的方案的图片。

 ![](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##先决条件
此示例要求如下：  

  - Visual Studio 2013 Update 5 或 Visual Studio 2015。  
  - 运行至少具有一个电子邮件帐户或 Office 365 帐户的 Exchange 2013 的计算机。你可以注册 [Office 365 开发人员订阅](http://aka.ms/ro9c62)，并通过该订阅获得 Office 365 帐户。
  - 必须安装 Internet Explorer 9 或更高版本，但不一定作为默认浏览器。为了支持 Office 外接程序，作为主机的 Office 客户端使用属于 Internet Explorer 9 或更高版本的一部分的浏览器组件。
  - 使用下列浏览器之一作为默认浏览器：Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13 或这些浏览器的更高版本。
  - 熟悉 JavaScript 编程和 Web 服务。

<a name="components"></a>
##主要组件

此解决方案在 [Visual Studio](https://msdn.microsoft.com/zh-cn/library/office/fp179827.aspx#Tools_CreatingWithVS) 中创建。它包含两个项目 - ScanForMe 和 ScanForMeWeb。以下是这些项目中关键文件的列表。 
#### ScanForMe 项目

* [```ScanForMe.xml```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMe/ScanForMeManifest/ScanForMe.xml) Word 外接程序的[清单文件](https://msdn.microsoft.com/zh-cn/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)。

#### ScanForMeWeb 项目

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.html) Word 外接程序的 HTML 用户界面。
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.js)由 Home.html 使用的与使用适用于 Office 的 JavaScript API 的 Word 进行交互的 JavaScript 代码。 


<a name="codedescription"></a>
##代码说明

此示例的核心逻辑在于 ScanForMeWeb 项目中的 [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.js) 文件。 

外接程序进行初始化后，扫描`item.to` 和 `item.cc` 属性以检查用户的电子邮件地址是否存在。从 [```Office.context.mailbox.userProfile```](https://msdn.microsoft.com/zh-cn/library/office/fp160976.aspx) 属性检索用户电子邮件地址。如果在此电子邮件的“收件人”行或“抄送”行找到用户，则该事实将注册到外接程序的 UI 上。 

Body 对象的 [```getAsync()```](https://msdn.microsoft.com/zh-cn/library/office/mt269089.aspx) 方法随后被用于检索文本格式的电子邮件正文。此异步操作完成时，调用内联回调函数。此函数使用一个正则表达式对电子邮件正文的文本进行扫描，以查看用户名字的出现情况。如果找到一个或多个，则外接程序的 UI 将注释“用户在电子邮件正文中被提及”。 

> 注意：有关使用 getAsync 检索 HTML 格式的电子邮件正文的示例，请参阅 [Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer) 示例。 


<a name="build"></a>
##构建和调试
1. 打开 Visual Studio 中的 [```ScanForMe.sln```](ScanForMe.sln) 文件。
2. 按 F5 构建并部署示例外接程序 
3. 当 Outlook 启动时，从收件箱中选择一封电子邮件
4. 通过从外接程序应用栏选择外接程序将其启动。

 - 需要屏幕截图


5. [描述接下来发生的事项]


<a name="troubleshooting"></a>
## 疑难解答

- 如果任务窗格中未显示外接程序，请选择“**插入 > 我的外接程序 > 为我扫描**”。

<a name="questions"></a>
## 问题和意见

- 如果你在运行此示例时遇到任何问题，请[记录问题](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues)。
- 与 Office 外接程序开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)。确保你的问题或意见使用了 [office-addins] 标记。


<a name="contribute"></a>
## 参与 ##
我们鼓励你参与我们的示例。有关如何继续的指南，请参阅我们的[参与指南](./Contributing.md)

此项目采用 [Microsoft 开源行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅 [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/)（行为准则常见问题解答），有任何其他问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。


<a name="additional-resources"></a>
## 其他资源 ##

- <a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=-Add-in">更多外接程序示例</a>
- [Office 外接程序](http://msdn.microsoft.com/zh-cn/library/office/jj220060.aspx)
- [外接程序剖析](https://msdn.microsoft.com/zh-cn/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [使用 Visual Studio 创建 Office 外接程序](https://msdn.microsoft.com/zh-cn/library/office/fp179827.aspx#Tools_CreatingWithVS)


## 版权
版权所有 (c) 2015 Microsoft。保留所有权利。

