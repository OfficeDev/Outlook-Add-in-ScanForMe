# Outlook 增益集：讀取案例的郵件增益集，它會檢查使用者是否在 [收件人] 行、[副本] 行或電子郵件的本文提及。

**目錄**

* [摘要](#summary)
* [必要條件](#prerequisites)
* [範例的主要元件](#components)
* [程式碼的描述](#codedescription)
* [建置和偵錯](#build)
* [疑難排解](#troubleshooting)
* [問題和建議](#questions)
* [參與](#contribute)
* [其他資源](#additional-resources)

<a name="summary"></a>
##摘要

在這個範例中，為您示範如何使用 [JavaScript API for Office](https://msdn.microsoft.com/zh-tw/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) 以建立 Outlook 增益集，剖析尋找超連結的電子郵件的本文。下列是有問題的案例圖片。

 ![](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##必要條件
此範例需要下列項目：  

  - Visual Studio 2013 (具有更新 5) 或 Visual Studio 2015。  
  - 執行 Exchange 2013 的電腦且具有至少一個電子郵件帳戶，或 Office 365 帳戶。您可以註冊 [Office 365 開發人員訂用帳戶](http://aka.ms/ro9c62)，並且透過它取得 Office 365 帳戶。
  - Internet Explorer 9 或更新版本，必須先安裝，但不一定是預設瀏覽器。若要支援 Office 增益集，做為主機的 Office 用戶端會使用 Internet Explorer 9 或更新版本的瀏覽器元件。
  - 預設瀏覽器為下列其中一項︰Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13 或其中一個瀏覽器的更新版本。
  - 熟悉 JavaScript 程式設計和 Web Services。

<a name="components"></a>
##主要元件

此解決方案是在 [Visual Studio](https://msdn.microsoft.com/zh-tw/library/office/fp179827.aspx#Tools_CreatingWithVS) 中建立。它包含兩個專案 - ScanForMe 和 ScanForMeWeb。以下是這些專案中主要檔案的清單。 
#### ScanForMe 專案

* [```ScanForMe.xml```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMe/ScanForMeManifest/ScanForMe.xml) Word 增益集的[資訊清單檔案](https://msdn.microsoft.com/zh-tw/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)。

#### ScanForMeWeb 專案

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.html) Word 增益集的 HTML 使用者介面。
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.js) Home.html 用於與使用 JavaScript for Office API 的 Word 互動的 JavaScript 程式碼。 


<a name="codedescription"></a>
##程式碼的描述

這個範例的核心邏輯位於 ScanForMeWeb 專案中的 [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.js) 檔案。 

一旦初始化增益集，`item.to` 和 `item.cc` 屬性都會被掃描是否有使用者的電子郵件地址存在。使用者電子郵件地址是從 [```Office.context.mailbox.userProfile```](https://msdn.microsoft.com/zh-tw/library/office/fp160976.aspx) 屬性擷取。如果在此電子郵件上的收件人或副本行中找到使用者，該事實已已登錄在增益集的 UI。 

然後會使用本文物件的 [```GetAsync()```](https://msdn.microsoft.com/zh-tw/library/office/mt269089.aspx) 方法，擷取文字格式的電子郵件本文。當這個非同步作業完成時，會叫用我們的內嵌回呼函式。這個函式會使用規則運算式來掃描電子郵件本文文字中，使用者名字的出現次數。如果找到一或多個出現次數，增益集的 UI 會注意到使用者已在電子郵件本文中提及。 

>注意︰如需使用 getAsync 來擷取 HTML 格式的電子郵件本文的範例，請參閱 [Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer) 範例。 


<a name="build"></a>
##建置和偵錯
1. 在 Visual Studio 中開啟 [```ScanForMe.sln```](ScanForMe.sln) 檔案。
2. 按 F5 建置及部署範例增益集 
3. 當 Outlook 啟動時，從您的收件匣中選取一封電子郵件
4. 藉由從增益集應用程式列中選取它以啟動增益集

 - 需要螢幕擷取畫面


5. [說明接下來發生的動作]


<a name="troubleshooting"></a>
## 疑難排解

- 如果增益集未出現在工作窗格中，選擇 [插入 > 我的增益集 > 為我掃描]<e />。

<a name="questions"></a>
## 問題和建議

- 如果執行此範例有任何問題，請[開立問題](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues)。
- 請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins) 提出有關 Office 增益集開發的一般問題。務必以 [office-addins] 標記您的問題或意見。


<a name="contribute"></a>
## 參與 ##
我們鼓勵您參與我們的範例。如需如何繼續的指引，請參閱我們的[參與指南](./Contributing.md)

此專案已採用 [Microsoft 開放原始碼執行](https://opensource.microsoft.com/codeofconduct/)。如需詳細資訊，請參閱[程式碼執行常見問題集](https://opensource.microsoft.com/codeofconduct/faq/)，如果有其他問題或意見，請連絡 [opencode@microsoft.com](mailto:opencode@microsoft.com)。


<a name="additional-resources"></a>
## 其他資源 ##

- <a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=-Add-in">更多增益集範例</a>
- [Office 增益集](http://msdn.microsoft.com/zh-tw/library/office/jj220060.aspx)
- [增益集的架構](https://msdn.microsoft.com/zh-tw/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [使用 Visual Studio 建立 Office 增益集](https://msdn.microsoft.com/zh-tw/library/office/fp179827.aspx#Tools_CreatingWithVS)


## 著作權
Copyright (c) 2015 Microsoft.著作權所有，並保留一切權利。

