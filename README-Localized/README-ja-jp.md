---
page_type: sample
products:
- office
- office-outlook
- office-365
languages:
- javascript
description: "JavaScript API for Office を使用して、メールの本文を解析してハイパーリンクを検索する Outlook アドインを作成する方法。"
urlFragment: outlook-scan-sample
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: "8/29/2015 7:47:24 PM"
---

# Outlook アドイン:メールの宛先行、CC 行、または本文のいずれかでユーザーが言及されているかを確認するための、読み取りシナリオ用のメール アドイン。

**目次**

* [概要](#summary)
* [前提条件](#prerequisites)
* [サンプルの主要なコンポーネント](#components)
* [コードの説明](#codedescription)
* [ビルドとデバッグ](#build)
* [トラブルシューティング](#troubleshooting)
* [質問とコメント](#questions)
* [投稿](#contribute)
* [その他のリソース](#additional-resources)

<a name="summary"></a>
##Summary

このサンプルでは、[JavaScript API for Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) を使用して、メールの本文を解析してハイパーリンクを検索する Outlook アドインを作成する方法を示します。このシナリオの画像を次に示します。

 ![](/readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##前提条件
このサンプルを実行するには次のものが必要です。  

  - Visual Studio 2015。  
  - 少なくとも 1 つの電子メール アカウントまたは Office 365 アカウントがある Exchange 2013 を実行するコンピューター。[Office 365 Developer サブスクリプション](https://aka.ms/devprogramsignup)にサインアップし、これを使用して Office 365 アカウントを取得することができます。
  - Internet Explorer 9 以降 (既定のブラウザーである必要はありませんが、インストールされている必要があります)。Office アドインをサポートできるよう、ホストとして動作する Office のクライアントでは、Internet Explorer 9 以降に組み込まれているブラウザー コンポーネントが使用されています。
  - 既定のブラウザーとして次のいずれかのブラウザーが必要です:Edge、Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13、これらのブラウザーのいずれかの最新バージョン。
  - JavaScript プログラミングと Web サービスに精通していること。

<a name="components"></a>
##主要なコンポーネント

このソリューションは、[Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS) で作成されました。ScanForMe と ScanForMeWeb という 2 つのプロジェクトにより構成されています。以下に、それらのプロジェクト内のキー ファイルの一覧を示します。 
#### ScanForMe プロジェクト

* [```ScanForMe.xml```](/ScanForMe/ScanForMeManifest/ScanForMe.xml) Outlook アドインの[マニフェスト ファイル](https://dev.office.com/docs/add-ins/outlook/manifests/manifests)。

#### ScanForMeWeb プロジェクト

* [```ItemRead.html```](/ScanForMeWeb/ItemRead.html) Outlook アドインの HTML ユーザー インターフェイス。
* [```ItemRead.js```](/ScanForMeWeb/ItemRead.js) Office API の JavaScript を使用して Word と対話するために Home.html によって使用される JavaScript コード。 


<a name="codedescription"></a>
##コードの説明

このサンプルのコア ロジックは、ScanForMeWeb プロジェクトの [```ItemRead.js```](/ScanForMeWeb/ItemRead.js) ファイルに含まれています。 

アドインが初期化されると、ユーザーのメール アドレスが含まれているかどうかについて、`item.to` プロパティと `item.cc` プロパティがスキャンされます。ユーザーのメール アドレスは [```Office.context.mailbox.userProfile```](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.userProfile) プロパティから取得されます。ユーザーがこのメールの宛先行または CC 行で検出された場合、このことがアドインの UI に登録されます。 

Body オブジェクトの [```getAsync()```](http://dev.office.com/reference/add-ins/outlook/Body) メソッドを使用して、テキスト形式のメールの本文が取得されます。この非同期操作が完了すると、インライン コールバック関数が呼び出されます。この関数は正規表現を使用して、ユーザーの名が使用されていないか電子メール本文のテキストをスキャンします。1 回以上使用されていることが検出された場合、アドインの UI に、ユーザーの名前が電子メールの本文に記載されていることが示されます。 

>注:HTML 形式のメールの本文を取得する getAsync を使用する例については、[Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer) のサンプルを参照してください。 


<a name="build"></a>
##ビルドとデバッグ
1.Visual Studio で [```ScanForMe.sln```](ScanForMe.sln) を開きます。
2.F5 キーを押して、サンプル アドインをビルドし、展開します。
3.Outlook が起動したら、受信トレイからメールを 1 通選択します。
4.アドイン アプリ バーからアドインを選択して、起動します。

<a name="questions"></a>
## 質問とコメント

- このサンプルの実行で問題が発生した場合は、[問題を報告](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues)してください。
- Office アドイン開発に関する全般的な質問は、「[Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)」に投稿してください。質問やコメントには、必ず [office-addins] のタグを付けてください。


<a name="contribute"></a>
## 投稿 ##
Microsoft のサンプルに是非投稿してください。投稿方法のガイドラインについては、[投稿ガイド](./Contributing.md)を参照してください。

このプロジェクトでは、[Microsoft Open Source Code of Conduct (Microsoft オープン ソース倫理規定)](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[Code of Conduct の FAQ (倫理規定の FAQ)](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。


<a name="additional-resources"></a>
## その他のリソース ##

- [その他のアドイン サンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Office アドイン](https://dev.office.com/reference/add-ins)
- [アドインの構造](https://dev.office.com/docs/add-ins/overview/office-add-ins#StartBuildingApps_AnatomyofApp)
- [Visual Studio で Office アドインを作成する](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio)


## 著作権
Copyright (c) 2015 Microsoft.All rights reserved.
