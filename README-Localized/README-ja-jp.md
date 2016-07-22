# Outlook アドイン:読み取りシナリオ用のメール アドイン。ユーザーが電子メールの宛先行、CC 行、または本文のいずれに記載されているかを確認します。

**目次**

* [概要](#summary)
* [前提条件](#prerequisites)
* [サンプルの主要なコンポーネント](#components)
* [コードの説明](#codedescription)
* [ビルドとデバッグ](#build)
* [トラブルシューティング](#troubleshooting)
* [質問とコメント](#questions)
* [投稿](#contribute)
* [その他の技術情報](#additional-resources)

<a name="summary"></a>
##まとめ

これから示すこのサンプルでは、[JavaScript API for Office](https://msdn.microsoft.com/ja-jp/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15) を使用して、ハイパーリンクを参照する電子メールの本文を解析する Outlook アドインの作成方法を示します。次に、このサンプルのシナリオの図を示します。

 ![](../https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##前提条件
このサンプルを実行するには次のものが必要です。  

  - Visual Studio 2013 更新プログラム 5 または Visual Studio 2015。  
  - 少なくとも 1 つの電子メール アカウントまたは Office 365 アカウントがある Exchange 2013 を実行するコンピューター。[Office 365 Developer サブスクリプション](http://aka.ms/ro9c62)にサインアップし、そこから Office 365 アカウントを取得することができます。
  - Internet Explorer 9 以降。インストールが必要ですが、既定のブラウザーである必要はありません。Office アドインをサポートするために、ホストとして機能する Office クライアントは Internet Explorer 9 以上を構成しているブラウザー コンポーネントを使用します。
  - 次のいずれかの既定のブラウザー: Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13、またはこれらのブラウザーのそれ以降のバージョン。
  - JavaScript プログラミングと Web サービスに精通していること。

<a name="components"></a>
##主要なコンポーネント

このソリューションは、[Visual Studio](https://msdn.microsoft.com/ja-jp/library/office/fp179827.aspx#Tools_CreatingWithVS) で作成されました。ScanForMe と ScanForMeWeb の 2 つのプロジェクトで構成されています。以下に、それらのプロジェクト内のキー ファイルの一覧を示します。 
#### ScanForMe プロジェクト

* [```ScanForMe.xml```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMe/ScanForMeManifest/ScanForMe.xml) Word アドインの[マニフェスト ファイル](https://msdn.microsoft.com/ja-jp/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)。

#### ScanForMeWeb プロジェクト

* [```Home.html```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.html) Word アドインの HTML ユーザー インターフェイス。
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.js) Office API の JavaScript を使用して Word と対話するために Home.html によって使用される JavaScript コード。 


<a name="codedescription"></a>
##コードの説明

このサンプルのコア ロジックは、ScanForMeWeb プロジェクトの [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.js) ファイルです。 

アドインを初期化すると、ユーザーの電子メール アドレスが存在するかどうか、`item.to` と `item.cc` プロパティをスキャンします。ユーザーの電子メール アドレスは、[```Office.context.mailbox.userProfile```](https://msdn.microsoft.com/ja-jp/library/office/fp160976.aspx) プロパティから取得します。ユーザーがこの電子メールの宛先行または CC 行で検出された場合、それらはアドインの UI に登録されます。 

body オブジェクトの [```getAsync()```](https://msdn.microsoft.com/ja-jp/library/office/mt269089.aspx) メソッドは、テキスト形式の電子メールの本文を取得するために使用されます。この非同期操作が完了すると、インライン コールバック関数が呼び出されます。この関数は正規表現を使用して、ユーザーの名が使用されていないか電子メール本文のテキストをスキャンします。1 回以上使用されていることが検出された場合、アドインの UI に、ユーザーの名前が電子メールの本文に記載されていることをが示されます。 

>注意:HTML 形式の電子メールの本文を取得する getAsync を使用する例については、[Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer) のサンプルを参照してください。 


<a name="build"></a>
##ビルドとデバッグ
1. [```ScanForMe.sln```](ScanForMe.sln) ファイルを Visual Studio で開きます。
2. F5 キーを押して、サンプル アドインをビルドおよび展開します。 
3. Outlook が起動したら、受信トレイから電子メールを選択します。
4. アドイン アプリ バーからアドインを選択して、起動します。

 - スクリーンショットが必要


5. [次に起こることの説明]


<a name="troubleshooting"></a>
## トラブルシューティング

- アドインが作業ウィンドウに表示されない場合、**[挿入] > [個人用アドイン] > [自分の名前をスキャン]** を選択します。

<a name="questions"></a>
## 質問とコメント

- このサンプルの実行について問題がある場合は、[問題をログに記録](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues)してください。
- Office アドイン開発全般の質問については、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office-addins)」に投稿してください。質問またはコメントには、必ず [office-addins] のタグを付けてください。


<a name="contribute"></a>
## 投稿 ##
当社のサンプルに是非投稿してください。投稿方法のガイドラインについては、[投稿ガイド](./Contributing.md)を参照してください。

このプロジェクトでは、[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[規範に関する FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。または、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までにお問い合わせください。


<a name="additional-resources"></a>
## その他の技術情報 ##

- <a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=-Add-in">その他のアドインのサンプル</a>
- [Office アドイン](http://msdn.microsoft.com/ja-jp/library/office/jj220060.aspx)
- [アドインの構造](https://msdn.microsoft.com/ja-jp/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Visual Studio で Office アドインを作成する](https://msdn.microsoft.com/ja-jp/library/office/fp179827.aspx#Tools_CreatingWithVS)


## 著作権
Copyright (c) 2015 Microsoft.All rights reserved.

