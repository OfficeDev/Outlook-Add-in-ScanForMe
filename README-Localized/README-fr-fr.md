# <a name="outlook-add-in-a-mail-add-in-for-a-read-scenario-that-checks-whether-the-user-is-mentioned-on-the-to-line-cc-line-or-body-of-an-email"></a>Complément Outlook : complément de messagerie pour un scénario de lecture qui vérifie si l’utilisateur est mentionné sur la ligne À, sur la ligne Cc ou dans le corps d’un message électronique.

**Table des matières**

* [Résumé](#summary)
* [Conditions préalables](#prerequisites)
* [Composants clés de l’exemple](#components)
* [Description du code](#codedescription)
* [Création et débogage](#build)
* [Résolution des problèmes](#troubleshooting)
* [Questions et commentaires](#questions)
* [Contribution](#contribute)
* [Ressources supplémentaires](#additional-resources)

<a name="summary"></a>
##<a name="summary"></a>Résumé

Dans cet exemple, nous allons vous montrer comment utiliser l’[interface API JavaScript pour Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) afin de créer un complément Outlook qui analyse le corps d’un message électronique pour rechercher des liens hypertextes. Voici une image du scénario en question.

 ![](../readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##<a name="prerequisites"></a>Conditions préalables
Cet exemple nécessite les éléments suivants :  

  - Visual Studio 2015.  
  - Un ordinateur exécutant Exchange 2013 avec au moins un compte de messagerie ou un compte Office 365. Vous pouvez souscrire un [abonnement de développeur Office 365](https://aka.ms/devprogramsignup) et obtenir un compte Office 365 par son intermédiaire.
  - Internet Explorer 9 ou une version ultérieure, qui doit être installé mais pas nécessairement défini comme le navigateur par défaut. Pour prendre en charge des Compléments Office, le client Office qui sert d’hôte utilise des composants de navigateur faisant partie d’Internet Explorer 9 ou d’une version ultérieure.
  - L’un des éléments suivants en tant que navigateur par défaut : Edge, Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 ou une version ultérieure de l’un de ces navigateurs.
  - Être familiarisé avec les services web et de programmation JavaScript.

<a name="components"></a>
##<a name="key-components"></a>Composants clés

Cette solution a été créée dans [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). Elle consiste en deux projets : ScanForMe et ScanForMeWeb. Voici la liste des fichiers clés compris dans ces projets. 
#### <a name="scanforme-project"></a>Projet ScanForMe

* [```ScanForMe.xml```](/ScanForMe/ScanForMeManifest/ScanForMe.xml) [Fichier manifeste](https://dev.office.com/docs/add-ins/outlook/manifests/manifests) pour le complément Outlook.

#### <a name="scanformeweb-project"></a>Projet ScanForMeWeb

* [```ItemRead.html```](/ScanForMeWeb/ItemRead.html) Interface utilisateur HTML pour le complément Outlook.
* [```ItemRead.js```](/ScanForMeWeb/ItemRead.js) Code JavaScript utilisé par Home.html pour interagir avec Word à l’aide de JavaScript pour l’API Office. 


<a name="codedescription"></a>
##<a name="description-of-the-code"></a>Description du code

La logique de base de cet exemple se trouve dans le fichier [```ItemRead.js```](/ScanForMeWeb/ItemRead.js) dans le projet ScanForMeWeb. 

Une fois que le complément est initialisé, les propriétés `item.to` et `item.cc` sont analysées pour détecter l’adresse de messagerie de l’utilisateur. L’adresse de messagerie de l’utilisateur est récupérée à partir de la propriété [```Office.context.mailbox.userProfile```](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.userProfile). Si l’utilisateur a été trouvé sur les lignes À ou Cc de ce message, cela signifie qu’il est enregistré dans l’interface utilisateur du complément. 

La méthode [```getAsync()```](http://dev.office.com/reference/add-ins/outlook/Body) de l’objet corps est ensuite utilisée pour récupérer le corps du message électronique au format texte. Lorsque cette opération asynchrone est terminée, notre fonction de rappel en ligne est invoquée. Cette fonction utilise une expression régulière pour analyser le texte du corps du courrier électronique afin de détecter les occurrences du prénom de l’utilisateur. Si au moins une occurrence est détectée, l’interface utilisateur du complément prend note du fait que l’utilisateur a été mentionné dans le corps du message électronique. 

>Remarque : Pour obtenir un exemple de l’utilisation de getAsync pour récupérer le corps d’un message électronique au format HTML, voir l’exemple [Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer). 


<a name="build"></a>
##<a name="build-and-debug"></a>Création et débogage
1. Ouvrez le fichier [```ScanForMe.sln```](ScanForMe.sln) dans Visual Studio.
2. Appuyez sur F5 pour créer et déployer l’exemple de complément. 
3. Lors du démarrage d’Outlook, sélectionnez un message électronique dans votre boîte de réception.
4. Lancez le complément en le sélectionnant dans la barre d’application du complément.

<a name="questions"></a>
## <a name="questions-and-comments"></a>Questions et commentaires

- Si vous rencontrez des difficultés pour exécuter cet exemple, [consignez un problème](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues).
- Si vous avez des questions générales sur le développement de compléments Office, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Posez vos questions avec la balise [office-addins].


<a name="contribute"></a>
## <a name="contributing"></a>Contribution ##
Nous vous invitons à contribuer à nos exemples. Pour obtenir des instructions sur la façon de procéder, consultez notre [guide de contribution](./Contributing.md).

Ce projet a adopté le [code de conduite Microsoft Open Source](https://opensource.microsoft.com/codeofconduct/). Pour plus d’informations, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.


<a name="additional-resources"></a>
## <a name="additional-resources"></a>Ressources supplémentaires ##

- [Autres exemples de compléments](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Compléments Office](https://dev.office.com/reference/add-ins)
- [Structure d’un complément](https://dev.office.com/docs/add-ins/overview/office-add-ins#StartBuildingApps_AnatomyofApp)
- [Création d’un complément Office avec Visual Studio](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio)


## <a name="copyright"></a>Copyright
Copyright (c) 2015 Microsoft. Tous droits réservés.
