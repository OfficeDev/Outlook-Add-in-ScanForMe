# Suplemento do Outlook: Um suplemento de email para um cenário de leitura que verifica se o usuário está sendo mencionado na linha Para, na linha Cc ou no corpo de um email.

**Sumário**

* [Resumo](#summary)
* [Pré-requisitos](#prerequisites)
* [Componentes principais do exemplo](#components)
* [Descrição do código](#codedescription)
* [Criar e depurar](#build)
* [Solução de problemas](#troubleshooting)
* [Perguntas e comentários](#questions)
* [Colaboração](#contribute)
* [Recursos adicionais](#additional-resources)

<a name="summary"></a>
##Resumo

Neste exemplo mostraremos como usar a [API JavaScript para Office](https://msdn.microsoft.com/pt-br/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) para criar um suplemento do Outlook que analisa o corpo de um email procurando hiperlinks. Veja uma imagem do cenário em questão.

 ![](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##Pré-requisitos
Esse exemplo requer o seguinte:  

  - Visual Studio 2013 com Atualização 5 ou Visual Studio 2015.  
  - Um computador executando o Exchange 2013 com pelo menos uma conta de email ou uma conta do Office 365. Você pode se inscrever para [uma assinatura do Office 365 Developer](http://aka.ms/ro9c62) e, através da assinatura, obter uma conta do Office 365.
  - Internet Explorer 9 ou posterior, que deve estar instalado, mas não precisa ser o navegador padrão. Para oferecer suporte aos Suplementos do Office, o cliente do Office que atua como host usa os componentes do navegador que fazem parte do Internet Explorer 9 ou posterior.
  - Um dos seguintes como o navegador padrão: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 ou uma versão mais recente de um desses navegadores.
  - Familiaridade com programação em JavaScript e serviços Web.

<a name="components"></a>
##Componentes principais

Essa solução foi criada no [Visual Studio](https://msdn.microsoft.com/pt-br/library/office/fp179827.aspx#Tools_CreatingWithVS). Ela é formada por dois projetos - ScanForMe e ScanForMeWeb. Veja uma lista dos principais arquivos dentro desses projetos. 
#### Projeto ScanForMe

* [```QbAdd inDotNet.xml```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMe/ScanForMeManifest/ScanForMe.xml) o [ arquivo de manifesto](https://msdn.microsoft.com/pt-br/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp) do suplemento do Word.

#### Projeto ScanForMeWeb

* [```Home.HTML```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.html) interface do usuário HTML para o suplemento do Word.
* [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.js) o código JavaScript usado por Home.html para interagir com o Word usando o a API JavaScript para Office. 


<a name="codedescription"></a>
##Descrição do código

A lógica principal deste exemplo está no arquivo [```Home.js```](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/blob/master/ScanForMeWeb/AppRead/Home/Home.js) no projeto ScanForMeWeb. 

Depois que o suplemento for inicializado, as propriedades `item.to` e `item.cc` serão verificadas em busca do endereço de email do usuário. O endereço de email do usuário é recuperado na propriedade [```Office.context.mailbox.userProfile```](https://msdn.microsoft.com/pt-br/library/office/fp160976.aspx). Se o usuário foi encontrado nas linhas Para ou Cc deste email, esse fato é registrado na interface do usuário do suplemento. 

O método [```getAsync()```](https://msdn.microsoft.com/pt-br/library/office/mt269089.aspx) do objeto Corpo é então usado para recuperar o corpo do email no formato Texto. Quando essa operação assíncrona for concluída, a função de retorno de chamada em linha é chamada. Esta função usa uma expressão regular para verificar o texto do corpo do email para verificar as ocorrências do nome do usuário. Se uma ou mais ocorrências forem encontradas, a interface do usuário do suplemento anota as ocorrências nas quais o usuário foi mencionado no corpo do email. 

>Observação: Para obter um exemplo de uso do getAsync para recuperar o corpo de um email no formato HTML, confira o exemplo [Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer). 


<a name="build"></a>
##Criar e depurar
1. Abra o arquivo [```ScanForMe.sln```](ScanForMe.sln) no Visual Studio.
2. Pressione F5 para compilar e implantar o suplemento de exemplo 
3. Quando o Outlook iniciar, escolha um email de sua caixa de entrada
4. Inicie o suplemento selecionando-o na barra de aplicativo do suplemento

 - requer captura de tela


5. [Descreve o que acontece a seguir]


<a name="troubleshooting"></a>
## Solução de problemas

- Se o suplemento não for exibido no painel de tarefas, escolha **Inserir > Meus Suplementos > Digitalizar para Mim**.

<a name="questions"></a>
## Perguntas e comentários

- Se você tiver problemas para executar este exemplo, [relate um problema](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues).
- As perguntas sobre o desenvolvimento de Suplementos do Office em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Não deixe de marcar as perguntas ou comentários com [office-addins].


<a name="contribute"></a>
## Colaboração ##
Recomendamos que você contribua para nossos exemplos. Para obter diretrizes sobre como proceder, confira nosso [guia de contribuição](./Contributing.md)

Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira as [Perguntas frequentes do Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.


<a name="additional-resources"></a>
## Recursos adicionais ##

- <a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=-Add-in">Mais exemplos de Suplementos</a>
- [Suplementos do Office](http://msdn.microsoft.com/pt-br/library/office/jj220060.aspx)
- [Anatomia de um Suplemento](https://msdn.microsoft.com/pt-br/library/office/jj220082.aspx#StartBuildingApps_AnatomyofApp)
- [Criação de um Suplemento do Office com o Visual Studio](https://msdn.microsoft.com/pt-br/library/office/fp179827.aspx#Tools_CreatingWithVS)


## Direitos autorais
Copyright © 2015 Microsoft. Todos os direitos reservados.

