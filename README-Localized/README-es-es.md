---
page_type: sample
products:
- office
- office-outlook
- office-365
languages:
- javascript
description: "Cómo usar la API de JavaScript para Office para crear un complemento de Outlook que analiza el cuerpo de un correo electrónico en busca de hipervínculos"
urlFragment: outlook-scan-sample
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: "8/29/2015 7:47:24 PM"
---

# Complemento de Outlook: Un complemento de correo electrónico para un escenario de lectura que comprueba si se menciona al usuario en la línea Para, en la línea CC o en el cuerpo de un correo electrónico.

**Tabla de contenido**

* [Resumen](#summary)
* [Requisitos previos](#prerequisites)
* [Componentes clave del ejemplo](#components)
* [Descripción del código](#codedescription)
* [Compilar y depurar](#build)
* [Solución de problemas](#troubleshooting)
* [Preguntas y comentarios](#questions)
* [Colaboradores](#contribute)
* [Recursos adicionales](#additional-resources)

<a name="summary"></a>
##Resumen

En este ejemplo se muestra cómo usar la [API de JavaScript para Office](https://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298(v=office.15)) para crear un complemento de Outlook que analice el cuerpo de un correo electrónico para encontrar hipervínculos. La siguiente imagen corresponde al escenario en cuestión.

 ![](/readme-images/screenshot1.PNG)

<a name="prerequisites"></a>
##Requisitos
previos Este ejemplo necesita lo siguiente:  

  - Visual Studio 2015.  
  - Un equipo que ejecute Exchange 2013 y, como mínimo, una cuenta de correo electrónico o una cuenta de Office 365. Puede registrarse para [una suscripción a Office 365 Developer](https://aka.ms/devprogramsignup) y obtener una cuenta de Office 365.
  - Internet Explorer 9 o posterior, que debe estar instalado, pero no es necesario que sea el navegador predeterminado. Para admitir los complementos de Office, el cliente de Office que actúa como host utiliza componentes del explorador que forman parte de Internet Explorer 9 o posterior.
  - Uno de los siguientes como explorador predeterminado: Edge, Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 o una versión posterior de alguno de estos exploradores.
  - Familiaridad con los servicios web y la programación de JavaScript.

<a name="components"></a>
##Componentes clave

Esta solución se creó en [Visual Studio](https://msdn.microsoft.com/library/office/fp179827.aspx#Tools_CreatingWithVS). Consta de dos proyectos: ScanForMe y ScanForMeWeb. Esta es una lista de los archivos de claves en esos proyectos. 
#### Proyecto ScanForMe

* [```ScanForMe.xml```](/ScanForMe/ScanForMeManifest/ScanForMe.xml) El [archivo de manifiesto](https://dev.office.com/docs/add-ins/outlook/manifests/manifests) del complemento de Outlook.

#### Proyecto ScanForMeWeb

* [```ItemRead.html```](/ScanForMeWeb/ItemRead.html) La interfaz de usuario HTML del complemento de Outlook.
* [```ItemRead.js```](/ScanForMeWeb/ItemRead.js) El código de JavaScript que usó Home.html para interactuar con Word mediante la API de JavaScript para Office. 


<a name="codedescription"></a>
##Descripción del código

La lógica básica de este ejemplo está en el archivo [```ItemRead.js```](/ScanForMeWeb/ItemRead.js) en el proyecto ScanForMeWeb. 

Una vez que se inicializa el complemento, las propiedades `item.to` y `item.cc` se analizan en busca de la presencia de la dirección de correo electrónico del usuario. La dirección de correo electrónico del usuario se recupera de la propiedad [```Office.context.mailbox.userProfile```](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.userProfile). Si el usuario se encuentra en las líneas Para o CC de este correo electrónico, este hecho se registra en la interfaz de usuario del complemento. 

El método [```getAsync()```](http://dev.office.com/reference/add-ins/outlook/Body) del objeto Body se usa después para recuperar el cuerpo del correo electrónico en formato de texto. Cuando se completa esta operación asincrónica, se invoca la función de devolución de llamada insertada. Esta función usa una expresión regular para analizar el texto del cuerpo del correo en busca de repeticiones del nombre del usuario. Si se encuentran una o más repeticiones, la interfaz de usuario del complemento detecta que se ha mencionado al usuario en el cuerpo del correo. 

>Nota: Para obtener un ejemplo del uso de getAsync para recuperar el cuerpo de un correo electrónico en formato HTML, consulte el ejemplo [Outlook-Add-in-LinkRevealer](https://github.com/OfficeDev/Outlook-Add-in-LinkRevealer). 


<a name="build"></a>
##Compilar y depurar
1. Abra el archivo [```ScanForMe.sln```](ScanForMe.sln) en Visual Studio.
2. Pulse F5 para crear e implementar el complemento de ejemplo
3. Cuando se inicie Outlook, seleccione un correo electrónico de la Bandeja de entrada
4. Inicie el complemento seleccionándolo en la barra de la aplicación del complemento

<a name="questions"></a>
## Preguntas y comentarios

- Si tiene algún problema para ejecutar este ejemplo, [registre un problema](https://github.com/OfficeDev/Outlook-Add-in-ScanForMe/issues).
- Las preguntas sobre el desarrollo de complementos para Office en general deben enviarse a [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Asegúrese de que sus preguntas o comentarios se etiquetan con [office-addins].


<a name="contribute"></a>
## Contribuciones ##
Le animamos a contribuir a nuestros ejemplos. Para obtener instrucciones sobre cómo continuar, consulte nuestra [guía de contribución](./Contributing.md)

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea las [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.


<a name="additional-resources"></a>
## Recursos adicionales ##

- [Más ejemplos de complementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Complementos de Office](https://dev.office.com/reference/add-ins)
- [Componentes de un complemento](https://dev.office.com/docs/add-ins/overview/office-add-ins#StartBuildingApps_AnatomyofApp)
- [Crear un complemento de Office con Visual Studio](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio)


## Copyright
Copyright (c) 2015 Microsoft. Todos los derechos reservados.
