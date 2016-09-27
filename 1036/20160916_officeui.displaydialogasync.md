# Méthode UI.displayDialogAsync

Affiche une boîte de dialogue dans un hôte Office. 

## Configuration requise

|Hôte|Nouveauté de|Dernière modification dans |
|:---------------|:--------|:----------|
|Word, Excel, PowerPoint|1.1|1.1|
|Outlook|Mailbox 1.4|Mailbox 1.4|

Cette méthode est disponible dans l’[ensemble de conditions requises](../../docs/overview/specify-office-hosts-and-api-requirements.md) DialogAPI. Pour spécifier l’ensemble de conditions requises DialogAPI, utilisez le code suivant dans votre manifeste.

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="DialogAPI"/> 
   </Sets> 
 </Requirements> 

```

Pour détecter cette API lors de son exécution, utilisez le code suivant.

```js
 if (Office.context.requirements.isSetSupported('DialogAPI', 1.1)) 
    {  
         // Use Office UI methods; 
    } 
 else 
     { 
         // Alternate path 
     } 
```



### Plateformes prises en charge
L’ensemble de conditions requises DialogAPI est actuellement pris en charge sur les plateformes suivantes :

  - Office pour Windows 2016 pour ordinateur de bureau (version 16.0.6741.0000 ou ultérieure)
  - Office pour iPad (version 1.22 ou ultérieure)
  - Office pour Mac (version 15.20 ou ultérieure) 

D’autres plateformes seront bientôt disponibles. 

## Syntaxe

```js
office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##Exemples

Pour obtenir un exemple simple qui utilise la méthode **displayDialogAsync**, consultez l’[exemple de boîte de dialogue API de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/) sur GitHub.

Pour obtenir un exemple qui illustre un scénario d’authentification, consultez l’exemple d’[authentification client Office 365 de complément Office pour AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth) sur GitHub.

 
## Paramètres

| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|startAddress|string|Accepte l’URL HTTPS(TLS) initiale qui s’ouvre dans la boîte de dialogue. <ul><li>La page initiale doit figurer sur le même domaine que la page parent. Après le chargement de la page initiale, vous pouvez également accéder à d’autres domaines.</li><li>Toute page appelant [office.context.ui.messageParent](officeui.messageparent.md) doit également figurer sur le même domaine que la page parent.</li></ul>|
|options|object|Facultatif. Accepte un objet options pour définir les comportements de la boîte de dialogue.|
|callback|object|Accepte une méthode de rappel pour gérer la tentative de création de boîte de dialogue.|
    
### Options de configuration
Les options de configuration suivantes sont disponibles pour une boîte de dialogue.


| Propriété     | Type   |Description|
|:---------------|:--------|:----------|
|**width**|object|Facultatif. Définit la largeur de la boîte de dialogue sous forme de pourcentage de l’affichage actuel. La valeur par défaut est 80 %. La résolution minimale est de 250 pixels.|
|**height**|object|Facultatif. Définit la hauteur de la boîte de dialogue sous forme de pourcentage de l’affichage actuel. La valeur par défaut est 80 %. La résolution minimale est de 150 pixels.|
|**displayInIFrame**|object|Facultatif. Détermine si la boîte de dialogue doit être affichée dans un IFrame dans les clients Office Online. Ce paramètre est ignoré par les clients de bureau. Les valeurs possibles sont les suivantes :<ul><li>False (valeur par défaut) : la boîte de dialogue s’affichera dans une nouvelle fenêtre de navigateur (fenêtre contextuelle). Recommandé pour les pages d’authentification qui ne peuvent pas être affichées dans un IFrame. </li><li>True : la boîte de dialogue s’affichera sous la forme d’une fenêtre flottante avec un IFrame. Recommandé pour une expérience utilisateur et des performances optimales.</li>|


## Valeur de rappel
Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **displayDialogAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à l’objet [Dialog](../../reference/shared/officeui.dialog.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou objet défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|


## Considérations relatives à la conception
Les considérations relatives à la conception ci-dessous s’appliquent aux boîtes de dialogue :

- Un complément Office ne peut comporter qu’une seule boîte de dialogue ouverte à la fois.
- Toutes les boîtes de dialogue peuvent être déplacées et redimensionnées par l’utilisateur.
- Toutes les boîtes de dialogue s’affichent au centre de l’écran à l’ouverture.
- Les boîtes de dialogue s’affichent au-dessus de l’application hôte et dans l’ordre dans lequel elles ont été créées.

Utilisez une boîte de dialogue pour :

- Afficher les pages d’authentification permettant de collecter les informations d’identification de l’utilisateur.
- Afficher un écran d’erreur/de progression/de saisie à partir d’une commande ShowTaskpane ou ExecuteAction.
- Augmenter provisoirement la surface dont un utilisateur dispose pour effectuer une tâche.

N’utilisez pas de boîte de dialogue pour interagir avec un document. Il est préférable d’utiliser un volet des tâches. 

Pour obtenir un exemple de modèle de conception à utiliser pour créer une boîte de dialogue, consultez [Boîte de dialogue client](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md) dans le référentiel relatif aux modèles de conception de l’expérience utilisateur du complément Office sur GitHub.
