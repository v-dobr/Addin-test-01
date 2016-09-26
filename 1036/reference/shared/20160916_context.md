
# Context, objet
Représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans **|1.1|

```
Office.context
```


## Membres

|||
|:-----|:-----|
|Name|Description|
|[commerceAllowed](../../reference/shared/office.context.commerceallowed.md)|Obtient des informations indiquant si le complément est exécuté sur une plateforme qui autorise les liens vers des systèmes de paiement externes.|
|[contentLanguage](../../reference/shared/office.context.contentlanguage.md)|Obtient les paramètres régionaux (langue) des données, tels qu’ils sont stockés dans le document ou l’élément.|
|[displayLanguage](../../reference/shared/office.context.displaylanguage.md)|Obtient les paramètres régionaux (langue) de l’interface utilisateur de l’application hôte.|
|[document](../../reference/shared/office.context.document.md)|Obtient un objet qui représente le document avec lequel le complément de contenu ou de volet de tâches interagit.|
|[boîte aux lettres](../../reference/shared/office.context.mailbox.md)|Obtient l’objet **mailbox** qui donne accès aux membres de l’API spécifiquement destinés aux compléments Outlook.|
|[officeTheme](../../reference/shared/office.context.officetheme.md)|Permet d’accéder aux propriétés pour les couleurs du thème Office.|
|[ui](../../reference/shared/officeui)|Fournit des objets et des méthodes permettant de créer et de manipuler des composants d’interface utilisateur, comme les boîtes de dialogue.|
|[roamingSettings](../../reference/shared/office.context.roamingsettings.md)|Obtient un objet qui représente les paramètres personnalisés enregistrés du complément.|
|[touchEnabled](../../reference/shared/office.context.touchenabled.md)|Obtient des informations indiquant si le complément est exécuté dans une application hôte Office tactile.|

## Remarques

L’objet **Context** donne accès aux objets clés de l’API JavaScript pour Office.


## Informations de prise en charge



|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Ajout des propriétés **commerceAllowed** et **touchEnabledAdded** (Excel, PowerPoint et Word dans Office pour iPad uniquement).|
|1.1|Prise en charge supplémentaire des compléments avec Excel et Word sur Office pour iPad.|
|1.1|Pour [contentLanguage](../../reference/shared/office.context.contentlanguage.md), [displayLanguage](../../reference/shared/office.context.displaylanguage.md) et [document](../../reference/shared/office.context.document.md), prise en charge supplémentaire des compléments de contenu pour Access.|
|1.0|Introduit|
