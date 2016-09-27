# Vue d’ensemble de la programmation de l’API JavaScript d’Excel

Cet article décrit comment utiliser l’API JavaScript Excel pour créer des compléments pour Excel 2016. Il présente des concepts fondamentaux pour l’utilisation d’API, notamment concernant les objets RequestContext, les objets de proxy JavaScript, ainsi que les méthodes sync(), Excel.run() et load(). Les exemples de code à la fin de l’article vous montrent comment appliquer les concepts.

## RequestContext

L’objet RequestContext facilite les demandes auprès de l’application Excel. L’exécution du complément Office et de l’application Excel faisant appel à deux processus différents, il est nécessaire de fournir le contexte des demandes pour accéder à Excel et aux objets associés, tels que les feuilles de calcul et les tableaux, à partir du complément. L’exemple suivant illustre la création d’un contexte de demande.

```js
var ctx = new Excel.RequestContext();
```

## Objets de proxy

Les objets JavaScript Excel déclarés et utilisés dans un complément sont des objets de proxy correspondant aux objets réels d’un document Excel. Toutes les actions effectuées sur les objets de proxy ne sont pas réalisées dans Excel et l’état du document Excel n’est pas répercuté sur les objets de proxy tant que cet état n’a pas été synchronisé. L’état de document est synchronisé lors de l’exécution de la méthode context.sync() (voir ci-dessous).

Par exemple, l’objet `selectedRange` JavaScript local est déclaré pour référencer la plage sélectionnée. Cela permet par exemple de mettre en file d’attente la valeur de ses propriétés et méthodes d’appel. Les actions appliquées à ces objets ne sont pas réalisées jusqu’à l’exécution de la méthode sync().

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

La méthode sync() disponible dans le contexte de demande synchronise l’état des objets de proxy JavaScript et des objets réels d’Excel en exécutant les instructions mises en file d’attente sur le contexte et en récupérant les propriétés des objets Office chargés à utiliser dans votre code. Cette méthode renvoie une promesse, qui est résolue lorsque la synchronisation est terminée.

## Excel.run(function(context) { batch })

Excel.run() exécute un script de commandes qui effectue des actions sur le modèle objet Excel. Les commandes de traitement par lots incluent les définitions des objets de proxy JavaScript locaux et des méthodes sync() qui synchronisent l’état des objets locaux et Excel, ainsi que la résolution de la promesse. L’avantage de traiter les demandes par lots avec Excel.run() est que, une fois la promesse résolue, tous les objets de plage faisant l’objet d’un suivi qui ont été alloués lors de l’exécution sont automatiquement libérés.

La méthode d’exécution utilise le contexte de demande et renvoie une promesse (en général, le résultat de la méthode ctx.sync()). Il est possible d’exécuter l’opération par lots en dehors de la méthode Excel.run(). Toutefois, dans ce cas, toutes les références d’objet de plage doivent être suivies et gérées manuellement.

## load()

La méthode load() permet de remplir les objets de proxy créés dans le calque JavaScript du complément. Lorsque vous essayez de récupérer un objet, une feuille de calcul par exemple, un objet de proxy local est tout d’abord créé dans le calque JavaScript. Cet objet peut être utilisé pour mettre en file d’attente la valeur de ses propriétés et méthodes d’appel. Toutefois, pour la lecture des propriétés ou des relations de l’objet, les méthodes load() et sync() doivent d’abord être appelées. La méthode load() utilise les propriétés et les relations qui doivent être chargées lors de l’appel de la méthode sync().

_Syntaxe :_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
où :

* `properties` est la liste des propriétés et/ou des noms de relation à charger, fournie sous forme de chaînes séparées par des virgules ou de tableau de noms. Pour plus d’informations, consultez les méthodes .load() décrites sous chaque objet.
* `loadOption` spécifie un objet qui décrit les propriétés select, expand, top et skip. Pour plus d’informations, reportez-vous aux [options](../../reference/excel/loadoption.md) de chargement d’objet.

## Exemple : écrire des valeurs d’un tableau vers un objet de plage

L’exemple suivant vous montre comment écrire des valeurs d’un tableau vers un objet de plage.

La méthode Excel.run() contient un lot d’instructions. Dans le cadre de ce traitement par lots, un objet de proxy faisant référence à une plage (adresse A1:B2) est créé dans la feuille de calcul active. La valeur de cet objet de plage de proxy est définie localement. Pour pouvoir lire les valeurs, la propriété `text` de la plage est chargée sur l’objet proxy. Toutes ces commandes sont mises en file d’attente et sont exécutées lorsque la méthode ctx.sync() est appelée. La méthode sync() renvoie une promesse qui peut être utilisée pour y adjoindre d’autres opérations.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) {

    // Create a proxy object for the sheet
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    // Values to be updated
    var values = [
                 ["Type", "Estimate"],
                 ["Transportation", 1670]
                 ];
    // Create a proxy object for the range
    var range = sheet.getRange("A1:B2");

    // Assign array value to the proxy object's values property.
    range.values = values;

    // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
    return ctx.sync().then(function() {
            console.log("Done");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Exemple : copier des valeurs

L’exemple suivant montre comment copier les valeurs de la plage A1:A2 vers la plage B1:B2 de la feuille de calcul en utilisant la méthode load() sur l’objet de plage.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) {

    // Create a proxy object for the range
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");

    // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
    return ctx.sync().then(function() {
        // Assign the previously loaded values to the new range proxy object. The values will be updated once the following .then() function is invoked.
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = range.values;
    });
}).then(function() {
      console.log("done");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Sélection de relations et de propriétés

Par défaut, la méthode object.load() sélectionne toutes les propriétés scalaires et complexes de l’objet qui est chargé. Les relations ne sont pas chargées par défaut (par exemple, le format est un objet de relation de l’objet de plage). Toutefois, nous vous recommandons de marquer de façon explicite les propriétés et les relations à charger afin d’améliorer les performances. Pour cela, indiquez (dans le paramètre `load()`) un sous-ensemble de propriétés et de relations à inclure dans la réponse. La méthode Load autorise deux types d’entrées :

* Des noms de propriété et de relation, sous forme de chaînes séparées par des virgules _ou_ de tableau de chaînes contenant des noms de propriété ou de relation.
* Un objet qui décrit les options select, expand, top et skip. Pour plus d’informations, reportez-vous aux [options](../../reference/excel/loadoption.md) de chargement d’objet.

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### Exemple

L’instruction de chargement suivante charge toutes les propriétés de la plage, puis étend les valeurs de format et de format/remplissage.

```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:B2";
    var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load(["address", "format/*", "format/fill", "entireRow" ]);
    return ctx.sync().then(function() {
        console.log (myRange.address); //ok
        console.log (myRange.format.wrapText); //ok
        console.log (myRange.format.fill.color); //ok
        //console.log (myRange.format.font.color); //not ok as it was not loaded

    });
}).then(function() {
      console.log("done");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Entrée null

### entrée de valeurs null dans un tableau 2D

L’entrée `null` dans un tableau à deux dimensions (pour des valeurs, des formats de nombre ou des formules) est ignorée dans l’API de mise à jour. Aucune mise à jour n’est appliquée à la cible voulue quand l’entrée `null` est envoyée dans des grilles de valeurs, de formats de nombre ou de formules.

Exemple : afin de mettre à jour uniquement des parties spécifiques de la plage, par exemple pour modifier le format de nombre d’une cellule tout en conservant le format de nombre existant pour d’autres parties de la plage, indiquez le format de nombre souhaité là où un changement est nécessaire et envoyez `null` pour les autres cellules.

Dans la demande définie suivante, seules certaines parties du format de nombre de la plage sont définies et le format de nombre existant est conservé sur la partie restante (en indiquant des valeurs null).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### Entrée null pour une propriété

Vous ne pouvez pas indiquer `null` comme valeur unique pour l’ensemble de la propriété. Par exemple, l’exemple suivant n’est pas valide car vous ne pouvez pas ignorer ou définir sur null l’ensemble des valeurs.

```js
 range.values= null;

```

L’exemple ci-dessous n’est pas valide non plus, car null n’est pas une valeur de couleur valide.

```js
 range.format.fill.color =  null;
```

### Réponse null

La représentation de propriétés de mise en forme composées de valeurs non uniformes renvoie une valeur null comme réponse.

Exemple : une plage peut se composer de plusieurs cellules. Si des cellules de la plage spécifiée ont des valeurs de mise en forme différentes, aucune représentation ne pourra être définie au niveau de la plage entière.

```js
  "size" : null,
  "color" : null,
```

### Entrées et sorties vides

Les valeurs vides dans les demandes de mise à jour sont traitées comme des instructions visant à effacer ou à restaurer la valeur de la propriété concernée. Une valeur vide est représentée par des guillemets droits non séparés par un espace. `""`

Exemple :

* pour `values`, la valeur de plage est effacée. Cela revient à supprimer du contenu dans l’application.

* Pour `numberFormat`, le format de nombre est défini sur `General`.

* Pour `formula` et `formulaLocale`, les valeurs de formule sont effacées.


Pour les opérations de lecture, il est normal d’obtenir des valeurs vides si les cellules le sont également. Si la cellule ne contient aucune donnée ou valeur, l’API renvoie une valeur vide. Une valeur vide est représentée par des guillemets droits non séparés par un espace. `""`.

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## Plage illimitée

### Lecture

Une adresse de plage illimitée ne contient que des identificateurs de ligne ou de colonne, ainsi que des identificateurs de lignes ou de colonnes non spécifiées (respectivement), comme dans l’exemple ci-dessous :

* `C:C`, `A:F`, `A:XFD` (contient des lignes non spécifiées)
* `2:2`, `1:4`, `1:1048546` (contient des colonnes non spécifiées)

Lorsque l’API effectue une demande pour récupérer une plage illimitée (par exemple, `getRange('C:C')`), la réponse renvoyée contient `null` pour les propriétés définies au niveau des cellules, telles que `values`, `text`, `numberFormat`, `formula`, etc. D’autres propriétés de plage, telles que `address`, `cellCount`, etc. refléteront la plage illimitée.

### Écriture

Vous n’êtes **pas autorisé** à définir des propriétés de niveau cellule (par exemple, values, numberFormat, etc.) sur une plage illimitée, car la demande d’entrée risque d’être trop lourde à gérer.

Exemple : la demande de mise à jour suivante n’est pas valide, car la plage demandée est illimitée.

```js
...
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
    range.values = 'Due Date';
...
```

Lorsqu’une opération de mise à jour est tentée sur une plage de ce type, l’API renvoie une erreur.


## Plage de grande taille

Une plage de grande taille est une plage trop grande pour pouvoir être gérée par un seul appel à l’API. De nombreux facteurs tels que le nombre de cellules, les valeurs, le format de nombre, et les formules contenues dans la plage peuvent entraîner une réponse trop lourde pour l’interaction avec l’API. L’API tente de renvoyer ou d’écrire les données requises en faisant au mieux de ses possibilités. Toutefois, la taille importante de la réponse peut provoquer une erreur de l’API à cause de la grande quantité de ressources mobilisées.

Pour éviter cela, nous vous recommandons de fractionner les plages de grande taille en plusieurs plages plus petites pour vos opérations de lecture et d’écriture.


## Copie d’une seule entrée

Pour mettre à jour une plage contenant des formats de nombre ou des valeurs uniformes, ou pour appliquer une même formule à l’ensemble d’une plage, la convention suivante est utilisée dans l’API définie. Dans Excel, cela revient à attribuer des valeurs ou des formules à une plage en mode CTRL + ENTRÉE.

L’API recherche une *valeur de cellule unique* et, si la dimension de la plage cible ne correspond pas à la dimension de la plage d’entrée, elle applique la mise à jour à la plage entière en mode CTRL + ENTRÉE avec la valeur ou la formule indiquée dans la demande.

### Exemples

La demande suivante met à jour la plage sélectionnée en y ajoutant le texte « Due Date ». Notez que la plage comporte 20 cellules, tandis que l’entrée fournie comporte seulement 1 valeur de cellule.

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.values = 'Due Date';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

La demande suivante met à jour la plage sélectionnée en y ajoutant la date « 3/11/2015 ».

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.numberFormat = 'm/d/yyyy';
    range.values = '3/11/2015';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
La requête suivante met à jour la plage sélectionnée en y ajoutant une formule qui sera appliquée en mode CTRL + ENTRÉE.

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.numberFormat = 'm/d/yyyy';
    range.values = '3/11/2015';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


## Messages d’erreur

Les erreurs sont renvoyées à l’aide d’un objet d’erreur qui se compose d’un code et d’un message. Le tableau suivant fournit la liste des erreurs qui peuvent se produire.

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |L’argument est manquant ou non valide, ou a un format incorrect.|
|InvalidRequest  |Impossible de traiter la demande.|
|InvalidReference|Cette référence n’est pas valide pour l’opération en cours.|
|InvalidBinding  |Cette liaison d’objets n’est plus valide en raison de mises à jour précédentes.|
|InvalidSelection|La sélection en cours est incorrecte pour cette action.|
|Unauthenticated |Les informations d’authentification requises sont manquantes ou incorrectes.|
|AccessDenied   |Vous ne pouvez pas effectuer l’opération demandée.|
|ItemNotFound   |La ressource demandée n’existe pas.|
|ActivityLimitReached|La limite d’activité a été atteinte.|
|GeneralException|Une erreur interne s’est produite lors du traitement de la demande.|
|NotImplemented  |La fonctionnalité demandée n’est pas implémentée|
|ServiceNotAvailable|Le service n’est pas disponible.|
|Conflict   |La demande n’a pas pu être traitée en raison d’un conflit.|
|ItemAlreadyExists|La ressource en cours de création existe déjà.|
|UnsupportedOperation|L’opération tentée n’est pas prise en charge.|
|RequestAborted|La demande a été interrompue pendant l’exécution.|
|ApiNotAvailable|L’API demandée n’est pas disponible.|
|InsertDeleteConflict|L’opération d’insertion ou de suppression tentée a créé un conflit.|
|InvalidOperation|L’opération tentée n’est pas valide sur l’objet.|

## Ressources supplémentaires

* [Création de votre premier complément Excel](build-your-first-excel-add-in.md)
* [Explorateur d’extraits de code](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Exemples de code pour les compléments Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Référence de l’API JavaScript pour les compléments Excel](excel-add-ins-javascript-api-reference.md)
