# Übersicht über die JavaScript-API-Programmierung für Excel

In diesem Artikel wird beschrieben, wie Sie die Excel-JavaScript-API verwenden können, um Add-Ins für Excel 2016 zu erstellen. Der Artikel enthält eine Einführung in die wichtigsten Konzepte zur Verwendung der APIs, z. B. RequestContext, JavaScript-Proxyobjekte, sync(), Excel.run() und load(). In den Codebeispielen am Ende des Artikels gezeigt, wie die Konzepte gelten.

## RequestContext

Das RequestContext-Objekt erleichtert das Senden von Anforderungen an die Excel-Anwendung. Da das Office-Add-In und die Excel-Anwendung in zwei verschiedenen Prozessen ausgeführt werden, ist Anforderungskontext erforderlich, um auf Excel und die zugehörigen Objekte, wie z. B. Arbeitsblätter und Tabellen aus dem Add-In zuzugreifen. Wie dargestellt wird ein Anforderungskontext erstellt.

```js
var ctx = new Excel.RequestContext();
```

## Proxyobjekte

Die in einem Add-In deklarierten und verwendeten Excel-JavaScript-Objekte sind Proxyobjekte für die realen Objekte in einem Excel-Dokument. Alle Aktionen zu Proxyobjekten werden in Excel nicht realisiert, und der Status des Excel-Dokuments wird in den Proxyobjekten erste realisiert, wenn der Status des Dokuments synchronisiert wurde. Der Dokumentstatus wird synchronisiert, wenn die context.sync() ausgeführt wird (siehe unten).

Das lokale JavaScript-Objekt `selectedRange` wird beispielsweise deklariert, um auf den ausgewählten Bereich zu verweisen. Damit kann die Einstellung der Eigenschaften und das Abrufen von Methoden in der Warteschlange verwendet werden. Die Aktionen für solche Objekte werden nicht realisiert, bis die Methode sync() ausgeführt wird.

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

Die zum Anforderungskontext verfügbare sync()-Methode synchronisiert den Status zwischen JavaScript-Proxyobjekten und realen Objekten in Excel, indem es die zum Kontext und zum Abrufen von Eigenschaften der geladenen Office-Objekte zur Verwendung in Ihrem Code in die Warteschlange gestellten Anweisungen ausführt.  Diese Methode gibt eine Zusage zurück, die nach Abschluss der Synchronisierung aufgelöst wird.

## Excel.run(function(context) { batch })

Excel.run() führt ein Batch-Skript aus, das Aktionen zum Excel-Objektmodell durchführt. Die Batchbefehle enthalten Definitionen lokaler JavaScript-Proxy-Objekte und sync()-Methoden, die den Status zwischen den lokalen und Excel-Objekten und der Zusage-Auflösung synchronisieren. Der Vorteil von Batch-Anforderungen in Excel.run() ist, dass bei der Zusage-Auflösung aller nachverfolgten Bereichsobjekte, die während der Ausführung zugeordnet wurden, automatisch freigegeben werden.

Die Run-Methode nimmt in RequestContext und gibt eine Zusage zurück (in der Regel nur das Ergebnis der ctx.sync()). Es ist möglich, den Batchvorgang außerhalb der Excel.run() auszuführen. In einem solchen Szenario müssen die Bereichsobjektverweise jedoch manuell nachverfolgt und verwaltet werden.

## load()

Die load()-Methode dient zum Auffüllen der Proxyobjekte, die auf der JavaScript-Ebene im Add-In erstellt werden. Beim Abrufen eines Objekts, z. B. eines Arbeitsblatts, wird ein lokales Proxyobjekt zunächst auf der JavaScript-Ebene erstellt. Damit kann die Einstellung der Eigenschaften und das Abrufen von Methoden in der Warteschlange verwendet werden. Zum Lesen von Objekteigenschaften oder Beziehungen müssen die load()- und sync()-Methoden zuerst aufgerufen werden. Die load()-Methode nimmt die Eigenschaften und Beziehungen auf, die beim Aufrufen der sync()-Methode geladen werden müssen.

_Syntax:_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
Dabei gilt:

* `properties` ist eine Liste der zu ladenden Eigenschaften und/oder Beziehungsnamen, die als durch Trennzeichen getrennte Zeichenfolgen oder Namen-Arrays angegeben wurden. Weitere Informationen dazu finden Sie in den .load()-Methoden unter den einzelnen Objekten.
* `loadOption` gibt ein Objekt an, das die Optionen für Auswahl, Erweiterungs, oben und Überspringen beschreibt. Weitere Informationen finden Sie im Objekt [Ladeoptionen](../../reference/excel/loadoption.md).

## Beispiel: Schreiben von Werten aus einem Array in ein Bereichsobjekt

Im folgenden Beispiel wird gezeigt, wie Sie Werte von einem Array in ein Bereichsobjekt schreiben

Die Excel.run() enthält eine Reihe von Anweisungen. Als Teil dieses Batches wird ein Proxyobjekt erstellt, das auf einen Bereich (Adresse A1:B2) im aktiven Arbeitsblatt verweist. Der Wert dieses Proxybereichsobjekts wird lokal festgelegt. Um die Werte auszulesen, wird die `text`-Eigenschaft des Bereichs angewiesen, in das Proxyobjekt geladen zu werden. Alle diese Befehle werden in die Warteschlange gestellt und beim Aufrufen von ctx.sync() ausgeführt. Die sync()-Methode gibt eine Zusage zurück, mit der sie mit anderen Vorgängen verkettet werden kann.

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

## Beispiel: Kopieren von Werten

Im folgende Beispiel wird gezeigt, wie Sie die Werte aus dem Bereich A1:A2 bis B1:B2 des aktiven Arbeitsblatts mithilfe der load()-Methode in das Bereichsobjekt kopieren können.

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

## Auswahl von Eigenschaften und Beziehungen

Standardmäßig wählt object.load() alle skalaren und komplexen Eigenschaften des Objekts aus, das geladen wird. Die Beziehungen werden standardmäßig nicht geladen (z. B. ist das Format ein Beziehungsobjekt des Bereichsobjekts). Allerdings empfiehlt es sich, die zu ladenden Eigenschaften und Beziehungen explizit zu markieren, um die Leistung zu verbessern. Dies kann durch Angeben (im `load()`-Parameter) einer Teilmenge der Eigenschaften und Beziehungen in der Antwort erfolgen. Die Load-Methode lässt zwei Arten von Eingaben zu:

* Eigenschafts- und Beziehungsnamen als durch Trennzeichen getrennte Zeichenfolgennamen _oder_ als Array von Zeichenfolgen, die Eigenschafts- oder Beziehungsnamen enthalten.
* Ein Objekt, das die Auswahl, Erweiterung, obere und Überspringen-Optionen beschreibt. Weitere Informationen finden Sie im Objekt [Ladeoptionen](../../reference/excel/loadoption.md).

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### Beispiel

Die folgende Load-Anweisung lädt alle Eigenschaften des Bereichs und wird dann zum Format und Format/Füllung erweitert.

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

## NULL-Eingabe

### NULL-Eingabe im 2D-Array

`null` Eingabe in zweidimensionalen Arrays (für Werte, Zahlenformate, Formeln) werden in der Update-API ignoriert. Am beabsichtigten Ziel findet keine Aktualisierung statt, wenn die `null`-Eingabe in Werten, im Zahlenformat oder Formelraster von Werten.

Beispiel: Um nur bestimmte Teile des Bereichs zu aktualisieren, wie z. B. das Zahlenformat einer Zelle, und das vorhandene Zahlenformat in anderen Teilen des Bereichs beizubehalten, legen Sie das gewünschte Zahlenformat an entsprechender Stelle fest und senden Sie `null` für die anderen Zellen.

In der folgenden Set-Anforderung werden nur einige Teile des Bereichszahlenformats festgelegt, dabei wird das vorhandene Zahlenformat im verbleibenden Teil beibehalten.

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### NULL-Eingabe für eine Eigenschaft

`null` ist keine gültige einzelne Eingabe für die gesamte Eigenschaft. Folgende Eingabe ist beispielsweise ungültig, da die gesamten Werte nicht auf NULL oder ignoriert festgelegt werden können.

```js
 range.values= null;

```

Folgendes ist nicht gültig, da NULL kein gültiger Farbwert ist.

```js
 range.format.fill.color =  null;
```

### NULL-Antwort

Darstellungen der Formatierungseigenschaften, die aus ungleichmäßigen Werten bestehen, ergeben einen NULL-Wert in der Antwort.

Beispiel: Ein Bereich kann aus einer oder mehreren Zellen bestehen. In Fällen, in denen im angegebenen Bereich einzelne Zellen enthalten sind, besitzen keine gleichmäßigen Formatierungswerte, deshalb wird die Bereichsebenendarstellung ungleichmäßig.

```js
  "size" : null,
  "color" : null,
```

### Leere Eingabe und Ausgabe

Leere Werte in Aktualisierungsanforderungen werden als Anweisung behandelt, um die entsprechende Eigenschaft zu löschen oder zurückzusetzen. Ein Leerer Wert wird durch zwei doppelte Anführungszeichen ohne Leerzeichen dazwischen dargestellt. `""`

Beispiel:

* Für `values` wird der Bereichswert gelöscht. Das ist genauso wie das Löschen des Inhalts in der Anwendung.

* Für `numberFormat` wird das Zahlenformat auf `General` festgelegt.

* Für `formula` und `formulaLocale` werden die Formelwerte gelöscht.


Für Lesevorgänge können Sie leere Werte erwarten, wenn der Inhalt der Zellen leer ist. Wenn die Zelle keine Daten oder keinen Wert enthält, gibt die API einen leeren Wert zurück. Ein Leerer Wert wird durch zwei doppelte Anführungszeichen ohne Leerzeichen dazwischen dargestellt. `""`

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## Ungebundener Bereich

### Lesen

Eine Ungebundener Bereichsadresse enthält nur Bezeichner für die Spalte oder Zeile und nicht angegebene Zeilenbezeichner oder Spaltenbezeichner, wie etwa:

* `C:C`, `A:F`, `A:XFD` (enthält nicht angegebene Zeilen)
* `2:2`, `1:4`, `1:1048546` (enthält nicht angegebene Spalten)

Wenn die API zum Abrufen eines ungebundenen Bereichs (z. B. `getRange('C:C')` eine Anforderung ausführt, enthält die zurückgegebene Antwort `null` für Zellenebeneneigenschaften wie `values`, `text`, `numberFormat`, `formula`, usw. Andere Bereichseigenschaften wie `address`, `cellCount` usw. spiegeln den ungebundenen Bereich wider.

### Write

Das Festlegen von Zelleigenschaftsebenen (z. B. Werte, NumberFormat usw.) für einen ungebundener Bereich ist **nicht zulässig,** da die Eingabeanforderung möglicherweise zu groß zum Verarbeiten ist.

Beispiel: Folgendes ist keine gültige Aktualisierungsanforderung, da der angeforderte Bereich ungebunden ist.

```js
...
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
    range.values = 'Due Date';
...
```

Wenn ein Aktualisierungsvorgang für so einen Bereich ausgeführt wird, gibt die API einen Fehler zurück.


## Großer Bereich

Ein großer Bereich impliziert einen Bereich, dessen Größe für einen einzelnen API-Aufruf zu groß ist. Viele Faktoren wie z. B. die Anzahl der Zellen, Werte, numberFormat und Formeln, die im Bereich enthalten sind, können die Antwort so groß werden lassen, dass es zur API-Interaktion ungeeignet werden kann. Die API ermöglicht einen optimalen Versuch zum Zurückgeben oder Schreiben der angeforderten Daten. Allerdings kann die Größe aufgrund der hohen Ressourcenverwendung zu einem API-Fehlerzustand führen.

Um einen solchen Zustand zu vermeiden, wird empfohlen, das Lesen oder Schreiben für einen großen Bereich in mehreren kleineren Bereichsgrößen zu verwenden.


## Einzelne Eingabekopie

Zur Unterstützung der Aktualisierung eines Bereichs mit denselben Werten oder demselben Zahlenformat oder des Anwendens derselben Formel für einen kompletten Bereich, wird in der Set-API folgende Konvention verwendet. In Excel ähnelt dieses Verhalten dem Eingeben von Werten oder Formeln in einen Bereich im Modus STRG + EINGABETASTE.

Die API sucht nach einem *einzelnen Zellenwert* und wenn die Ziel-Bereichsdimension nicht mit der Eingabebereichsdimension übereinstimmt, wendet sie die Aktualisierung im STRG + EINGABE-Modell mit dem in der Anforderung bereitgestellten Wert oder der Formel auf den gesamten Bereich an.

### Beispiele

Die folgende Anforderung aktualisiert den ausgewählten Bereich mit dem Text „Fälligkeitsdatum“. Beachten Sie, dass der Bereich 20 Zellen aufweist, während die angegebene Eingabe nur einen Zellwert besitzt.

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

Die folgende Anforderung aktualisiert den ausgewählten Bereich mit dem Datum „3.11..2015“.

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
Durch die folgende Anforderung wird der ausgewählte Bereich mit einer Formel aktualisiert, die im Modus STRG + EINGABETASTE auf den Bereich angewendet wird.

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


## Fehlermeldungen

Fehler werden mithilfe eines Fehlerobjekts zurückgegeben, das aus einem Code und einer Nachricht besteht. Die folgende Tabelle enthält eine Liste möglicher Fehlerzustände, die auftreten können.

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |Das Argument ist ungültig oder fehlt oder weist ein falsches Format auf.|
|InvalidRequest  |Die Anforderung kann nicht verarbeitet werden.|
|InvalidReference|Dieser Verweis ist für den aktuellen Vorgang nicht gültig.|
|InvalidBinding  |Die Objektbindung ist aufgrund von früheren Updates nicht mehr gültig.|
|InvalidSelection|Die aktuelle Auswahl ist für diesen Vorgang nicht gültig.|
|Nicht authentifiziert |Erforderliche Authentifizierungsinformationen fehlen oder sind ungültig.|
|AccessDenied   |Sie können den angeforderten Vorgang nicht durchzuführen.|
|ItemNotFound   |Die angeforderte Ressource ist nicht vorhanden.|
|ActivityLimitReached|Der Aktivitätsgrenzwert wurde erreicht.|
|GeneralException|Beim Verarbeiten der Anforderung ist ein interner Fehler aufgetreten.|
|NotImplemented  |Das angeforderte Feature ist nicht implementiert.|
|ServiceNotAvailable|Der Dienst ist nicht verfügbar.|
|Conflict   |Anforderung konnte aufgrund eines Konflikts nicht verarbeitet werden.|
|ItemAlreadyExists|Die erstellte Ressource ist bereits vorhanden.|
|UnsupportedOperation|Dieser Vorgang wird nicht unterstützt.|
|RequestAborted|Die Anforderung wurde während der Laufzeit abgebrochen.|
|ApiNotAvailable|Die angeforderte API ist nicht verfügbar.|
|InsertDeleteConflict|Der Einfüge- oder Löschvorgang führte zu einem Konflikt.|
|InvalidOperation|Dieser Vorgang ist für das Objekt ungültig.|

## Weitere Ressourcen

* [Erstellen Ihres ersten Excel-Add-Ins](build-your-first-excel-add-in.md)
* [Codeausschnitt-Explorer](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Codebeispiele zu Excel-Add-Ins](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [JavaScript-API-Referenz zu Excel-Add-Ins](excel-add-ins-javascript-api-reference.md)
