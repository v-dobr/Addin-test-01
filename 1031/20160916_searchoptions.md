# SearchOptions-Objekt (JavaScript-API für Word)

Gibt die Optionen an, die in einem Suchvorgang eingeschlossen werden sollen.

__Gilt für: Word 2016, Word für iPad, Word für Mac_

## Eigenschaften
| Eigenschaft     | Typ   |Beschreibung
|:---------------|:--------|:----------|
|ignorePunct|bool|Ruft einen Wert ab oder legt ihn fest, um anzugeben, dass alle Interpunktionszeichen zwischen Wörtern ignoriert werden sollen. Entspricht dem Kontrollkästchen Interpunktionszeichen ignorieren im Dialogfeld Suchen und Ersetzen.|
|ignoreSpace|bool|Ruft einen Wert ab oder legt ihn fest, um anzugeben, dass alle Leerzeichen zwischen Wörtern ignoriert werden sollen. Entspricht dem Kontrollkästchen Leerzeichen ignorieren im Dialogfeld Suchen und Ersetzen.|
|matchCase|bool|Ruft einen Wert ab oder legt ihn fest, um anzugeben, dass die Suche mit Berücksichtigung von Groß-/Kleinschreibung durchgeführt werden soll. Entspricht dem Kontrollkästchen Groß-und Kleinschreibung beachten im Dialogfeld Suchen und Ersetzen (Menü Bearbeiten).|
|matchPrefix|bool|Ruft einen Wert ab oder legt ihn fest, um anzugeben, ob Wörter beachtet werden sollen, die mit der Suchzeichenfolge beginnen. Entspricht dem Kontrollkästchen Präfix beachten im Dialogfeld Suchen und Ersetzen (Menü Bearbeiten).|
|matchSoundsLike|bool|**Diese Option ist im Update vom Juni 2016 veraltet**. Ruft einen Wert ab oder legt ihn fest, um anzugeben, ob Wörter gesucht werden sollen, die ähnlich wie die Suchzeichenfolge klingen. Entspricht dem Kontrollkästchen Klingt wie im Dialogfeld Suchen und Ersetzen|
|matchSuffix|bool|Ruft einen Wert ab oder legt ihn fest, um anzugeben, ob Wörter beachtet werden sollen, die mit der Suchzeichenfolge enden. Entspricht dem Kontrollkästchen Suffix beachten im Dialogfeld Suchen und Ersetzen.|
|matchWholeWord|bool|Ruft einen Wert ab oder legt ihn fest, um anzugeben, dass nur ganze Wörter und nicht nach Textelementen in längeren Wörtern gesucht werden soll. Entspricht dem Kontrollkästchen Nur ganzes Wort suchen im Dialogfeld Suchen und Ersetzen.|
|matchWildCards|bool|Ruft einen Wert ab oder legt ihn fest, um anzugeben, ob die Suche mithilfe von speziellen Suchoperatoren ausgeführt wird. Entspricht dem Kontrollkästchen Platzhalter verwenden im Dialogfeld Suchen und Ersetzen.|

_Weitere Informationen finden Sie in den Eigenschaftszugriffs[beispielen.](#property-access-examples)_

Suchoptionen sind optional. Die Suchoptionen sollten in allen Suchmethoden mithilfe eines Objektsliterals angegeben werden:

```js
    search('searchstring', {searchOption1:bool, ...searchOptionN:bool}
```

Sie können eine oder mehrere Suchoptionseigenschaften im Objektliteral angeben, um Suchoptionen anzugeben. 

## Beziehungen
Keine


## Methoden

| Methode           | Rückgabetyp    |Beschreibung|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.|

## Methodendetails

### load(param: object)
Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Objektwerten.

#### Syntax
```js
object.load(param);
```

#### Parameter
| Parameter    | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|Parameter|object|Optional. Akzeptiert Parameter und Beziehungsnamen als getrennte Zeichenfolge oder Array. Oder geben Sie das [loadOption](loadoption.md)-Objekt an.|

#### Gibt 
void

## Eigenschaftszugriffsbeispiele

### Interpunktionssuche ignorieren
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Suche auf Grundlage eines Präfixes
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document based on a prefix.
    var searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Suche auf Grundlage eines Suffixes
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document for any string of characters after 'ly'.
    var searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'orange';
            searchResults.items[i].font.highlightColor = 'black';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Suche mithilfe eines Platzhalters
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildCards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


## Leitfaden zu Platzhaltern 

| für die Suche nach:         | Platzhalter |  Beispiel |
|:-----------------|:--------|:----------|
| Ein beliebiges Zeichen| ? |s? t sucht sat und legt es fest. |
|Eine beliebige Zeichenfolge| * |s*d findet sad und started.|
|Der Anfang eines Worts|< |<(inter) findet interesting und intercept, jedoch nicht splintered.|
|Das Ende eines Wortes |> |(in)> finds in und within, jedoch nicht interesting.|
|Eines der angegebenen Zeichen|[ ] |w[io]n findet win und won.|
|Beliebiges einzelnes Zeichen in diesem Bereich| [-] |[r-t]ight findet right und sight. Bereiche müssen in aufsteigender Reihenfolge sortiert sein.|
|Ein beliebiges einzelnes Zeichen mit Ausnahme der Zeichen im Bereich in eckigen Klammern|[!x-z] |t[!a-m]ck findet tock und tuck, jedoch nicht tack oder tick.|
|Genau n Vorkommen des vorhergehenden Zeichens oder Ausdrucks|{n} |fe\{2\}d findet feed jedoch nicht fed.|
|Mindestens n Vorkommen des vorhergehenden Zeichens oder Ausdrucks|{n,} |fe{1,}d findet fed und feed.|
|Von n bis m Vorkommen des vorhergehenden Zeichens oder Ausdrucks|{n,m} |10{1,3} findet 10, 100 und 1000.|
|Ein oder mehr Vorkommen des vorhergehenden Zeichens oder Ausdrucks|@ |lo@t finds lot und loot.|


## Supportdetails
Verwenden Sie den [Anforderungssatz](../office-add-in-requirement-sets.md) in Laufzeitüberprüfungen, um sicherzustellen, dass die Anwendung von der Hostversion von Word unterstützt wird. Weitere Informationen zur Office-Hostanwendung und den Serveranforderungen finden Sie unter [Anforderungen für die Ausführung von Office-Add-Ins](../../docs/overview/requirements-for-running-office-add-ins.md).
