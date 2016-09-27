# LoadOption-Objekt (JavaScript-API für Word)

Ein Objekt, das Paginginformationen und Eigenschaften enthält, die bei Aufrufen von „context.sync()“ geladen werden.

__Gilt für: Word 2016, Word für iPad, Word für Mac_

## Eigenschaften
| Eigenschaft     | Typ   |Beschreibung|
|:---------------|:--------|:----------|
|select|object|Enthält eine durch Trennzeichen getrennte Liste oder ein Array von Parametern/Verhältnisnamen. Optional.|
|expand|object|Enthält eine durch Trennzeichen getrennte Liste oder ein Array von Verhältnisnamen. Optional.|
|top|int| Gibt die maximale Anzahl der Sammlungselemente an, die im Ergebnis enthalten sein können. Optional. Sie können diese Option nur verwenden, wenn Sie die Option „Objektnotation“ verwenden.|
|skip|int|Gibt die Anzahl der Elemente in der Sammlung an, die übersprungen und nicht im Ergebnis miteinbezogen werden sollen. Wenn `top` angegeben ist, wird das Resultset gestartet, nachdem die angegebene Anzahl von Elementen übersprungen wurde. Optional. Sie können diese Option nur verwenden, wenn Sie die Option „Objektnotation“ verwenden.|

## Weitere Informationen

Die bevorzugte Methode zum Angeben der Eigenschaften und Paginginformationen ist das Zeichenfolgenliteral. In den folgenden zwei Beispielen wird die bevorzugte Methode zum Anfordern der Text- und Schriftgradeigenschaften für Absätze in einer Absatzsammlung dargestellt:

<code>context.load(paragraphs, 'text, font/size');</code>

<code>paragraphs.load('text, font/size');</code>

Hier sehen Sie ein ähnliches Beispiel mit Objektnotation (einschließlich Paging):

<code>context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>

<code>paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

Hinweis: Wenn keine bestimmten Eigenschaften des Schriftartenobjekts in der select-Anweisung angegeben sind, gibt die expand-Anweisung an, dass alle Schriftarteneigenschaften geladen werden.

## Beispiele

In diesem Beispiel werden die Absätze im Word-Dokument mit den dazugehörigen Text- und Schriftgradeigenschaften abgerufen.

```js
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the paragraphs collection.
            var paragraphs = context.document.body.paragraphs;

            // Queue a commmand to load the text and font properties.
            // It is best practice to always specify the property set. Otherwise, all properties are
            // returned in on the object.
            context.load(paragraphs, 'text, font/size');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            // Insert code that works with the paragraphs loaded by context.load().
           })
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });

```

## Supportdetails
Verwenden Sie den [Anforderungssatz](../office-add-in-requirement-sets.md) in Laufzeitüberprüfungen, um sicherzustellen, dass die Anwendung von der Hostversion von Word unterstützt wird. Weitere Informationen zur Office-Hostanwendung und den Serveranforderungen finden Sie unter [Anforderungen für die Ausführung von Office-Add-Ins](../../docs/overview/requirements-for-running-office-add-ins.md).
