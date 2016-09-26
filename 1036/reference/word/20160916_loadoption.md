# Objet LoadOption (interface API JavaScript pour Word)

Objet permettant de définir les informations de pagination et les propriétés à charger lors de l’appel de la méthode context.sync().

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Propriétés
| Propriété     | Type   |Description|
|:---------------|:--------|:----------|
|select|object|Contient une liste délimitée par des virgules ou un tableau de noms de paramètres/relations. Facultatif.|
|expand|object|Contient une liste délimitée par des virgules ou un tableau de noms de relations. Facultatif.|
|top|int| Spécifie le nombre maximal d’éléments de collection qui peuvent être inclus dans le résultat. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.|
|skip|int|Indiquez le nombre d’éléments de la collection devant être ignorés et exclus du résultat. Si une valeur est définie pour `top`, la sélection du jeu de résultats démarre une fois que le nombre spécifié d’éléments a été ignoré. Facultatif. Vous pouvez utiliser cette option uniquement lorsque vous utilisez l’option de notation d’objet.|

## Plus d’informations

La méthode privilégiée pour spécifier les propriétés et les informations de pagination consiste à utiliser un littéral de chaîne. Les deux premiers exemples illustrent la méthode recommandée pour demander les propriétés de texte et de taille de police pour les paragraphes d’une collection :

<code>context.load(paragraphs, 'text, font/size');</code>

<code>paragraphs.load('text, font/size');</code>

Voici un exemple similaire qui utilise la notation d’objet (y compris la pagination) :

<code>context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>

<code>paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

Notez que si aucune propriété spécifique n’est définie pour l’objet de police dans l’instruction select, l’instruction expand, si elle est définie seule, indique que toutes les propriétés de police sont chargées.

## Exemples

Cet exemple montre comment obtenir les paragraphes du document Word, ainsi que les propriétés de texte et de taille de police.

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

## Informations de prise en charge
Utilisez l’[ensemble de conditions requises](../office-add-in-requirement-sets.md) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).
