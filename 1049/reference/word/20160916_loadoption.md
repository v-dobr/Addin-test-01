# Объект LoadOption (API JavaScript для Word)

Объект, определяющий сведения о разбивке по страницам и свойства для загрузки при вызове context.sync().

_Свойства_

## Свойства
| Свойство     | Тип   |Описание|
|:---------------|:--------|:----------|
|select|object|Содержит массив или разделенный запятыми список имен параметров и связей. Необязательный параметр.|
|expand|object|Содержит массив или разделенный запятыми список имен связей. Необязательный параметр.|
|top|int| Указывает максимальное число элементов в коллекции, которые можно включить в результат. Необязательный параметр. Его можно применять, только если используется параметр нотации объектов.|
|skip|int|Укажите количество элементов в коллекции, которые необходимо пропустить и исключить из результата. Если указан параметр `top`, результирующий набор начнется после пропуска заданного числа элементов. Необязательный. Его можно применять, только если используется параметр нотации объектов.|

## Подробнее

Для указания свойств и сведений о разбивке на страницы рекомендуется использовать строковый литерал. В первых двух примерах показан предпочтительный способ запроса свойств размера текста и шрифта для абзацев в коллекции абзацев:

<code>context.load(paragraphs, 'text, font/size');</code>

<code>paragraphs.load('text, font/size');</code>

Вот аналогичный пример с использованием нотации объектов (включающий подкачку):

<code>context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>

<code>paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

Обратите внимание, что если не задать определенные свойства объекта шрифта в инструкцию select, инструкция expand сама по себе означает, что загружаются все свойства шрифта.

## Примеры

В этом примере показано, как получить абзацы в документе Word, а также свойства размера текста и шрифта для них.

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

## Сведения о поддержке
Используйте [набор требований](../office-add-in-requirement-sets.md) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](../../docs/overview/requirements-for-running-office-add-ins.md).
