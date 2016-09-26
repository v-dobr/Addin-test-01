# Объект SearchOptions (API JavaScript для Word)

Указывает параметры, которые необходимо включить в операцию поиска.

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|ignorePunct|bool|Возвращает или задает значение, которое указывает, следует ли пропустить все знаки препинания между словами. Соответствует установленному флажку "Не учитывать знаки препинания" в диалоговом окне "Найти и заменить".|
|ignoreSpace|bool|Возвращает или задает значение, которое указывает, следует ли пропустить все пробелы между словами. Соответствует установленному флажку "Не учитывать пробелы" в диалоговом окне "Найти и заменить".|
|matchCase|bool|Возвращает или задает значение, которое указывает, следует ли выполнять поиск с учетом регистра. Соответствует установленному флажку "Учитывать регистр" в диалоговом окне "Найти и заменить".|
|matchPrefix|bool|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, начинающиеся со строки поиска. Соответствует установленному флажку "Учитывать префикс" в диалоговом окне "Найти и заменить".|
|matchSoundsLike|bool|**Этот параметр не поддерживается в обновлении за июнь 2016 года**. Возвращает или задает значение, которое указывает, нужно ли учитывать слова, имеющие схожее произношение со строкой поиска. Соответствует установленному флажку "Произносится как" в диалоговом окне "Найти и заменить".|
|matchSuffix|bool|Возвращает или задает значение, указывающее, нужно ли учитывать слова, которые заканчиваются строкой поиска. Соответствует установленному флажку "Учитывать суффикс" в диалоговом окне "Найти и заменить".|
|matchWholeWord|bool|Возвращает или задает значение, которое указывает, следует ли искать только целые слова, а не текст, являющийся частью большего слова. Соответствует установленному флажку "Только слово целиком" в диалоговом окне "Найти и заменить".|
|matchWildCards|bool|Возвращает или задает значение, которое указывает, будет ли выполняться поиск с использованием специальных операторов поиска. Соответствует установленному флажку "Подстановочные знаки" в диалоговом окне "Найти и заменить".|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

Параметры поиска являются необязательными. Параметры поиска должны быть указаны во всех методах поиска с помощью литерала объекта:

```js
    search('searchstring', {searchOption1:bool, ...searchOptionN:bool}
```

Чтобы задать параметры поиска, можно указать одно или несколько свойств параметров поиска в литерале объекта. 

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### load(param: object)
Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.

#### Синтаксис
```js
object.load(param);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void

## Примеры доступа к свойствам

### Поиск без учета знаков препинания
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

### Поиск на основе префикса
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

### Поиск на основе суффикса
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

### Поиск с использованием подстановочных знаков
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


## Руководство по подстановочным знакам 

| Чтобы найти:         | Подстановочный знак |  Пример |
|:-----------------|:--------|:----------|
| Любой знак| ? |"л?с" находит "лес" и "лис". |
|Любая строка знаков| * |"к*т" находит "кот" и "компот".|
|Начало слова|< |"<(интер)" находит "интересный" и "интермедия", но не "заинтересованный".|
|Конец слова |> |"(ель)>" находит "ель" и "портфель", но не "ельник".|
|Один из указанных знаков|[ ] |"п[оы]л" находит "пол" и "пыл".|
|Любой символ из этого диапазона| [-] |"[б-с]оль" находит "боль" и "соль". Диапазон должен быть указан в алфавитном порядке.|
|Любой символ, кроме символов из диапазона, указанного в скобках|[!э-я] |"ко[!а-п]а" находит "кора" и "коса", но не "коза" или "кожа".|
|Точное количество повторений (n) предыдущего знака или выражения|{n} |"жарен\{2\}ый" находит "жаренный", но не "жареный".|
|Количество повторений предыдущего знака или выражения не менее n раз|{n,} |"жарен{1,}ый" находит и "жареный" и "жаренный".|
|Количество повторений предыдущего знака или выражения в диапазоне от n до m|{n,m} |10{1,3} находит 10, 100 и 1000.|
|Одно или несколько повторений предыдущего знака или выражения|@ |"жарен@ый" находит "жареный" и "жаренный".|


## Сведения о поддержке
Используйте [набор требований](../office-add-in-requirement-sets.md) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](../../docs/overview/requirements-for-running-office-add-ins.md).
