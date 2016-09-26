# Общие сведения о создании кода с помощью API JavaScript для Excel

В этой статье описано, как создавать надстройки для Excel 2016 с помощью API JavaScript для Excel. В ней представлены основные понятия, такие как "RequestContext", "прокси-объекты JavaScript", "sync()", "Excel.run()" и "load()", имеющие ключевое значение при использовании интерфейсов API. Примеры кода в конце статьи иллюстрируют применение этих понятий.

## RequestContext

Объект RequestContext упрощает отправку запросов приложению Excel. Так как надстройка Office и приложение Excel — это два разных процесса, для доступа из надстройки к Excel и связанным объектам, например листам и таблицам, необходим контекст запроса. Контекст запроса создается, как показано ниже.

```js
var ctx = new Excel.RequestContext();
```

## Прокси-объекты

Объекты JavaScript Excel, объявленные и использованные в надстройке, — это прокси-объекты для реальных объектов в документе Excel. Никакие действия над прокси-объектами не реализуются в Excel, а состояние документа Excel не реализуется на прокси-объектах до его синхронизации. Состояние документа синхронизируется при выполнении метода context.sync(). (См. ниже).

Например, локальный объект JavaScript `selectedRange` объявлен в качестве ссылки на выбранный диапазон. Это можно использовать для постановки в очередь настройки его свойств и вызова методов. Действия над такими объектами не реализуются до выполнения метода sync().

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

Метод sync(), доступный в контексте запроса, синхронизирует состояние прокси-объектов JavaScript и реальных объектов в Excel путем выполнения поставленных в очередь инструкций над контекстом и получения свойств загруженных объектов Office для их использования в коде. Этот метод возвращает обещание, которое выполняется после завершения синхронизации.

## Excel.run(function(context) { batch })

Метод Excel.run() выполняет пакетный сценарий, выполняющий действия над моделью объекта Excel. Пакетные команды включают определения локальных прокси-объектов JavaScript и методов sync(), синхронизирующих состояние локальных объектов и объектов Excel, а также выполнение обещания. Преимущество пакетной обработки запросов в Excel.run() в том, что при выполнении обещания любые отслеживаемые объекты диапазона, выделенные во время выполнения, автоматически отпускаются.

Выполняемый метод использует объект RequestContext и возвращает обещание (как правило, просто результат метода ctx.sync()). Пакетную операцию можно выполнить вне метода Excel.run(). Однако при таком сценарии любые ссылки на объекты диапазона требуют отслеживания и управления вручную.

## load()

Метод load() используется для заполнения прокси-объектов, созданных на уровне JavaScript надстройки. При попытке получения объекта, например листа, сначала на уровне JavaScript создается локальный прокси-объект. Такой объект можно использовать для постановки в очередь настройки его свойств и методов вызова. Но для чтения свойств или связей объекта сначала необходимо вызвать методы load() и sync(). Метод load() использует свойства и связи, которые требуется загрузить при вызове метода sync().

_Синтаксис:_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
Где:

* `properties` — это список имен свойств и/или связей, которые требуется загрузить, разделенных запятыми, или массив имен. Дополнительные сведения см. в методах .load() под каждым объектом.
* `loadOption` указывает объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](../../reference/excel/loadoption.md).

## Пример. Запись значений из массива в объект диапазона

В приведенном ниже примере показано, как записать значения из массива в объект диапазона.

Метод Excel.run() содержит пакет инструкций. В рамках этого пакета создается прокси-объект, который ссылается на диапазон (адрес A1:B2) на активном листе. Значение этого прокси-объекта диапазона устанавливается локально. Чтобы прочитать значения, свойство `text` диапазона загружается в прокси-объект. Все эти команды ставятся в очередь и выполняются при вызове метода ctx.sync(). Метод sync() возвращает обещание, с помощью которого его можно связать с другими операциями.

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

## Пример. Копирование значений

В следующем примере показано, как скопировать значения из диапазона от A1:A2 до B1:B2 активного листа, используя метод load() на объекте диапазона.

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

## Выбор свойств и связей

По умолчанию метод object.load() выбирает все скалярные и сложные свойства загружаемого объекта. По умолчанию связи не загружаются (например, format — это объект связи объекта Range). Однако рекомендуем явно помечать загружаемые свойства и связи, чтобы повысить производительность. Для этого укажите (в параметре `load()`) подмножество свойств и связей, которые требуется включить в ответ. Метод load() поддерживает два типа входных данных:

* Имена свойств и связей, разделенные запятыми, _или_ массив имен.
* Объект, описывающий параметры "выбрать", "развернуть", "сверху" и "пропустить". Дополнительные сведения см. в статье, посвященной [параметрам загрузки объектов](../../reference/excel/loadoption.md).

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### Пример

Приведенный ниже оператор загружает все свойства объекта Range, а затем добавляет format и format/fill.

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

## Отсутствие входных данных

### Входное значение null в двумерном массиве

Входное значение `null` в двумерном массиве (для значений, числового формата, формулы) игнорируется в API обновления. Предполагаемый целевой объект не будет обновлен, если входное значение `null` отправлено в виде значений, числового формата или сетки значений формулы.

Пример. Чтобы обновить только определенные фрагменты диапазона, такие как числовой формат ячейки, и сохранить существующий числовой формат в других фрагментах диапазона, установите требуемый числовой формат в нужных фрагментах и отправьте значение `null` для других ячеек.

В приведенном ниже запросе значения задаются лишь для некоторых фрагментов числового формата диапазона, в то время как в остальных фрагментах сохраняется имеющийся числовой формат (передаются значения null).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### Входное значение null для свойства

`null` не является допустимым входным значением для всего свойства. Например, следующий пример недопустим, так как целые значения нельзя устанавливать на null или игнорировать.

```js
 range.values= null;

```

Следующий пример также недопустим, потому что значение null недопустимо для цвета.

```js
 range.format.fill.color =  null;
```

### Null-Response

Представление свойств форматирования, состоящее из неоднородных значений, приведет к возврату значения null в отклике.

Пример: диапазон может состоять из одной или нескольких ячеек. Если отдельные ячейки в указанном диапазоне не содержат однородных значений форматирования, представление уровня диапазона будет неопределенным.

```js
  "size" : null,
  "color" : null,
```

### Пустые входные и выходные данные

Пустые значения в запросах на обновление считаются указанием на очистку или сброс соответствующего свойства. Пустое значение представляется двумя двойными кавычками, не разделенными пробелом. `""`

Пример:

* Для `values` значение диапазона очищено. Это аналогично очистке содержимого в приложении.

* Для `numberFormat` числовому формату присвоено значение `General`.

* Для `formula` и `formulaLocale` значения формулы очищены.


При операциях чтения будьте готовы получать пустые значения, если в ячейках нет содержимого. Если ячейка не содержит данных или значений, API возвращает пустое значение. Пустое значение представляется двумя двойными кавычками, не разделенными пробелом. `""`.

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## Неограниченный диапазон

### Чтение

Адрес неограниченного диапазона содержит только идентификаторы столбцов и строк, а также идентификаторы неопределенных строк и столбцов (соответственно), например:

* `C:C`, `A:F`, `A:XFD` (содержит неопределенные строки)
* `2:2`, `1:4`, `1:1048546` (содержит неопределенные столбцы)

Когда API запрашивает неограниченный диапазон (например, `getRange('C:C')`), ответ содержит значение `null` для свойств на уровне ячеек, например `values`, `text`, `numberFormat`, `formula` и т. д. Другие свойства диапазона, такие как `address`, `cellCount` и т. д., отражают неограниченный диапазон.

### Запись

Задание свойств уровня ячеек (например, значений, числового формата и т. д.) для неограниченного диапазона **не допускается**, так как запрос на ввод может оказаться слишком большим для обработки.

Пример: приведенный ниже запрос на обновление значений недопустим, поскольку запрашиваемый диапазон не ограничен.

```js
...
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
    range.values = 'Due Date';
...
```

Когда для такого объекта Range предпринимается попытка выполнить операцию обновления, API возвращает ошибку.


## Большой диапазон

Большой диапазон — это объект Range, размер которого слишком велик для одного вызова API. Множество факторов, например количество свойств объекта, таких как cells, values, numberFormat и formula, могут сделать запрос настолько большим, что он станет неподходящим для взаимодействия с API. Интерфейс API делает все возможное для возврата запрашиваемых данных или записи в них. Но обработка крупного запроса может привести к ошибке API из-за чрезмерного использования ресурсов.

Чтобы избежать этого, рекомендуем выполнять операции чтения и записи для нескольких объектов Range меньшего размера.


## Копирование одного входного значения

Для поддержки обновления диапазона с использованием одинаковых значений или числового формата либо для применения одной и той же формулы ко всему диапазону в установленном интерфейсе API используется следующее соглашение. В Excel этот принцип аналогичен вводу значений или формул в диапазон в режиме CTRL+ВВОД.

API ищет *значение одной ячейки* и, если размер целевого диапазона не соответствует размеру входного диапазона, обновление применяется ко всему диапазону в режиме CTRL+ВВОД с использованием значения или формулы в запросе.

### Примеры

Приведенный ниже запрос добавляет в выбранный диапазон текст "Due Date". Обратите внимание, что диапазон содержит 20 ячеек, в то время как входные данные — значение лишь для одной ячейки.

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

Указанный ниже запрос добавляет в выбранный диапазон дату "11.03.2015".

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
Следующий запрос добавляет в выбранный диапазон формулу, которая применяется ко всему диапазону при нажатии клавиш CTRL+ВВОД.

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


## Сообщения об ошибках

Ошибки возвращаются с помощью объекта ошибки, состоящего из кода и сообщения. В таблице ниже перечислены возможные ошибки.

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |Аргумент недопустим, отсутствует или имеет неправильный формат.|
|InvalidRequest  |Не удается обработать запрос.|
|InvalidReference|Эта ссылка недопустима для текущей операции.|
|InvalidBinding  |Эта привязка объектов недопустима из-за предыдущих обновлений.|
|InvalidSelection|Выбранный фрагмент недопустим для этой операции.|
|Unauthenticated |Требуемые сведения о проверке подлинности отсутствуют или недопустимы.|
|AccessDenied   |Вы не можете выполнить запрашиваемую операцию.|
|ItemNotFound   |Запрашиваемый ресурс не существует.|
|ActivityLimitReached|Достигнут предел действий.|
|GeneralException|При обработке запроса возникла внутренняя ошибка.|
|NotImplemented  |Запрашиваемая функция не реализована.|
|ServiceNotAvailable|Служба недоступна.|
|Conflict   |Запрос не удалось обработать из-за конфликта.|
|ItemAlreadyExists|Создаваемый ресурс уже существует.|
|UnsupportedOperation|Выполняемая операция не поддерживается.|
|RequestAborted|Запрос прерван во время выполнения.|
|ApiNotAvailable|Запрашиваемый интерфейс API недоступен.|
|InsertDeleteConflict|Операция вставки или удаления привела к конфликту.|
|InvalidOperation|Выполняемая операция недопустима для этого объекта.|

## Дополнительные ресурсы

* [Создание первой надстройки Excel](build-your-first-excel-add-in.md)
* [Обозреватель фрагментов кода](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Примеры кода надстроек Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Справочник по API JavaScript для надстроек Excel](excel-add-ins-javascript-api-reference.md)
