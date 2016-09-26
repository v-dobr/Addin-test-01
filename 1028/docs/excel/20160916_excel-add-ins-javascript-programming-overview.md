# Excel JavaScript API 程式設計的概觀

本文說明如何使用 Excel JavaScript API 為 Excel 2016 建置增益集。 會介紹使用 API 的基礎關鍵概念，例如 RequestContext、JavaScript proxy 物件 sync()、Excel.run() 及 load()。 在文件結尾的程式碼範例顯示如何套用概念。

## RequestContext

RequestContext 物件可協助向 Excel 應用程式提出要求。由於 Office 增益集和 Excel 應用程式在兩個不同的處理程序中執行，因此需要使用要求內容以便從增益集存取 Excel 及相關的物件，例如工作表、表格等。建立的要求內容如下所示。

```js
var ctx = new Excel.RequestContext();
```

## Proxy 物件

在增益集中宣告和使用的 Excel JavaScript 物件，是 Excel 文件中實際物件的 proxy 物件。除非同步處理文件狀態，否則對 proxy 物件採取的所有動作都不會在 Excel 中實現，而 Excel 文件的狀態也不會在 proxy 物件中實現。執行 context.sync() 時就會同步處理文件狀態 (請參閱以下說明)。

例如，本機 JavaScript 物件 `selectedRange` 宣告為參考選取的範圍。 這可用來佇列其屬性設定以及叫用方法。 直到執行 sync() 方法後，才會實現對此類物件執行的動作。

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

要求內容可用的 sync() 方法會透過執行在內容中排入佇列的指示以及擷取已載入供程式碼使用之 Office 物件的屬性，來同步處理 JavaScript proxy 物件和 Excel 中實際物件之間的狀態。此方法會傳回承諾，同步處理完成時就會將其解決。

## Excel.run(function(context) { batch })

Excel.run() 會執行批次指令碼，以對 Excel 物件模型執行動作。批次命令包括本機 JavaScript proxy 物件的定義，以及同步處理本機、Excel 物件及承諾解決之間狀態的 sync() 方法。在 Excel.run() 中批次處理要求的優點是，當承諾已解決時，任何在執行期間所配置的追蹤 range 物件就會自動釋出。

Run 方法發生在 RequestContext 中，並傳回一項承諾 (通常就是 ctx.sync() 的結果)。也有可能在 Excel.run() 外部執行批次作業。不過，在這種情況下，將需要手動追蹤並管理 range 物件參考。

## load()

Load() 方法是用於填入在增益集 JavaScript 層中建立的 proxy 物件。舉例來說，當工作表嘗試擷取物件時，會先在 JavaScript 層中建立一個本機 proxy 物件。這種物件可用來佇列其屬性設定以及叫用方法。不過，若要讀取物件屬性或關聯，則需先叫用 load() 和 sync() 方法。Load() 方法會發生在屬性中以及在呼叫 sync() 方法時需要載入的關聯中。

_語法：_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
其中：

* `properties` 是要載入的屬性和/或關聯性名稱清單，以逗點分隔的字串或名稱陣列來指定。 如需詳細資訊，請參閱每個物件下方的 .load() 方法。
* `loadOption` 指定的物件用於描述 selection、expansion、top 和 skip 選項。 如需詳細資訊，請參閱物件載入[選項](../../reference/excel/loadoption.md)。

## 範例：從陣列中寫入值到範圍物件

下列範例會示範如何從陣列中寫入值到範圍物件。

Excel.run() 包含指示批次。 在此批次中會建立一個 proxy 物件，參考作用中工作表上的範圍 (位址 A1:B2)。 這個 proxy range 物件的值是在本機設定。 為了讀回此值，命令會指示將 range 的 `text` 屬性載入至 proxy 物件。 所有這些命令都會排入佇列，並在呼叫 ctx.sync() 後執行。 Sync() 方法會傳回一項可用來將其鏈結至其他作業的承諾。

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

## 範例：複製值

下列範例示範如何在 range 物件上使用 load() 方法，將作用中工作表的範圍 A1:A2 的值複製到 B1:B2。

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

## 屬性和關聯性的選取範圍

根據預設，object.load() 會選取所載入之物件的所有純量與複雜屬性。 預設不會載入關聯性 (在範例中，格式是 Range 物件的關聯性物件)。 不過，建議您明確標記要載入的屬性與關聯性，以便改善效能。 若要這樣做，(在 `load()` 參數中) 指定在回應中包含屬性和關聯性的子集。 Load 方法可接受兩種輸入：

* 以逗點分隔字串指定的屬性和關聯性名稱，_或_其中包含屬性或關聯性名稱的字串陣列。
* 描述 selection、expansion、top 和 skip 選項的物件。 如需詳細資訊，請參閱物件載入[選項](../../reference/excel/loadoption.md)。

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### 範例

下列 load 陳述式會載入 Range 的所有屬性，接著再展開 format 及 format/fill。

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

## Null 輸入

### 2 維陣列中的 null 輸入

更新 API 會略過二維陣列內部的 `null` 輸入 (針對值、數字格式、公式)。 當值、數字格式或值的公式方格中傳送了 `null` 輸入時，不會對預定目標進行任何更新。

範例：為了只更新 Range 的特定部分，例如某些儲存格的數字格式，並保留 Range 其他部分的現有數字格式，在所需位置設定想要的數字格式，並對其他儲存格傳送 `null`。

下列設定的要求中，只會設定某些部分的 Range 數字格式，其餘部分則保留現有的數字格式 (藉由傳遞 null)。

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### 屬性的 null 輸入

對於整個屬性而言，`null` 不是有效的單一輸入。 例如，以下命令無效，因為不能將整個值設定為 null 或略過。

```js
 range.values= null;

```

以下命令無效，因為 null 不是有效的色彩值。

```js
 range.format.fill.color =  null;
```

### Null 回應

包含不一致值的格式屬性表示法，會導致在回應中傳回 null 值。

範例：Range 可以包含多個儲存格的其中一個。如果所指定 Range 中包含的個別儲存格沒有統一的格式值，則範圍層級表示法將會是未定義。

```js
  "size" : null,
  "color" : null,
```

### 空白的輸入和輸出

更新要求中的空白值會被視為清除或重設各別屬性的指示。 空白值以中間沒有空格的兩個雙引號來表示。 `""`

範例：

* 若是 `values`，會清除範圍值。 此效果與清除應用程式的內容相同。

* 若是 `numberFormat`，會將數字格式設定為 `General`。

* 若是 `formula` 和 `formulaLocale`，會清除公式的值。


若是讀取作業，如果儲存格的內容為空白，則應該收到空白值。 如果儲存格不包含資料或值，API 就會傳回空白值。 空白值以中間沒有空格的兩個雙引號來表示。 `""`.

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## 未繫結的範圍

### 讀取

未繫結的範圍位址只包含欄識別碼而未指定列識別碼，或只包含列識別碼而未指定欄識別碼，例如：

* `C:C`、`A:F`、`A:XFD` (包含未指定的列)
* `2:2`、`1:4`、`1:1048546` (包含未指定的欄)

當 API 發出要求以擷取未繫結的範圍時 (例如 `getRange('C:C')`)，傳回的回應針對儲存格層級屬性 (例如 `values`、`text`、`numberFormat`、`formula` 等) 將包含 `null`。 其他 Range 屬性 (例如 `address`、`cellCount` 等) 則會反映未繫結的範圍。

### 寫入

**不允許**對未繫結的範圍設定儲存格層級屬性 (例如 values、numberFormat 等)，因為輸入要求可能會過大而無法處理。

範例：下列是不正確的更新要求，因為要求的範圍是未繫結的。

```js
...
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
    range.values = 'Due Date';
...
```

對這類範圍嘗試更新作業時，API 會傳回錯誤。


## 大型範圍

大型範圍表示該範圍的大小對單一 API 呼叫而言過大。 範圍中包含的許多因素 (例如儲存格數目、值、numberFormat 和公式) 都會使回應過大，而不適合與 API 互動。 API 會盡量嘗試傳回或寫入要求的資料。 但是，大小過大時可能會因為利用大量資源而導致 API 錯誤狀況。

若要避免這種狀況，建議將大型範圍分為多個較小型範圍來使用讀取或寫入作業。


## 單一輸入複本

為了支援以相同的值或數字格式來更新範圍或是將相同的公式套用至整個範圍，在設定 API中採用下列慣例。在 Excel 中，這個行為類似於以 CTRL+Enter 模式在一個範圍內輸入值或公式。

API 會尋找*單一儲存格值*，而如果目標範圍維度不符合輸入範圍維度，它會使用要求中提供的值或公式，以 CTRL+Enter 模式套用更新至整個範圍。

### 範例

下列要求會使用 "Due Date" 文字更新選取的範圍。注意 Range 有 20 個儲存格，而提供的輸入只有 1 個儲存格值。

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

下列要求會使用 '3/11/2015' 日期更新選取的範圍。

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
下列要求會用公式更新選定的範圍，方法是在 CTRL+Enter 模式中將公式套用至整個範圍。

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


## 錯誤訊息

錯誤

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |引數無效或遺失，或格式不正確。|
|InvalidRequest  |無法處理要求。|
|InvalidReference|此參考對目前的作業無效。|
|InvalidBinding  |由於先前的更新，此物件繫結不再有效。|
|InvalidSelection|目前的選取範圍對這項作業無效。|
|Unauthenticated |必要的驗證資訊遺失或無效。|
|AccessDenied   |您不能執行要求的作業。|
|ItemNotFound   |要求的資源不存在。|
|ActivityLimitReached|已達到活動上限。|
|GeneralException|處理要求時發生內部錯誤。|
|NotImplemented  |要求的功能未實作。|
|ServiceNotAvailable|服務無法使用。|
|Conflict	   |因為發生衝突，無法處理要求。|
|ItemAlreadyExists|欲建立的資源已經存在。|
|UnsupportedOperation|不支援所嘗試的操作。|
|RequestAborted|要求在執行階段中止。|
|ApiNotAvailable|無法使用要求的 API。|
|InsertDeleteConflict|嘗試的插入或刪除作業導致衝突。|
|InvalidOperation|在物件上嘗試的操作無效。|

## 其他資源

* [建立第一個 Excel 增益集](build-your-first-excel-add-in.md)
* [程式碼片段總管](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Excel 增益集程式碼範例](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel 增益集 JavaScript API 參考](excel-add-ins-javascript-api-reference.md)
