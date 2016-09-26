# Excel JavaScript API 编程概述

本文介绍如何使用 Excel JavaScript API 生成适用于 Excel 2016 的外接程序。 它介绍了使用 API（例如 RequestContext、JavaScript 代理对象、sync()、Excel.run() 和 load()）的基本关键概念。 文章结尾部分的代码示例介绍了如何应用这些概念。

## RequestContext

RequestContext 对象可加快对 Excel 应用程序的请求。由于 Office 外接程序和 Excel 应用程序在两个不同的进程中运行，需要请求上下文来获得对 Excel 及外接程序中相关对象（如工作表、表）的访问权限。如下所示创建请求上下文。

```js
var ctx = new Excel.RequestContext();
```

## 代理对象

在外接程序中声明和使用的 Excel JavaScript 对象是 Excel 文档中真实对象的代理对象。对代理对象执行的所有操作都不会在 Excel 中实现，Excel 文档的状态不会在代理对象中实现，直至文档状态已同步。运行 context.sync() 时将同步文档状态（参见下文）。

例如，本地 JavaScript 对象 `selectedRange` 声明为引用所选区域。 这可以用于对其属性和调用方法的设置进行排队。 对此类对象执行的操作不会实现，除非运行 sync() 方法。

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

请求上下文中可用的 sync() 方法通过执行在上下文中排队的指令以及检索用于您代码中的已加载 Office 对象的属性，在 JavaScript 代理对象和 Excel 中的真实对象之间同步状态。此方法返回一个将在同步完成时实现的承诺。

## Excel.run(function(context) { batch })

Excel.run() 运行一个对 Excel 对象模型执行操作的批处理脚本。批处理命令包括定义本地 JavaScript 代理对象、在本地和 Excel 对象之间同步状态的 sync() 方法以及承诺实现。Excel.run() 中的批处理请求的优势在于，当实现承诺时，在执行期间分配的任何被跟踪的 range 对象将会自动释放。

运行方法包括 RequestContext 并返回一个承诺（通常只是 ctx.sync() 的结果）。可以在 Excel.run() 之外运行批处理操作。但是，在这种情况下，任何 range 对象引用需要手动进行跟踪和管理。

## load()

load() 方法用于填充在外接程序 JavaScript 层中创建的代理对象。尝试检索对象（例如工作表）时，将首先在 JavaScript 层中创建一个本地代理对象。此类对象可以用于对其属性和调用方法的设置进行排队。但是，要读取对象属性或关系，则需首先调用 load() 和 sync() 方法。load() 方法包括在调用 sync() 方法时需加载的属性和关系。

_语法：_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
其中：

* `properties` 列出了要加载的属性和/或关系名称，指定为逗号分隔的字符串或名称数组。 有关详细信息，请参阅每个对象下的 .load() 方法。
* `loadOption` 指定的对象描述了选择、展开、置顶和跳过选项。 有关详细信息，请参阅对象加载 [选项](../../reference/excel/loadoption.md)。

## 示例：将数组中的值写入一个范围对象

以下示例介绍了如何将数组中的值写入一个范围对象。

Excel.run() 包含一批指令。 作为此批次的一部分，将会创建一个代码对象，引用活动工作表上的区域（地址 A1:B2）。 此代理 range 对象的值在本地设置。 为了读回值，区域的 `text` 属性被指示为加载到代理对象上。 所有这些命令将在调用 ctx.sync() 时排队和运行。 sync() 方法返回一个承诺，可用于将其与其他操作链接起来。

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

## 示例：复制值

下面的示例演示如何通过对 range 对象执行 load() 方法，将值从活动工作表的区域 A1:A2 复制到 B1:B2。

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

## 属性和关系选择

默认情况下，object.load() 选择正在加载的对象的所有标量和复杂属性。 默认情况下不会加载关系（例如，格式是 Range 对象的关系对象）。 但是，建议将属性和关系标记为明确加载以提高性能。 这可以通过（在 `load()` 参数中）指定要包括在响应中的一小部分属性和关系来实现。 加载方法允许两种类型的输入：

* 属性和关系名称，作为逗号分隔的字符串名称_或_作为包含属性或关系名称的字符串数组。
* 用于描述选择、展开、置顶和跳过选项的对象。 有关详细信息，请参阅对象加载 [选项](../../reference/excel/loadoption.md)。

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### 示例

以下加载语句将加载区域的所有属性，然后在格式和格式/填充上展开。

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

## Null-Input

### 二维数组中的 null 输入

二维数组中的 `null` 输入（对于值、数字格式、公式）将在更新的 API 中忽略。 当 `null` 输入以值的形式或在值的数字格式或公式网格中发送时，不会对预期目标进行更新。

示例：要仅更新区域的特定部分（例如某些单元格的数字格式）并保留区域其他部分的现有数字格式，请根据需要设置所需的数字格式并对其他单元格发送 `null`。

在以下设置请求中，仅会设置区域数字格式的某些部分，同时保留其余部分的现有数字格式（通过传递 null）。

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### 属性的 null 输入

`null` 并非整个属性的有效的单个输入。 例如，以下输入无效，因为整个值不能设置为 null，也不能忽略。

```js
 range.values= null;

```

以下输入无效，因为 null 不是有效的颜色值。

```js
 range.format.fill.color =  null;
```

### Null 响应

由不一致的值组成的格式属性的表示形式将导致在响应中返回 null 值。

示例：区域可以由一个或多个单元格组成。如果指定区域中包含的单个单元格不具有一致的格式值，则不会定义区域级别表示形式。

```js
  "size" : null,
  "color" : null,
```

### 空白输入和输出

更新请求中的空白值视为清除或重置相应属性的指令。 空白值表示为两个双引号，中间没有空格。 `""`

示例：

* 对于 `values`，将清除区域值。 这与清除应用程序中的内容相同。

* 对于 `numberFormat`，数字格式设置为 `General`。

* 对于 `formula` 和 `formulaLocale`，将清除公式值。


对于读取操作，预计单元格内容为空时会收到空白值。 如果单元格不包含数据或值，该 API 将返回空白值。 空白值表示为两个双引号，中间没有空格。 `""`（）。

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## 无限区域

### 读取

无限区域地址仅包含列或行标识符和未指定的行标识符或列标识符（分别），例如：

* `C:C`、`A:F`、`A:XFD`（包含未指定的行）
* `2:2`、`1:4`、`1:1048546`（包含未指定的列）

当 API 发出检索无限区域的请求时（例如 `getRange('C:C')`，返回的响应包含以下单元格级别属性的 `null`：例如 `values`、`text`、`numberFormat`、`formula` 等） 其他区域属性（例如 `address`、`cellCount` 等）将反映无限区域。

### 写入

**不允许**在无限区域上设置单元格级别属性（例如值、numberFormat 等），因为输入请求可能太大而无法处理。

示例：下面不是一个有效的更新请求，因为所请求的区域是无限的。

```js
...
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
    range.values = 'Due Date';
...
```

当尝试对此类区域执行更新操作时，API 将返回错误。


## 大区域

大区域意味着区域大小对于单个 API 调用来说太大。 区域中包含的很多因素（例如单元格数量、值、numberFormat 和公式）都会使响应太大，而不适用于 API 交互。 API 努力尝试返回或写入到请求的数据。 但是，涉及的大尺寸可能会导致由于资源利用率较高而产生 API 错误条件。

为了避免出现这种情况，建议以多个较小的区域大小对大区域执行读取或写入。


## 单个输入副本

为了支持使用相同的值或数字格式更新区域或在整个区域应用相同的公式，在一组 API 中使用以下约定。在 Excel 中，此行为与在 CTRL+Enter 模式下将值或公式输入到区域中相似。

API 将查找*单个单元格值*，如果目标区域尺寸与输入区域尺寸不符，它将在 CTRL+Enter 模式下，使用请求中提供的值或公式更新整个区域。

### 示例

以下请求使用文本“Due Date”更新选定的区域。请注意，区域包含 20 个单元格，而提供的输入仅包含一个单元格值。

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

以下请求使用日期“3/11/2015”更新选定的区域。

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
以下请求在 CTRL+Enter 模式下使用将应用于整个区域的公式来更新选定区域。

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


## 错误消息

使用包含代码和消息的错误对象返回错误。下表列出了可能发生的错误情况。

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |参数无效或缺少或格式不正确。|
|InvalidRequest  |无法处理此请求。|
|InvalidReference|此引用对于当前操作无效。|
|InvalidBinding  |由于之前的更新，此对象绑定不再有效。|
|InvalidSelection|当前选定内容对于此操作无效。|
|Unauthenticated |所需的身份验证信息缺少或无效。|
|AccessDenied   |无法执行所请求的操作。|
|ItemNotFound   |所请求的资源不存在。|
|ActivityLimitReached|已达到活动限制。|
|GeneralException|处理请求时出现内部错误。|
|NotImplemented  |所请求的功能未实现。|
|ServiceNotAvailable|服务不可用。|
|Conflict   |由于冲突，无法处理请求。|
|ItemAlreadyExists|所创建的资源已存在。|
|UnsupportedOperation|不支持正在尝试的操作。|
|RequestAborted|请求在运行时已中止。|
|ApiNotAvailable|请求的 API 不可用。|
|InsertDeleteConflict|尝试的插入或删除操作导致冲突。|
|InvalidOperation|尝试的操作对于对象无效。|

## 其他资源

* [构建你的第一个 Excel 外接程序](build-your-first-excel-add-in.md)
* [代码段资源管理器](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Excel 外接程序代码示例](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel 外接程序 JavaScript API 参考](excel-add-ins-javascript-api-reference.md)
