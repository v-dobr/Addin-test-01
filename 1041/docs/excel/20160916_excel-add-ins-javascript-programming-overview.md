# Excel の JavaScript API のプログラミングの概要

この記事では、Excel JavaScript API を使用して Excel 2016 のアドインをビルドする方法について説明します。 また、RequestContext、JavaScript プロキシ オブジェクト、sync()、Excel.run()、load() などの API を使用するために知っておくべき主な概念について紹介します。 この記事の末尾に掲載したコード例では、この概念を応用する方法を示します。

## RequestContext

RequestContext オブジェクトは、Excel アプリケーションへの要求を容易にします。Office アドインと Excel アプリケーションは 2 つの異なるプロセスで実行されているため、アドインから Excel やその関連オブジェクト (ワークシートやテーブル) にアクセスするには、要求のコンテキストが必要になります。要求のコンテキストは次のように作成されます。

```js
var ctx = new Excel.RequestContext();
```

## プロキシ オブジェクト

アドインで宣言され使用される Excel の JavaScript オブジェクトは、Excel ドキュメントの実際のオブジェクトのプロキシ オブジェクトになります。プロキシ オブジェクトで実行されるすべてのアクションは、Excel では認識されません。また、Excel ドキュメントの状態は、ドキュメントの状態が同期されるまでプロキシ オブジェクトで認識されません。ドキュメントの状態は、context.sync() の実行時に同期されます (以下を参照)。

たとえば、ローカルの JavaScript オブジェクト `selectedRange` は、選択された範囲を参照するように宣言されています。 これは、このオブジェクトのプロパティと呼び出しメソッドの設定をキューに入れるために使用できます。 Sync() メソッドが実行されるまで、これらのオブジェクトのアクションは認識されません。

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

要求のコンテキストで使用可能な Sync() メソッドは、JavaScript のプロキシ オブジェクトと Excel の実際のオブジェクトの間で状態を同期させます。これは、このコンテキストでキューに入れられた命令を実行し、コードで使用するために読み込まれた Office オブジェクトのプロパティを取得することによって行われます。このメソッドは、同期処理が完了したときに解決される約束を返します。

## Excel.run(function(context) { batch })

Excel.run() は、Excel オブジェクト モデルに対してアクションを実行するバッチ スクリプトを実行します。このバッチ コマンドには、JavaScript のローカル プロキシ オブジェクトの定義と、ローカル オブジェクトと Excel オブジェクトの間で状態を同期し、解決される約束を返す sync() メソッドが含まれます。Excel.run() で要求をバッチ処理する利点は、約束が解決されるときに、実行中に割り当てられたすべての追跡範囲オブジェクトが自動的に解放されることです。

run メソッドは、RequestContext を取り込み、約束  (通常は、単なる ctx.sync() の結果) を返します。バッチ操作は Excel.run() の外部で実行することができます。ただし、このようなシナリオでは、範囲オブジェクトの参照は、手動で追跡および管理する必要があります。

## load()

load() メソッドは、アドインの JavaScript レイヤーで作成されたプロキシ オブジェクトに設定を取り込むために使用されます。オブジェクト、たとえばワークシート、を取得しようとすると、まず JavaScript レイヤーでローカル プロキシ オブジェクトが作成されます。このようなオブジェクトは、そのプロパティと呼び出しメソッドの設定をキューに登録するために使用できます。しかし、オブジェクトのプロパティや関係を読み取るためには、最初に load() メソッドと sync() メソッドを呼び出す必要があります。load() メソッドは、sync() メソッドが呼び出されたときに読み込まれる必要があるプロパティと関係を取り込みます。

_構文:_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
ここで

* `properties` は、読み込まれるプロパティ名やリレーションシップ名の一覧で、名前のコンマ区切りの文字列または配列として指定されます。 詳細は、各オブジェクトの下の .load() メソッドを参照してください。
* `loadOption` は、selection、expansion、top、skip の各オプションについて説明するオブジェクトを指定します。 詳細については、オブジェクトの読み込みの[オプション](../../reference/excel/loadoption.md)を参照してください。

## 例: 配列の値を範囲オブジェクトに書き込む

次の例は、配列の値を範囲オブジェクトに書き込む方法を示しています。

Excel.run() には、命令のバッチが含まれています。 このバッチの一部として、作業中のワークシートの範囲 (アドレス A1:B2) を参照するプロキシ オブジェクトが作成されます。 この範囲のプロキシ オブジェクトの値は、ローカルに設定されています。 値を読み取って返すようにするため、この範囲の `text` プロパティが、プロキシ オブジェクトに読み込まれるように指示されています。 これらのすべてのコマンドがキューに入れられ、ctx.sync() が呼び出されたときに実行されます。 sync() メソッドが返す約束は、このメソッドを他の操作とチェーンにするために使用できます。

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

## 例: 値をコピーする

次の例は、範囲オブジェクトで load() メソッドを使用して、作業中のワークシートの A1:A2 から B1:B2 までの範囲の値をコピーする方法を示しています。

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

## プロパティとリレーションシップの選択

既定では、object.load() は、読み込まれるオブジェクトのスカラー プロパティと複合プロパティをすべて選択します。 リレーションシップは既定では読み込まれません (たとえば、書式は範囲オブジェクトのリレーションシップ オブジェクトです)。 ただし、パフォーマンスの向上のために、プロパティと関係が読み込まれるように明示的にマークすることをお勧めします。 これを行うには、(`load()` パラメーターで) プロパティとリレーションシップのサブセットを応答に含めるように指定します。 load メソッドは、次の 2 種類の入力を受け入れます。

* プロパティ名とリレーションシップ名をコンマ区切りの文字列の名前として入力するか、_または_プロパティやリレーションシップの名前を含んだ文字列の配列として入力することができます。
* selection、expansion、top、skip の各オプションについて説明するオブジェクト。 詳細については、オブジェクトの読み込みの[オプション](../../reference/excel/loadoption.md)を参照してください。

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### 例

次の load ステートメントでは、範囲のプロパティをすべて読み込んでから、書式と塗りつぶしの書式を拡張しています。

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

## null の入力

### 2 次元配列での null の入力

更新 API では、2 次元配列内の (値、番号書式、数式に対する) `null` の入力は無視されます。 値や値の番号書式または数式のグリッドに `null` の入力を送信する場合、指定の対象に対しては更新が行われません。

例:範囲の特定の部分 (一部のセルの番号書式など) のみを更新し、範囲のその他の部分では既存の番号書式を保持する場合は、必要な部分で希望する番号書式を設定し、他のセルに対しては `null` を送信します。

次の設定要求では、範囲内のある部分の番号書式のみを設定し、残りの部分では (null 値を渡すことで) 既存の番号書式を保持します。

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### プロパティに対する null の入力

`null` を、プロパティ全体に対する単独の入力として指定することはできません。 たとえば、値全体を null に設定したり無視したりすることはできないため、以下は無効になります。

```js
 range.values= null;

```

null は有効なカラー値ではないため、以下も無効になります。

```js
 range.format.fill.color =  null;
```

### null 応答

均一でない値で構成された書式設定プロパティの表示形式では、null 値が応答で返されます。

例:範囲は、1 つ以上のセルで構成できます。指定した範囲に含まれる個々のセルの書式設定値が均一でない場合、その範囲のレベルの表示形式は定義されません。

```js
  "size" : null,
  "color" : null,
```

### 空の入力と出力

更新要求にある空の値は、それぞれのプロパティをクリアまたはリセットする命令として扱われます。 空の値は、間にスペースを入れない 2 つの二重引用符によって表されます。 `""`

例:

* `values` の場合は、範囲の値がクリアされています。 これは、アプリケーションの内容をクリアするのと同じです。

* `numberFormat` の場合は、番号書式が `General` に設定されます。

* `formula` および `formulaLocale` の場合は、数式の値がクリアされます。


読み取り操作では、セルの内容が空白の場合に空白の値を受け取ることが予想されます。 セルにデータや値が含まれていない場合、API は 空の値を返します。 空の値は、間にスペースを入れない 2 つの二重引用符によって表されます。 `""`

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## 無制限の範囲

### 読み取り

無制限の範囲のアドレスは、列識別子のみがあって行識別子が未指定であるか、あるいは行識別子のみがあって列識別子が未指定です。たとえば、次のとおりです。

* `C:C`、`A:F`、`A:XFD` (行が未指定)
* `2:2`、`1:4`、`1:1048546` (列が未指定)

API が無制限のセル範囲を取得する要求を行う場合 (`getRange('C:C')` など)、返される応答では、`values`、`text`、`numberFormat`、`formula` などのセル レベルのプロパティに `null` が含まれます。 `address`、`cellCount` などその他の範囲プロパティは、無制限の範囲を反映します。

### 書き込み

無制限のセル範囲にセル レベルのプロパティ ( values、numberFormat など) を設定することは、入力要求が長すぎて処理できない可能性があるため、**許可されていません**。

例: 要求された範囲が無制限であるため、次の更新要求は無効です。

```js
...
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
    range.values = 'Due Date';
...
```

このような範囲に対して更新操作を実行しようとすると、API はエラーを返します。


## 広い範囲

広い範囲とは、1 つの API の呼び出しに対してサイズが大きすぎる範囲を意味します。 範囲に含まれる、セル数、値、番号書式、数式などの多くの要因によって、応答のサイズが大きくなりすぎて API での操作に適さなくなることがあります。 API は、要求されたデータを返したり、それに書き込んだりしようと最善を尽くします。 しかし、大きなサイズが関係していると、リソース使用率が大きくなるために、API エラー状態になることがあります。

これを回避するには、広い範囲の読み取りや書き込みは、よりサイズの小さい複数の範囲に分けて実行することをお勧めします。


## 単一の入力のコピー

同じ値または番号書式での範囲の更新や、範囲全体への同じ数式の適用をサポートするため、set API では以下の方法が用いられています。Excel では、この動作は、Ctrl+Enter モードで範囲に値や数式を入力することに似ています。

API は *1 つのセル値*を探し、対象の範囲ディメンションが入力の範囲ディメンションと一致しない場合は、CTRL+Enter モードで範囲全体を、要求で指定された値または数式で更新します。

### 例

次の要求では、"Due Date" というテキストで選択範囲が更新されます。範囲に 20 個のセルがある一方、指定された入力には 1 つのセルの値のみがあることに注意してください。

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

次の要求では、選択範囲の日付が 2015 年 3 月 11 日に更新されます。

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
次の要求では、CTRL+Enter モードで範囲全体に適用される数式で選択範囲が更新されます。

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


## エラー メッセージ

エラーは、コードとメッセージで構成される error オブジェクトを使用して返されます。次の表は、発生する可能性があるエラー状態の一覧を示しています。

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |引数が無効であるか、存在しません。または形式が正しくありません。|
|InvalidRequest  |要求を処理できません。|
|InvalidReference|この参照は、現在の操作に対して無効です。|
|InvalidBinding  |このオブジェクトのバインドは、以前の更新プログラムが原因で無効になっています。|
|InvalidSelection|現在の選択内容は、この操作では無効です。|
|Unauthenticated |必要な認証情報が見つからないか、無効です。|
|AccessDenied   |要求された操作を実行できません。|
|ItemNotFound   |要求されたリソースは存在しません。|
|ActivityLimitReached|アクティビティの制限に達しました。|
|GeneralException|要求の処理中に内部エラーが発生しました。|
|NotImplemented  |要求された機能は実装されていません。|
|ServiceNotAvailable|サービスを利用できません。|
|Conflict   |競合のため、要求を処理できませんでした。|
|ItemAlreadyExists|作成中のリソースはすでに存在しています。|
|UnsupportedOperation|試行中の操作はサポートされていません。|
|RequestAborted|実行時に要求が中止されました。|
|ApiNotAvailable|要求された API は使用できません。|
|InsertDeleteConflict|試行された挿入操作または削除操作で競合が発生しました。|
|InvalidOperation|試行された操作はこのオブジェクトでは無効です。|

## その他のリソース

* [最初の Excel アドインをビルドする](build-your-first-excel-add-in.md)
* [コード スニペット エクスプローラー](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Excel アドインのコード サンプル](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel アドインの JavaScript API リファレンス](excel-add-ins-javascript-api-reference.md)
