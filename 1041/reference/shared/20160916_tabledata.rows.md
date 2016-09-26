
# TableData.rows プロパティ
テーブル内の行を取得または設定します。

|||
|:-----|:-----|
|**ホスト:**|Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|TableBindings|
|**で追加**|1.1|

```
var myRows = tableBindingObj.rows;
```


## 戻り値

テーブル内のデータを含む配列の配列を返します。行がない場合は、空の  **array**`[]` を返します。


## 解説

行を指定するには、テーブルの構造に対応した配列の配列を指定する必要があります。 たとえば、2 列で構成されるテーブルに **string** 値の 2 行を指定するには **rows** プロパティに ` [['a', 'b'], ['c', 'd']]` を設定します。

**rows** プロパティに **null** を指定した場合 (または **TableData** オブジェクトの作成時にこのプロパティを空のままにした場合)、コードを実行すると次の結果になります。


- 新しいテーブルを挿入した場合は、空行が挿入されます。
    
- 既存のテーブルを上書きまたは更新した場合は、既存の行は変更されません。
    

## 例

次の例では、ヘッダーと 3 つの行から成る単一列のテーブルを作成します。


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}
```


## サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|


|||
|:-----|:-----|
|**要件セットに指定できるもの**|TableBindings|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アプリの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|Word Online のサポートが追加されました。|
|1.1|Office for iPad における Excel と Word のサポートが追加されました。|
|1.0|導入|
