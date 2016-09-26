
# TableData.rows 属性
获取或设置表中的行。

|||
|:-----|:-----|
|**主机：**|Excel 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|TableBindings|
|**在其中添加**|1.1|

```
var myRows = tableBindingObj.rows;
```


## 返回值

返回包含表中数据的数组的数组。如果没有行，则返回空的  **array**`[]`。


## 备注

若要指定行，则必须指定对应于表结构的数组的数组。 例如，若要在两列表中指定两行**字符串**值，需要将 **rows** 属性设置为 ` [['a', 'b'], ['c', 'd']]`。

如果您将  **rows** 指定为 **null**（或者在构造  **TableData** 对象时将该属性留空），执行代码时会出现以下结果：


- 如果插入新表，将插入一个空白行。
    
- 如果覆盖或更新现有表，则不会改动现有行。
    

## 示例

以下示例将创建具有一个标题和三行的单列表。


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}
```


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|


|||
|:-----|:-----|
|**在要求集中可用**|TableBindings|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Word Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel和 Word 的支持|
|1.0|引入|
