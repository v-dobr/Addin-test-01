# SearchOptions 对象（适用于 Word 的 JavaScript API）

指定要包括在搜索操作中的选项。

_适用于：Word 2016、Word for iPad、Word for Mac_

## 属性
| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|ignorePunct|bool|获取或设置指示是否忽略单词之间的所有标点符号的值。对应于“查找和替换”对话框中的“忽略标点符号”复选框。|
|ignoreSpace|bool|获取或设置指示是否忽略单词之间的所有空格的值。对应于“查找和替换”对话框中的“忽略空格字符”复选框。|
|matchCase|bool|获取或设置指示是否执行区分大小写的搜索的值。对应于“查找和替换”对话框（“编辑”菜单）中的“区分大小写”复选框。|
|matchPrefix|bool|获取或设置指示是否匹配以搜索字符串开头的单词。对应于“查找和替换”对话框中的“匹配前缀”复选框。|
|matchSoundsLike|bool|**此选项已在 2016 年 6 月更新中停用**。 获取或设置指示是否查找读音与搜索字符串类似的字词。 对应于“查找和替换”对话框中的“读音类似”复选框。|
|matchSuffix|bool|获取或设置指示是否匹配以搜索字符串结尾的单词。对应于“查找和替换”对话框中的“匹配后缀”复选框。|
|matchWholeWord|bool|获取或设置指示是否只查找整个单词，而不查找长单词的一部分的值。对应于“查找和替换”对话框中的“全字匹配”复选框。|
|matchWildCards|bool|获取或设置指示搜索是否使用特殊搜索操作符执行的值。对应于“查找和替换”对话框中的“使用通配符”复选框。|

_请参阅属性访问[示例](#property-access-examples)。_

搜索选项是可选的。应使用对象文本，在所有搜索方法中指定搜索选项：

```js
    search('searchstring', {searchOption1:bool, ...searchOptionN:bool}
```

您可以在对象文本中提供一个或多个搜索选项属性以指定搜索选项。 

## Relationships
无


## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息

### load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

## 属性访问示例

### 忽略标点符号搜索
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

### 基于前缀搜索
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

### 基于后缀搜索
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

### 使用通配符搜索
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


## 通配符指导 

| 若要查找：         | 通配符 |  示例 |
|:-----------------|:--------|:----------|
| 任意单个字符| ? |s?t 找到 sat 和 set。 |
|任何字符的字符串| * |s*d 找到 sad 和 started。|
|单词的开头|< |<(inter) 找到 interesting 和 intercept，而不是 splintered。|
|单词结尾 |> |(in)> 找到 in 和 within，而不是 interesting。|
|一个指定的字符|[ ] |w[io]n 找到 win 和 won。|
|此区域中的任何单个字符| [-] |[r-t]ight 找到 right 和 sight。区域必须按升序排列。|
|除括号中区域内的字符以外的任何单个字符|[!x-z] |t[!a-m]ck 找到 tock 和 tuck，而不是 tack 或 tick。|
|前一个字符或表达式出现 n 次|{n} |fe\{2\}d 找到 feed，而不是 fed。|
|前一个字符或表达式至少出现 n 次|{n,} |fe{1,}d 找到 fed 和 feed。|
|前一个字符或表达式出现 n 到 m 次|{n,m} |10{1,3} 找到 10、100 和 1000。|
|前一个字符或表达式出现一次或多次|@ |lo@t 找到 lot 和 loot。|


## 支持详细信息
在运行时检查过程中使用[要求设置](../office-add-in-requirement-sets.md)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](../../docs/overview/requirements-for-running-office-add-ins.md)。
