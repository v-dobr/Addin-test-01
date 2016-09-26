# SearchOptions 物件 (適用於 Word 的 JavaScript API)

指定搜尋作業中要包含的選項。

_適用版本：Word 2016、Word for iPad、Word for Mac_

## 屬性
| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|ignorePunct|bool|取得或設定值，指出是否忽略文字之間的所有標點符號。相當於 [尋找及取代] 對話方塊中的 [略過標點符號] 核取方塊。|
|ignoreSpace|bool|取得或設定值，指出是否忽略文字之間的所有空格。相當於 [尋找及取代] 對話方塊中的 [略過空格字元] 核取方塊。|
|matchCase|bool|取得或設定值，指出是否執行區分大小寫的搜尋。相當於 [尋找及取代] 對話方塊 ([編輯] 功能表) 中的 [大小寫須相符] 核取方塊。|
|matchPrefix|bool|取得或設定值，指出是否比對符合搜尋字串開頭的文字。相當於 [尋找及取代] 對話方塊中的 [前置詞須相符] 核取方塊。|
|matchSoundsLike|bool|**此選項已在 2016 年 6 月更新中被取代**。 取得或設定值，指出是否尋找發音類似於搜尋字串的文字。 相當於 [尋找及取代] 對話方塊中的 [類似拼音 ] 核取方塊|
|matchSuffix|bool|取得或設定值，指出是否比對符合搜尋字串結尾的文字。相當於 [尋找及取代] 對話方塊中的 [後置詞須相符] 核取方塊。|
|matchWholeWord|bool|取得或設定值，指出是否只尋找整個字，而非屬於較長字詞一部分的文字。相當於 [尋找及取代] 對話方塊中的 [全字拼寫須相符] 核取方塊。|
|matchWildCards|bool|取得或設定值，指出是否使用特殊搜尋運算子來執行搜尋。相當於 [尋找及取代] 對話方塊中的 [使用萬用字元] 核取方塊。|

_請參閱屬性存取[範例。](#property-access-examples)_

搜尋選項是選用的。應使用物件常值在所有搜尋方法中指定搜尋選項：

```js
    search('searchstring', {searchOption1:bool, ...searchOptionN:bool}
```

您可以用物件常值提供一或多個搜尋選項屬性，來指定搜尋選項。 

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void

## 屬性存取範例

### 忽略標點符號搜尋
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

### 根據前置詞進行搜尋
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

### 根據後置詞進行搜尋
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

### 使用萬用字元搜尋
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


## 萬用字元指引 

| 尋找：         | 萬用字元 |  範例 |
|:-----------------|:--------|:----------|
| 任何單一字元| ? |s?t 可尋找 sat 和 set。 |
|任何字元字串| * |s*t 可尋找 sad 和 started。|
|文字的開頭|< |<(inter) 可尋找 interesting 和 intercept，但無法尋找 splintered。|
|文字的結尾 |> |(in)> 可尋找 in 和 within，但無法尋找 interesting。|
|其中一個指定字元|[ ] |w[io]n 可尋找 win 和 won。|
|此範圍內的任何單一字元| [-] |[r-t]ight 可尋找 right 和 sight。範圍必須以遞增順序排列。|
|方括弧內範圍中字元以外的任何單一字元|[!x-z] |t[!a-m]ck 可尋找 tock 和 tuck，但無法尋找 tack 和 tick。|
|前一字元或運算式的剛好 n 個發生次數|{n} |fe\{2\}d 可尋找 feed，但無法尋找 fed。|
|前一字元或運算式的至少 n 個發生次數|{n,} |fe{1,}d 可尋找 fed 和 feed。|
|前一字元或運算式的 n 到 m 個發生次數|{n,m} |10{1,3} 可尋找 10、100 和 1000。|
|前一字元或運算式的一或多發生次數|@ |lo@t 可尋找 lot 和 loot。|


## 支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。
