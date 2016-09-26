# SearchOptions オブジェクト (JavaScript API for Word)

検索操作に含めるオプションを指定します。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ     | 型   |説明
|:---------------|:--------|:----------|
|ignorePunct|bool|単語間のすべての区切り記号を無視するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [句読点を無視する] チェック ボックスに相当します。|
|ignoreSpace|bool|単語間のすべての空白を無視するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [空白文字を無視する] チェック ボックスに相当します。|
|matchCase|bool|大文字と小文字を区別する検索を実行するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックス ([編集] メニュー) の [大文字と小文字を区別する] チェック ボックスに相当します。|
|matchPrefix|bool|検索文字列で始まる単語と一致するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [接頭辞に一致する] チェック ボックスに相当します。|
|matchSoundsLike|bool|**このオプションは 2016 年 6 月の更新で廃止されました**。 検索文字列と似ている単語を検出するかどうかを示す値を取得または設定します。 [検索と置換] ダイアログ ボックスの [あいまい検索] に相当します。|
|matchSuffix|bool|検索文字列で終わる語句と一致するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [接尾辞に一致する] に相当します。|
|matchWholeWord|bool|長い単語の一部ではなく、単語全体のみを検索操作の対象にするかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [完全に一致する単語だけを検索する] チェック ボックスに相当します。|
|matchWildCards|bool|特殊な検索演算子を使用して検索を実行するかどうかを示す値を取得または設定します。[検索と置換] ダイアログ ボックスの [ワイルドカードを使用する] チェック ボックスに相当します。|

_プロパティのアクセスの[例](#property-access-examples)を参照してください。_

検索のオプションは、省略可能です。検索のオプションは、すべての検索方法でオブジェクト リテラルを使用して指定する必要があります。

```js
    search('searchstring', {searchOption1:bool, ...searchOptionN:bool}
```

1 つ以上の検索オプションのプロパティをオブジェクト リテラルで指定して、検索オプションを指定できます。 

## 関係
なし


## メソッド

| メソッド           | 戻り値の型    |説明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。|

## メソッドの詳細

### load(param: object)
JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。

#### 構文
```js
object.load(param);
```

#### パラメーター
| パラメーター    | 型   |説明|
|:---------------|:--------|:----------|
|param|object|省略可能。パラメーターとリレーションシップ名を、区切られた文字列または 1 つの配列として受け入れます。あるいは、[loadOption](loadoption.md) オブジェクトを提供します。|

#### 戻り値
void

## プロパティのアクセスの例

### 句読点を無視する検索
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

### 接頭辞に基づく検索
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

### 接尾辞に基づく検索
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

### ワイルドカードを使用する検索
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


## ワイルドカードのガイダンス 

| 検索対象         | ワイルドカード |  サンプル |
|:-----------------|:--------|:----------|
| 任意の 1 文字| ? |s?t は、sat や set を検出します。 |
|文字からなる任意の文字列| * |s*d は、sad や started を検出します。|
|単語の先頭|< |<(inter) では、interesting や intercept が検出されますが、splintered は検出されません。|
|単語の末尾 |> |(in)> では、in や within が検出されますが、interesting は検出されません。|
|指定した文字のいずれか 1 つ|[ ] |w[io]n では、win と won が検出されます。|
|この範囲に含まれる任意の 1 文字| [-] |[r-t]ight では、right や sight が検出されます。範囲は、昇順にする必要があります。|
|角括弧で囲まれた範囲の文字を除く任意の 1 文字|[!x-z] |t[!a-m]ck では、tock や tuck が検出されますが、tack や tick は検出されません。|
|直前の文字または式の n 回の出現|{n} |fe\{2\}d では、feed が検出されますが、fed は検出されません。|
|直前の文字または式の n 回以上の出現|{n,} |fe{1,}d では、fed や feed が検出されます。|
|直前の文字または式の n 回から m 回までの出現|{n,m} |10{1,3} では、10、100、1000 が検出されます。|
|直前の文字または式の 1 回以上の出現|@ |lo@t では、lot や loot が検出されます。|


## サポートの詳細
実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。
