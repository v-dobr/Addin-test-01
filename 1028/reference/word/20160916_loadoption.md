# LoadOption 物件 (適用於 Word 的 JavaScript API)

一個物件，指定當呼叫 context.sync() 時要載入的分頁資訊和屬性。

_適用版本：Word 2016、Word for iPad、Word for Mac_

## 屬性
| 屬性	     | 類型	   |說明|
|:---------------|:--------|:----------|
|select|物件|包含參數/關聯性名稱的逗號分隔清單或陣列。選用。|
|expand|物件|包含關聯性名稱的逗號分隔清單或陣列。選用。|
|top|int| 指定結果中可包含的集合項目數上限。選用。當您使用物件標記選項時，您只能使用此選項。|
|略過|int|指定結果中要略過不予包含的集合項目數。 如果指定 `top`，則結果集會在略過指定的項目數後開始。 選用。 當您使用物件標記選項時，您只能使用此選項。|

## 詳細資訊

指定屬性和分頁資訊的慣用方法是使用字串常值。前兩個範例示範用來要求段落集合中的段落文字和字型大小屬性的偏好方式：

<code>context.load(paragraphs, 'text, font/size');</code>

<code>paragraphs.load('text, font/size');</code>

以下是使用物件標記法 (包括分頁) 的類似範例︰

<code>context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>

<code>paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

請注意，如果不在 select 陳述式中指定 font 物件的特定屬性，則 expand 陳述式本身會指定載入所有字型屬性。

## 範例

這個範例示範如何在 Word 文件中取得段落，以及其文字和字型大小屬性。

```js
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the paragraphs collection.
            var paragraphs = context.document.body.paragraphs;

            // Queue a commmand to load the text and font properties.
            // It is best practice to always specify the property set. Otherwise, all properties are
            // returned in on the object.
            context.load(paragraphs, 'text, font/size');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            // Insert code that works with the paragraphs loaded by context.load().
           })
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });

```

## 支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。
