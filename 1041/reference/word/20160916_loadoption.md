# LoadOption オブジェクト (JavaScript API for Word)

context.sync() が呼び出されたときに読み込まれるページング情報とプロパティを指定するオブジェクト。

_適用対象:Word 2016、Word for iPad、Word for Mac_

## プロパティ
| プロパティ     | 型   |説明|
|:---------------|:--------|:----------|
|select|object|パラメーター/リレーションシップの名前のコンマ区切りリストまたは配列が含まれます。省略可能。|
|expand|object|リレーションシップ名のコンマ区切りリストまたは配列が含まれています。省略可能。|
|top|int| 結果に含めることができるコレクション項目の最大数を指定します。省略可能。このオプションは、オブジェクト表記オプションを使用する場合にのみ使用できます。|
|skip|int|スキップされて結果に組み込まれないコレクション内の項目の数を指定します。 `top` が指定されている場合は、指定された数の項目がスキップされた後で結果セットが開始されます。 省略可能。 このオプションは、オブジェクト表記オプションを使用する場合にのみ使用できます。|

## 詳細情報

プロパティとページング情報を指定するための推奨される方法は、文字列リテラルの使用です。最初の 2 つの例は、段落コレクションの段落のテキストおよびフォント サイズのプロパティを要求するための推奨される方法を示しています。

<code>context.load(paragraphs, 'text, font/size');</code>

<code>paragraphs.load('text, font/size');</code>

次に、オブジェクト表記 (ページングを含む) を使用する、類似の例を示します。

<code>context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>

<code>paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

select ステートメントのフォント オブジェクトで特定のプロパティを指定しない場合、すべてのフォント プロパティが読み込まれることを expand ステートメントが単独で示します。

## 例

この例は、テキストおよびフォント サイズのプロパティとともに Word 文書の段落を取得する方法を示しています。

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

## サポートの詳細
実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。
