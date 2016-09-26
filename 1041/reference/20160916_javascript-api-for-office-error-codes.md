
# JavaScript API for Office のエラー コード
この記事では、JavaScript API for Office (Office.js) の使用時に発生する可能性のあるエラー メッセージについて説明します。

 _**適用対象:**Office アドイン | SharePoint アドイン | Excel | Outlook | PowerPoint | Project | Word_


## エラー コード

次の表に、エラー コード、名前、表示されるメッセージ、それらが示す状態を示します。



|**[Error.code](../reference/shared/error.code.md)**|**[Error.name](../reference/shared/error.name.md)**|**[Error.message](../reference/shared/error.message.md)**|**条件**|
|:-----|:-----|:-----|:-----|
|1000|無効な強制型変換|指定された強制型変換はサポートされていません。|ホスト アプリケーションではこの強制型変換はサポートされていません (たとえば、OOXML と HTML の強制型変換タイプは Excel でサポートされていません)。|
|1001|データの読み取りエラー|現在の選択項目はサポートされていません。|ユーザーの現在の選択項目はサポートされていません (つまり、サポートされている強制型変換と異なっている部分があります)。|
|1002|無効な強制型変換|指定された強制型変換は、このバインド タイプと互換性がありません。|ソリューション開発者が指定した強制型変換とバインド タイプの組み合わせには互換性がありません。|
|1003|データの読み取りエラー|指定した rowCount または columnCount の値が無効です。|ユーザーが無効な列数または行数を指定しています。|
|1004|データの読み取りエラー|現在の選択項目は指定された強制型変換と互換性がありません。|このアプリケーションでは、現在の選択項目は指定された強制型変換でサポートされていません。|
|1005|データの読み取りエラー|指定された startRow または startColumn の値が正しくありません。|ユーザーが無効な startRow または startCol の値を指定しています。|
|1006|データの読み取りエラー|テーブルに結合されたセルが含まれている場合、座標パラメーターを強制型変換タイプ "Table" と共に使用できません。|ユーザーは一様でないテーブル (つまり、マージされたセルを持つテーブル) から一部のデータを取得しようとしています。 |
|1007|データの読み取りエラー|ドキュメントのサイズが大きすぎます。|ユーザーが、現在サポートされているサイズより大きいドキュメントを取得しようとしています。|
|1008|データの読み取りエラー|要求されたデータ セットが大きすぎます。|ユーザーがホスト アドインで定義されたデータの制限を超えるデータの読み取りを要求しています。|
|1009|データの読み取りエラー|指定されたファイルの種類はサポートされていません。|ユーザーが、無効なファイルの種類を送信しています。|
|2000|データの書き込みエラー|指定されたデータ オブジェクトの型はサポートされていません。 |サポートされていないデータ オブジェクトが指定されています。|
|2001|データの書き込みエラー|現在の選択項目を書き込むことができません。|ユーザーの現在の選択項目は、書き込み操作でサポートされていません (ユーザーがイメージを選択した場合など)。|
|2002|データの書き込みエラー|指定されたデータ オブジェクトは、現在の選択項目の形状または次元と互換性がありません。|複数のセルが選択されています (また、選択項目の形状がデータの形状と一致しません)。複数のセルが選択されています (また、選択項目の次元がデータの次元と一致しません)。|
|2003|データの書き込みエラー|指定されたデータ オブジェクトがデータを上書きするため、設定操作に失敗しました。|1 つのセルが選択され、指定されたデータ オブジェクトが、ワークシート内のデータを上書きします。|
|2004|データの書き込みエラー|指定されたデータ オブジェクトが、現在の選択項目のサイズと一致しません。|ユーザーが、現在の選択項目のサイズよりも大きいオブジェクトを指定しています。|
|2005|データの書き込みエラー|指定された startRow または startColumn の値が正しくありません。|ユーザーが無効な startRow または startCol の値を指定しています。|
|2006|無効な形式のエラー|指定されたデータ オブジェクトの形式が正しくありません。|ソリューション開発者が、HTML または OOXML の無効な文字列、HTML の不正な文字列、または OOXML の無効な文字列を指定しています。|
|2007|無効なデータ オブジェクト|指定されたデータ オブジェクトの型は、現在の選択項目と互換性がありません。|ソリューション開発者が、指定された強制型変換と互換性のないデータ オブジェクトを指定しています。|
|2008|データの書き込みエラー|TBD|TBD|
|2009|データの書き込みエラー|指定されたデータ オブジェクトが大きすぎます。|ユーザーがホスト アドインで定義されたデータの制限を超えたデータを設定しようとしました。|
|2010|データの書き込みエラー|テーブルに結合されたセルが含まれている場合は、座標パラメーターを強制変換タイプ Table と共に使用できません。|ユーザーが一様でないテーブル (つまり、マージされたセルを持つテーブル) から一部のデータを設定しようとしています。|
|3000|バインディングの作成エラー|現在の選択項目をバインドできません。|ユーザーの選択項目は、バインディングではサポートされていません (たとえば、ユーザーがイメージまたはその他のサポートされていないオブジェクトを選択しています)。|
|3001|バインディングの作成エラー|TBD|TBD|
|3002|無効なバインド エラー|指定されたバインドが存在しません。|開発者は、存在しない、または削除されたバインディングにバインドしようとしています。|
|3003|バインディングの作成エラー|連続していない選択項目はサポートされません。|ユーザーが複数の選択を行っています。|
|3004|バインディングの作成エラー|現在の選択項目と指定されたバインド タイプでバインドを作成できません。|これが発生する条件はいくつかあります。この資料の後半の「バインディングの作成エラーの条件」セクションを参照してください。|
|3005|無効なバインド操作|このバインド タイプではサポートされていない操作です。|開発者が _テーブル_ ではないバインド タイプに行の追加操作または列の追加操作を送信しています。|
|3006|バインディングの作成エラー|名前付きアイテムが存在しません。|名前付きアイテムが見つかりませんでした。 その名前を持つコンテンツ コントロールまたはテーブルが存在しません。|
|3007|バインディングの作成エラー|同じ名前を持つ複数のオブジェクトが見つかりました。|競合エラー: 同じ名前を持つコンテンツ コントロールが複数存在し、競合の失敗が  **true** に設定されています。|
|3008|バインディングの作成エラー|指定されたバインド タイプは、指定された名前付きアイテムと互換性がありません。|名前付きアイテムは型にバインドできません。たとえば、コンテンツ コントロールにテキストが含まれていますが、開発者が型変換タイプ  _table_ を使用してバインドしようとしています。|
|3009|無効なバインド操作|バインド タイプがサポートされていません。|下位互換性のために使用されます。|
|3010|サポートされないバインド操作|選択するコンテンツはテーブル形式にする必要があります。 データをテーブルとして書式設定して、もう一度やり直してください。|開発者が強制型変換タイプ  **matrix** のデータ上の **TableBinding** オブジェクトの **addRowsAsynch** メソッドまたは _deleteAllDataValuesAsynch_ メソッドを使用しようとしています。|
|4000|設定の読み取りエラー|指定された設定の名前が存在しません。|存在しない設定の名前が指定されています。|
|4001|設定の保存エラー|設定を保存できませんでした。|設定を保存できませんでした。|
|4002|古い設定のエラー|設定が古いために保存できませんでした。|設定が古く、開発者が設定を上書きしないよう指定しています。|
|5000|古い設定のエラー|この操作はサポートされていません。|この操作は現在のホストではサポートされていません。たとえば、 **document.getSelectionAsync** が Outlook から呼び出されています。|
|5001|内部エラー|内部エラーが発生しました。|内部エラー条件は、次のいずれかの理由で発生する可能性があります。<br/><table><tr><td>ブックを共有している他のユーザーが使用しているアドインが、ほとんど同時にバインドを作成しました。使用しているアドインは、再バインドを行う必要があります。</tr></td><tr><td>不明なエラーが発生しました。</tr></td><tr><td>処理に失敗しました。</tr></td><tr><td>ユーザーが権限を持つロールのメンバーではないために、アクセスが拒否されました。</tr></td><tr><td>セキュリティで保護された、暗号化された通信が必要なために、アクセスが拒否されました。</tr></td><tr><td>データが古いので、クエリがデータを再取得できるよう確認する必要があります。</tr></td><tr><td>サイト コレクションの CPU クォータが限界を超えています。</tr></td><tr><td>サイト コレクションのメモリ クォータが限界を超えています。</tr></td><tr><td>セッションのメモリ クォータが限界を超えています。</tr></td><tr><td>ブックが無効な状態なので、操作を実行できません。</tr></td><tr><td>アイドル状態が続いてセッションがタイムアウトしました。ユーザーがブックを再読み込みする必要があります。</tr></td><tr><td>ユーザーごとに許可されるセッションの最大数を超えています。</tr></td><tr><td>操作はユーザーによって取り消されました。</tr></td><tr><td>時間がかかりすぎているため、操作を完了できません。</tr></td><tr><td>要求を完了できません。再試行する必要があります。</tr></td><tr><td>製品の試用期間の期限が切れています。</tr></td><tr><td>アイドル状態が続いたのでセッションがタイムアウトしました。</tr></td><tr><td>ユーザーは指定されたセル範囲に対する操作を実行する権限がありません。</tr></td><tr><td>現在のコラボレーションのセッションとユーザーの地域の設定が一致しません。</tr></td><tr><td>ユーザーはもはや接続されていません。ブックを更新し再度開く必要があります。</tr></td><tr><td>要求した範囲がシートに存在しません。</tr></td><tr><td>ユーザーは、ブックを編集する権限がありません。</tr></td><tr><td>ブックはロックされているので、編集できません。</tr></td><tr><td>セッションは、ブックを自動的に保存できません。</tr></td><tr><td>セッションは、ブック ファイルのロックを更新できません。</tr></td><tr><td>要求を処理できません。再試行する必要があります。</tr></td><tr><td>ユーザーのサインイン情報を検証できませんでした。再入力する必要があります。</tr></td><tr><td>ユーザーのアクセスが拒否されています。</tr></td><tr><td>共有ブックを更新する必要があります。</tr></td></table>|
|5002|アクセスが拒否されました|要求された操作は、現在のドキュメント モードでは許可されません。|ソリューション開発者が設定操作を送信しましたが、ドキュメントが "編集の制限" など、変更を許可しないモードになっています。|
|5003|イベント登録エラー|指定されたイベントの種類は、現在のオブジェクトではサポートされていません。|ソリューション開発者が、存在しないイベントにハンドラーを登録または登録解除しようとしています。|
|5004|無効な API 呼び出し|現在のコンテキストで無効な API 呼び出しです。|Excel で  **CustomXMLPart** オブジェクトを使用しようとするなど、コンテキストに対して無効な呼び出しが行われています。|
|5005|データが古い|サーバー上のデータが古いため、操作が失敗しました。|サーバー上のデータを更新する必要があります。|
|5006|セッションのタイムアウト|ドキュメント セッションがタイムアウトしました。 ドキュメントを再読み込みします。 |セッションがタイムアウトになりました。|
|5007|無効な API 呼び出し|列挙体は、現在のコンテキストではサポートされていません。|列挙体は、現在のコンテキストではサポートされていません。|
|5009|アクセスが拒否されました|アクセスが拒否されました|アドインに特定の API を呼び出すためのアクセス許可がありません。|
|6000|無効なノード|指定されたノードが見つかりませんでした。|**CustomXmlPart** ノードが見つかりませんでした。|
|6100|カスタム XML エラー|カスタム XML エラー|無効な API 呼び出し|
|7000|無効な ID|指定された ID が存在しません。|無効な ID|
|7001|無効なナビゲーション|ナビゲーションがサポートされていない場所にオブジェクトがあります。|ユーザーはオブジェクトを見つけることができますが、ナビゲーションできません (たとえば、Word でヘッダー、フッター、またはコメントにバインドされています)。|
|7002|無効なナビゲーション|オブジェクトがロックされているか、保護されています。|ロックまたは保護された範囲へ移動しようとしています。|
|7004|無効なナビゲーション|インデックスが範囲を超えているため、操作に失敗しました。|範囲外のインデックスに移動しようとしています。|
|8000|パラメーターがありません|一部のパラメーターの値が存在しないため、テーブルのセルの書式を設定できませんでした。 パラメーターを確認して、もう一度やり直してください。|cellFormat メソッドに一部のパラメーターがありません。 たとえば、cells、format、または tableOptions パラメーターがありません。|
|8010|無効な値|1 つ以上の cells パラメーター値が使用できません。値を確認して、もう一度やり直してください。|一般的なセル参照列挙型が定義されていません。 たとえば、All、Data、Headers などです。|
|8011|無効な値|1 つ以上の tableOptions パラメーター値が使用できません。値を確認して、もう一度やり直してください。|tableOptions の値のいずれかが無効です。|
|8012|無効な値|1 つ以上の format パラメーター値が使用できません。値を確認して、もう一度やり直してください。|foramt の値のいずれかが正しくありません。|
|8020|範囲外|行のインデックス値が許容範囲外です。行の数よりも少ない正の値 (0 以上) を使用してください。|行のインデックスが、テーブルの最大行のインデックスより大きいか、または 0 より小さいです。|
|8021|範囲外|列のインデックス値が許容範囲外です。列の数よりも少ない正の値 (0 以上) を使用してください。|列のインデックスが、テーブルの最大列のインデックスより大きいか、または 0 より小さいです。|
|8022|範囲外|値が許容範囲外です。|形式の値の一部がサポート範囲外です。|
|9016|アクセス許可が拒否されました|アクセスが拒否されました|アクセスが拒否されました。|

## バインディングの作成エラーの条件

API でバインディングが作成される場合、ソリューション開発者は使用するバインド タイプを示す必要があります。次の表は、様々な可能性と、バインドの動作の予測される結果をまとめたものです。


### Excel での動作

次の表は、Excel のバインドの動作をまとめたものです。



|**指定されたバインド タイプ**|**実際の選択**|**動作**|
|:-----|:-----|:-----|
|マトリックス|セルの範囲 (1 つのテーブル内にあり、単一セルの場合を含む)|選択されたセルでバインド タイプ  _matrix_ が作成されます。ドキュメント内の変更は想定されていません。|
|マトリックス|セル内の選択されたテキスト|セル全体でバインド タイプ  _matrix_ が作成されます。ドキュメント内の変更は想定されていません。|
|マトリックス|複数の選択項目または無効な選択項目 (たとえば、ユーザーが画像、オブジェクト、ワードアートなどを選択した場合)。|バインドを作成できません。|
|テーブル|セルの範囲 (単一セルの場合を含む)|バインドを作成できません。|
|テーブル|テーブル内のセルの範囲 (テーブル内の 1 つのセル、テーブル全体、またはテーブルのセル内のテキストの場合を含む)。|バインドがテーブル全体で作成されます。|
|テーブル|選択項目の半分がテーブルで、半分がテーブル外です|バインドを作成できません。|
|テーブル|(テーブル内ではなく) セル内で選択されたテキスト。|バインドを作成できません。|
|テーブル|複数の選択項目または無効な選択項目 (たとえば、ユーザーが画像、オブジェクト、ワードアートなどを選択した場合)。|バインドを作成できません。|
|テキスト|セルの範囲|バインドを作成できません。|
|テキスト|テーブル内のセルの範囲|バインドを作成できません。|
|テキスト|1 つのセル|バインド タイプ  _text_ が作成されます。|
|テキスト|テーブル内の 1 つのセル|バインド タイプ  _text_ が作成されます。|
|テキスト|セル内の選択されたテキスト|セル全体でバインド タイプ  _text_ が作成されます。|

### Word の動作

次の表は、Word のバインドの動作をまとめたものです。



|**指定されたバインド タイプ**|**実際の選択**|**動作**|
|:-----|:-----|:-----|
|マトリックス|テキスト|バインドを作成できません。|
|マトリックス|テーブル全体|バインド タイプ  _matrix_ が作成されます。ドキュメントが変更され、コンテンツ コントロールがテーブルをラップする必要があります。 |
|マトリックス|テーブル内の範囲|バインドを作成できません。|
|マトリックス|無効な選択項目 (たとえば、複数のオブジェクト、無効なオブジェクトなど)。|バインドを作成できません。|
|テーブル|テキスト|バインドを作成できません。|
|テーブル|テーブル全体|バインド タイプ  _text_ が作成されます。|
|テーブル|テーブル内の範囲|バインドを作成できません。|
|テーブル|無効な選択項目 (たとえば、複数のオブジェクト、無効なオブジェクトなど)。|バインドを作成できません。|
|テキスト|テーブル全体|バインド タイプ  _text_ が作成されます。|
|テキスト|テーブル内の範囲|バインドを作成できません。|
|テキスト|複数の選択項目|最後の選択項目がコンテンツ コントロール内でラップされ、そのコントロールにバインドされます。型  _text_ のコンテンツ コントロールが作成されます。|
|テキスト|無効な選択項目 (たとえば、複数のオブジェクト、無効なオブジェクトなど)。|バインドを作成できません。|

## その他のリソース


- [Office アドインの API とスキーマ参照](../reference/reference.md)
    
- [Office アドインの開発ライフ サイクル](../docs/design/add-in-development-lifecycle.md)
    