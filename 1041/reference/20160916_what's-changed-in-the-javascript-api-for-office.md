
# JavaScript API for Office の変更点
JavaScript API for Office は、Office アドインの機能を拡張するため、オブジェクト、メソッド、プロパティ、イベント、列挙体の新規追加や更新によって定期的に更新が加えられています。新規および更新された API のメンバーを確認するには、次のリンクを参照してください。

新しい API メンバーを使用してアドインを開発するには、[プロジェクトで JavaScript API for Office ファイルを更新する](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)必要があります。

前回の更新から変更されていない API メンバーを含むすべての API メンバーを表示するには、「[JavaScript API for Office](../reference/javascript-api-for-office.md)」を参照してください。


## 新規および更新された API

 **新規オブジェクトと更新されたオブジェクト**


|**Object**|**説明**|**追加または更新されたバージョン**|
|:-----|:-----|:-----|
|[Item](../reference/outlook/Office.context.mailbox.item.md)|次に対する更新および追加です。<br><ul><li><p>ユーザーの選択の取得と、メッセージまたは予定の件名と本文を上書きするための、<a href="../reference/outlook/Office.context.mailbox.item.md#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> および <a href="../reference/outlook/Office.context.mailbox.item.md#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> メソッド。</p></li><li><p>予定の返信フォームへの添付ファイルの追加をサポートする <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyAllForm" target="_blank">displayReplyAllForm</a> および <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyForm" target="_blank">displayReplyForm</a> メソッド。</p></li></ul>|Mailbox 1.2|
|[アイテム](../reference/outlook/Office.context.mailbox.item.md)|新規作成モードの Outlook アドインを作成するためのメソッドとフィールドを含めるよう更新されました。 |1.1|
|[Binding](../reference/shared/binding.md)|Access 用コンテンツ アドインにおけるテーブル バインドをサポートするよう更新されました。|1.1|
|[Bindings](../reference/shared/bindings.bindings.md)|Access 用コンテンツ アドインにおけるテーブル バインドをサポートするよう更新されました。|1.1|
|[本文](../reference/outlook/Body.md)|新規作成モードの Outlook アドインでメッセージや予定の本文を作成および編集できるよう追加されました。|1.1|
|[ドキュメント](../reference/shared/document.md)|次に対して更新および追加が行われました。 <ul><li><p>Access 用のコンテンツ アドインで <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>、<a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a>、および <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> の各プロパティをサポートします。</p></li><li><p>PowerPoint および Word 用アドインで、<a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> メソッドを使用してドキュメントを PDF として取得します。</p></li><li><p>Excel、PowerPoint、および Word 用アドインで、<a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> メソッドを使用してファイルのプロパティを取得します。</p></li><li><p>Excel および PowerPoint 用アドインで、<a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> メソッドを使用して、ドキュメント内の場所とオブジェクトに移動します。</p></li><li><p>PowerPoint 用アドインで、<a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> メソッドを使用して (新しい <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a> 列挙体を指定した場合)、選択したスライドの ID、タイトル、およびインデックスを取得します。</p></li></ul>|1.1|
|[Location](../reference/outlook/Location.md)|新規作成モードの Outlook アドインで予定の場所を設定できるよう追加されました。|1.1|
|[Office](../reference/shared/office.md)|Access 用コンテンツ アドインにおけるバインドの取得をサポートするよう select メソッドが更新されました。|1.1|
|[受信者](../reference/outlook/Recipients.md)|新規作成モードでメッセージや予定の受信者を取得および設定できるよう追加されました。|1.1|
|[設定値](../reference/shared/document.settings.md)|Access 用コンテンツ アドインにおけるカスタム設定の作成をサポートするよう更新されました。|1.1|
|[件名](../reference/outlook/Subject.md)|新規作成モードの Outlook アドインでメッセージや予定の件名を取得および設定できるよう追加されました。|1.1|
|[時刻](../reference/outlook/Time.md)|新規作成モードの Outlook アドインで予定の開始時刻および終了時刻を取得および設定できるよう追加されました。|1.1|



**新規列挙体および更新された列挙体**


|**Object**|**説明**|**バージョン**|
|:-----|:-----|:-----|
|[ActiveView](../reference/shared/activeview-enumeration.md)|ユーザーがドキュメントを編集できるかどうかなど、ドキュメントのアクティブなビューの状態を示します。PowerPoint 用アドインで、ユーザーがプレゼンテーション ( **スライド ショー**) を閲覧しているのか、スライドを編集しているのかを判断できるように追加されました。 |1.1|
|[CoercionType](../reference/shared/coerciontype-enumeration.md)|PowerPoint 用アドインで、 **Office.CoercionType.SlideRange** メソッドを使用して選択されたスライドの範囲の取得をサポートするよう **getSelectedDataAsync** が更新されました。|1.1|
|[EventType](../reference/shared/eventtype-enumeration.md)|新しい ActiveViewChanged イベントを含めるよう更新されました。|1.1|
|[FileType](../reference/shared/filetype-enumeration.md)|PDF 形式での出力を指定するよう更新されました。|1.1|
|[GoToType](../reference/shared/gototype-enumeration.md)|ドキュメントの移動先の場所またはオブジェクトを指定するよう追加されました。|1.1|

## その他のリソース


- [Office アドインの API とスキーマ参照](../reference/reference.md)
    
- [Office Add-ins?](../docs/overview/office-add-ins.md)
    
