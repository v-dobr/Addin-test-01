
# Context オブジェクト
アドインのランタイム環境を表し、API の主要なオブジェクトへのアクセスを提供します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```
Office.context
```


## メンバー

|||
|:-----|:-----|
|名前|説明|
|[commerceAllowed](../../reference/shared/office.context.commerceallowed.md)|外部の支払システムにリンクできるプラットフォーム上で アドイン が実行されているかどうかを取得します。|
|[contentLanguage](../../reference/shared/office.context.contentlanguage.md)|ドキュメントまたはアイテムに保存されているデータのロケール (言語) を取得します。|
|[displayLanguage](../../reference/shared/office.context.displaylanguage.md)|ホスト アプリケーションの UI のロケール (言語) を取得します。|
|[document](../../reference/shared/office.context.document.md)|コンテンツ アドインまたは作業ウィンドウ アドインによって操作するドキュメントを表すオブジェクトを取得します。|
|[mailbox](../../reference/shared/office.context.mailbox.md)|特に Outlook アドイン向けに API のメンバーへのアクセスを提供する  **mailbox** オブジェクトを取得します。|
|[officeTheme](../../reference/shared/office.context.officetheme.md)|Office テーマの色のプロパティにアクセスできるようにします。|
|[UI](../../reference/shared/officeui)|ダイアログ ボックスなどの UI コンポーネントの作成や操作に使用できるオブジェクトとメソッドを提供します。|
|[roamingSettings](../../reference/shared/office.context.roamingsettings.md)|アドインの保存されているカスタム設定を表すオブジェクトを取得します。|
|[touchEnabled](../../reference/shared/office.context.touchenabled.md)|タッチ対応 Office ホスト アプリケーションで アドイン が実行されているかどうかを取得します。|

## 注釈

**Context** オブジェクトは、JavaScript API for Office の主要なオブジェクトへのアクセスを提供します。


## サポートの詳細



|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ、Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## サポート履歴



****


|**変更内容**|**1.1**|
|:-----|:-----|
|1.1|**commerceAllowed** プロパティと **touchEnabledAdded** プロパティが追加されました (Office for iPad 上の Excel、PowerPoint、および Word のみ)。|
|1.1|Office for iPad 上の Excel と Word での アドイン のサポートが追加されました。|
|1.1|[contentLanguage](../../reference/shared/office.context.contentlanguage.md)、[displayLanguage](../../reference/shared/office.context.displaylanguage.md)、および[ドキュメント](../../reference/shared/office.context.document.md)で、Access 用コンテンツ アドインのサポートが追加されました。|
|1.0|導入|
