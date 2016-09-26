
# 內容物件
代表增益集的執行階段環境，並且提供 API 的主要物件之存取。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```
Office.context
```


## 成員

|||
|:-----|:-----|
|名稱|說明|
|[commerceAllowed](../../reference/shared/office.context.commerceallowed.md)|取得關於增益集是否在允許連結到外部付款系統的平台上執行的資訊。|
|[contentLanguage](../../reference/shared/office.context.contentlanguage.md)|取得當資料儲存在文件或項目中時的地區設定 (語言)。|
|[displayLanguage](../../reference/shared/office.context.displaylanguage.md)|取得主控應用程式 UI 的地區設定 (語言)。|
|[文件](../../reference/shared/office.context.document.md)|取得代表與內容或工作窗格增益集互動之文件的物件。|
|[信箱](../../reference/shared/office.context.mailbox.md)|取得提供特別針對 Outlook 增益集 API 成員存取的**信箱**物件。|
|[officeTheme](../../reference/shared/office.context.officetheme.md)|提供 Office 佈景主題色彩屬性的存取。|
|[UI](../../reference/shared/officeui)|提供物件和方法，您可以用來建立和操作 UI 元件，例如對話方塊。|
|[roamingSettings](../../reference/shared/office.context.roamingsettings.md)|取得代表增益集已儲存自訂設定的物件。|
|[touchEnabled](../../reference/shared/office.context.touchenabled.md)|取得關於增益集是否在具有觸控功能的 Office 主應用程式中執行的資訊。|

## 備註

**Context** 物件提供適用於 Office 的 JavaScript API 中索引鍵物件的存取。


## 支援詳細資料



|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 **commerceAllowed** 與 **touchEnabledAdded** 屬性 (僅限 iPad 版 Office 中的 Excel、PowerPoint 與 Word)。|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的增益集支援。|
|1.1|對於 [contentLanguage](../../reference/shared/office.context.contentlanguage.md)、[displayLanguage](../../reference/shared/office.context.displaylanguage.md) 與 [文件](../../reference/shared/office.context.document.md)，新增對 Access 內容增益集的支援。|
|1.0|已導入|
