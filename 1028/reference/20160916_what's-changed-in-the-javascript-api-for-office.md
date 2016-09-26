
# 適用於 Office 的 JavaScript API 中的變更項目
適用於 Office 的 JavaScript API 會定期以新的或更新的物件、方法、屬性、事件和列舉型別更新，來擴充您 Office 增益集的功能。請使用下列連結來查看全新及更新的 API 成員。

若要使用新的 API 成員開發增益集，您必須[更新專案中的 JavaScript API for Office 檔案](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。

若要檢視所有 API 成員，包括先前更新後未變更的成員，請參閱 [JavaScript API for Office](../reference/javascript-api-for-office.md)。


## 新增及更新的 API

 **新增及更新的物件**


|**物件**|**說明**|**新增或更新的版本**|
|:-----|:-----|:-----|
|[項目](../reference/outlook/Office.context.mailbox.item.md)|更新和加入的項目︰<br><ul><li><p><a href="../reference/outlook/Office.context.mailbox.item.md#getSelectedDataAsync" target="_blank">GetSelectedDataAsync</a> 和 <a href="../reference/outlook/Office.context.mailbox.item.md#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> 方法可支援取得使用者的選取範圍，並覆寫在主旨和郵件或約會的本文。</p></li><li><p><a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyAllForm" target="_blank">DisplayReplyAllForm</a> 和 <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyForm" target="_blank">displayReplyForm</a> 方法可支援將附件新增至約會的回覆表單。</p></li></ul>|Mailbox 1.2|
|[項目](../reference/outlook/Office.context.mailbox.item.md)|更新以包含方法和欄位，以建立撰寫模式 Outlook 增益集。 |1.1|
|[Binding](../reference/shared/binding.md)|更新支援 Access 內容增益集中的表格繫結。|1.1|
|[Bindings](../reference/shared/bindings.bindings.md)|更新支援 Access 內容增益集中的表格繫結。|1.1|
|[本文](../reference/outlook/Body.md)|新增以在撰寫模式 Outlook 增益集中，建立和編輯郵件或約會的本文。|1.1|
|[Document](../reference/shared/document.md)|更新和加入的項目︰ <ul><li><p>支援 Access 內容增益集內的 <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>、<a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a> 及 <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> 屬性。</p></li><li><p>以 PowerPoint 和 Word 增益集內的 <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> 方法取得 PDF 形式的文件。</p></li><li><p>以 Excel、PowerPoint 及 Word 增益集內的 <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> 方法取得檔案屬性。</p></li><li><p>以 Excel 和 PowerPoint 增益集內的 <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> 方法瀏覽至文件中的位置和物件。</p></li><li><p>以 PowerPoint 增益集內的 <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> 方法取得選定投影片的識別碼、標題及索引 (當您指定新的 <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a> 列舉時)。</p></li></ul>|1.1|
|[位置](../reference/outlook/Location.md)|新增在撰寫模式 Outlook 增益集中啟用約會的位置設定。|1.1|
|[Office](../reference/shared/office.md)|更新選取方法，支援 Access 的內容增益集中的取得繫結。|1.1|
|[收件者](../reference/outlook/Recipients.md)|新增在撰寫模式中，啟用郵件或約會的收件者之取得和設定。|1.1|
|[設定](../reference/shared/document.settings.md)|更新以支援在 Access 的內容增益集中建立自訂設定。|1.1|
|[主旨](../reference/outlook/Subject.md)|新增在撰寫模式 Outlook 增益集中，啟用郵件或約會之主旨的取得和設定。|1.1|
|[時間](../reference/outlook/Time.md)|新增在撰寫模式 Outlook 增益集中，啟用約會之開始和結束時間的取得和設定。|1.1|



**新增及更新的列舉**


|**物件**|**說明**|**版本**|
|:-----|:-----|:-----|
|[ActiveView](../reference/shared/activeview-enumeration.md)|指定文件的使用中檢視狀態，例如使用者是否可以編輯文件。已加入，讓 PowerPoint 的增益集可以判斷使用者是否正在檢視簡報 (**投影片放映**) 或編輯投影片。 |1.1|
|[CoercionType](../reference/shared/coerciontype-enumeration.md)|更新 **Office.CoercionType.SlideRange**，支援在 PowerPoint 的增益集中，使用 **getSelectedDataAsync** 方法取得選定的投影片範圍。|1.1|
|[EventType](../reference/shared/eventtype-enumeration.md)|更新以包含新的 ActiveViewChanged 事件。|1.1|
|[FileType](../reference/shared/filetype-enumeration.md)|更新以指定 PDF 格式的輸出。|1.1|
|[GoToType](../reference/shared/gototype-enumeration.md)|加入以指定文件中要移至的位置或物件。|1.1|

## 其他資源


- [Office 增益集 API 和結構描述參考](../reference/reference.md)
    
- [Office 增益集](../docs/overview/office-add-ins.md)
    
