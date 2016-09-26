
# JavaScript API for Office 中的更改内容
JavaScript API for Office 将定期更新新增和更新的对象、方法、属性、事件和枚举，以扩展 Office 外接程序的功能。使用下面的链接可查看新增和更新的 API 成员。

若要使用新的 API 成员开发外接项目，你需要 [在项目中更新适用于 Office 的 JavaScript API 文件](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。

若要查看所有 API 成员（包括与之前更新相比未变化的成员），请参阅 [适用于 Office 的 JavaScript API](../reference/javascript-api-for-office.md)。


## 新 API 和更新的 API

 **新增和更新的对象**


|**Object**|**说明**|**添加或更新了版本**|
|:-----|:-----|:-----|
|[项目](../reference/outlook/Office.context.mailbox.item.md)|更新和新增功能：<br><ul><li><p><a href="../reference/outlook/Office.context.mailbox.item.md#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> 和 <a href="../reference/outlook/Office.context.mailbox.item.md#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> 方法支持获取用户所选的内容并将其覆盖到邮件或约会的主题和正文中。</p></li><li><p><a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyAllForm" target="_blank">displayReplyAllForm</a> 和 <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyForm" target="_blank">displayReplyForm</a> 方法支持向约会的答复表单添加附件。</p></li></ul>|邮箱 1.2|
|[项目](../reference/outlook/Office.context.mailbox.item.md)|进行了更新以包括用于创建撰写模式 Outlook 外接程序的方法和字段。 |1.1|
|[Binding](../reference/shared/binding.md)|进行了更新以支持 Access 内容加载项中的表绑定。|1.1|
|[Bindings](../reference/shared/bindings.bindings.md)|进行了更新以支持 Access 内容加载项中的表绑定。|1.1|
|[Body](../reference/outlook/Body.md)|进行了添加以便能够在撰写模式 Outlook 外接程序中创建和编辑邮件或约会的正文。|1.1|
|[文档](../reference/shared/document.md)|进行了更新和添加以： <ul><li><p>支持 Access 内容加载项中的 <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>、<a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a> 和 <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> 属性。</p></li><li><p>在 PowerPoint 和 Word 加载项中通过 <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> 方法获取 PDF 格式文档。</p></li><li><p>在 Excel、PowerPoint 和 Word 加载项中通过 <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> 方法获取文件属性。</p></li><li><p>在 Excel 和 Powerpoint 加载项中通过 <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> 方法导航到文档中的位置和对象。</p></li><li><p>在 PowerPoint 加载项中通过 <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> 方法（当您指定新的 <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a> 枚举时）获取选定幻灯片的 ID、标题和索引。</p></li></ul>|1.1|
|[位置](../reference/outlook/Location.md)|进行了添加以便能够在撰写模式 Outlook 外接程序中设置约会的地点。|1.1|
|[Office](../reference/shared/office.md)|更新了选择方法以支持获取 Access 内容加载项中的绑定。|1.1|
|[收件人](../reference/outlook/Recipients.md)|进行了添加以便能够在撰写模式下获取和设置邮件或约会的收件人。|1.1|
|[Settings](../reference/shared/document.settings.md)|进行了更新以支持在 Access 内容加载项中创建自定义设置。|1.1|
|[主题](../reference/outlook/Subject.md)|进行了添加以便能够在撰写模式 Outlook 外接程序中获取和设置邮件或约会的主题。|1.1|
|[时间](../reference/outlook/Time.md)|进行了添加以便能够在撰写模式 Outlook 外接程序中获取和设置约会的开始和结束时间。|1.1|



**新增和更新的枚举**


|**Object**|**说明**|**版本**|
|:-----|:-----|:-----|
|[ActiveView](../reference/shared/activeview-enumeration.md)|指定文档活动视图的状态，例如，用户是否可以编辑 document.Added，以便 PowerPoint 的外接程序可以确定用户是否正在查看演示文稿（**幻灯片放映**）或编辑幻灯片。 |1.1|
|[CoercionType](../reference/shared/coerciontype-enumeration.md)|使用  **Office.CoercionType.SlideRange** 进行更新，以支持在 PowerPoint 加载项中通过 **getSelectedDataAsync** 方法获取选定幻灯片范围。|1.1|
|[EventType](../reference/shared/eventtype-enumeration.md)|进行了更新以包含新的 ActiveViewChanged 事件。|1.1|
|[FileType](../reference/shared/filetype-enumeration.md)|进行了更新以指定 PDF 格式的输出。|1.1|
|[GoToType](../reference/shared/gototype-enumeration.md)|进行了添加以指定要转到的文档位置或对象。|1.1|

## 其他资源


- [Office 外接程序 API 和架构参考](../reference/reference.md)
    
- [Office 外接程序](../docs/overview/office-add-ins.md)
    
