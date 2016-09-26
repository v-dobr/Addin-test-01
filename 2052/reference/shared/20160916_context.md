
# Context 对象
表示 外接程序 的运行时环境，并提供对 API 的关键对象的访问。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```
Office.context
```


## 成员

|||
|:-----|:-----|
|姓名|说明|
|[commerceAllowed](../../reference/shared/office.context.commerceallowed.md)|获取外接程序是否将运行在允许链接到外部付款系统的平台上。|
|[contentLanguage](../../reference/shared/office.context.contentlanguage.md)|存储到文档或项中时获取数据的区域设置（语言）。|
|[displayLanguage](../../reference/shared/office.context.displaylanguage.md)|获取宿主应用程序的 UI 的区域设置（语言）。|
|[文档](../../reference/shared/office.context.document.md)|获取表示正与内容或任务窗格外接程序交互的文档的对象。|
|[mailbox](../../reference/shared/office.context.mailbox.md)|获取提供专门针对 Outlook 外接程序的 API 的成员的访问的  **mailbox** 对象。|
|[officeTheme](../../reference/shared/office.context.officetheme.md)|提供了对 Office 主题颜色属性的访问权限。|
|[ui](../../reference/shared/officeui)|提供可用于创建和操作 UI 组件（如对话框）的对象和方法。|
|[roamingSettings](../../reference/shared/office.context.roamingsettings.md)|获取表示外接程序的已保存自定义设置的对象。|
|[touchEnabled](../../reference/shared/office.context.touchenabled.md)|获取外接程序是否将运行在已启用触控的 Office 主机应用程序中。|

## 备注

提供对 JavaScript API for Office 中的关键对象的访问的  **Context** 对象。


## 支持详细信息



|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格、Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|添加  **commerceAllowed** 和 **touchEnabledAdded** 属性（Excel、PowerPoint 和 Word 仅在 Office for iPad 上）。|
|1.1|增加了对支持Office for iPad 上外接程序与 Excel 和 Word 的支持。|
|1.1|对于 [contentLanguage](../../reference/shared/office.context.contentlanguage.md)， [displayLanguage](../../reference/shared/office.context.displaylanguage.md) 和[document](../../reference/shared/office.context.document.md)，增加了对 Access 相关内容外接程序的支持。|
|1.0|引入|
