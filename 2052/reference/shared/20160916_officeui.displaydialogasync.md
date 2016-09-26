# UI.displayDialogAsync 方法

在 Office 主机中显示一个对话框。 

## 要求

|主机|引入版本|包含最后一次更改的版本|
|:---------------|:--------|:----------|
|Word、Excel、PowerPoint|1.1|1.1|
|Outlook|Mailbox 1.4|Mailbox 1.4|

此方法适用于 DialogAPI [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)。 若要指定 DialogAPI 要求集，请在清单中使用以下内容。

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="DialogAPI"/> 
   </Sets> 
 </Requirements> 

```

若要在运行时检测此 API，请使用以下代码。

```js
 if (Office.context.requirements.isSetSupported('DialogAPI', 1.1)) 
    {  
         // Use Office UI methods; 
    } 
 else 
     { 
         // Alternate path 
     } 
```



### 支持的平台
目前，以下平台支持 DialogAPI 要求集：

  - Office for Windows Desktop 2016（版本 16.0.6741.0000 或更高版本）
  - Office for IPad（版本 1.22 或更高版本）
  - Office for Mac（版本 15.20 或更高版本） 

即将推出更多平台。 

## 语法

```js
office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##示例

有关使用 **displayDialogAsync** 方法的简单示例，请参阅 GitHub 上的 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/)。

有关显示身份验证方案的示例，请参阅 GitHub 上的 [AngularJS 的 Office 外接程序 Office 365 客户端身份验证](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth) 示例。

 
## 参数

| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|startAddress|字符串|接受在对话框中打开的初始 HTTPS(TLS) URL。 <ul><li>初始网页必须与父页位于相同的域。 初始网页加载后，你可以转到其他域。</li><li>调用 [office.context.ui.messageParent](officeui.messageparent.md) 的所有页也必须都与父页位于相同的域。</li></ul>|
|选项|object|可选。接受用于定义对话框行为的 options 对象。|
|callback|对象|接受用于处理对话框创建尝试的 callback 方法。|
    
### 配置选项
以下配置选项适用于对话框。


| 属性     | 类型   |说明|
|:---------------|:--------|:----------|
|**width**|对象|可选。 以占当前显示器的百分比的形式，定义对话框的宽度。 默认值为 80%。 最小分辨率为 250 像素。|
|**高度**|对象|可选。 以占当前显示器的百分比的形式，定义对话框的高度。 默认值为 80%。 最小分辨率为 150 像素。|
|**displayInIFrame**|对象|可选。 确定对话框是否应在 Office Online 客户端中的 IFrame 内显示。 桌面客户端会忽略此设置。 以下是可能的值：<ul><li>False（默认值）- 对话框将显示为一个新的浏览器窗口（弹出窗口）。 对于无法在 IFrame 中显示的身份验证页建议使用此值。 </li><li>True - 对话框将显示为使用 IFrame 的浮动重叠窗口。 对于用户体验和性能而言，这是最佳选择。</li>|


## 回调值
在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **displayDialogAsync** 方法的回调函数中，你可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问 [Dialog](../../reference/shared/officeui.dialog.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果你将用户定义的对象或值作为 _asyncContext_ 参数传递，则对其进行访问。|


## 设计注意事项
下列设计注意事项适用于对话框：

- Office 外接程序随时都可能有一个打开的对话框。
- 用户可以移动每个对话框和调整其大小。
- 每个对话框在打开时都在屏幕上居中显示。
- 对话框按照创建的顺序出现在主机应用程序顶部。

使用对话框可以执行以下操作：

- 显示身份验证页以收集用户凭据。
- 显示来自 ShowTaspane 或 ExecuteAction 命令的错误/进度/输入屏幕。
- 临时增加用户可用于完成一项任务的表面区域。

不要使用对话框与文档进行交互。 而是使用任务窗格。 

有关可以用于创建对话框的设计模式，请参阅 GitHub 的 Office 外接程序 UX 设计模式存储库中的 [客户端对话框](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md)。
