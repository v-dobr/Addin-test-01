
# 处理 Outlook 外接程序中的日期值的提示

适用于 Office 的 JavaScript API 将 JavaScript [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp) 对象用于大多数日期和时间存储和检索。该 **Date** 对象提供一些方法，如 [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp)、[getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp) 和 [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp)，它根据协调世界时 (UTC) 时间返回请求的日期或时间值。<br/><br/>
**Date** 对象还提供了其他方法，如 [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp)、[getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp) 和 [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp)，它根据本地时间返回请求的日期或时间。<br/><br/>“本地时间”概念很大程度上取决于客户端计算机上的浏览器和操作系统。例如，在大多数运行于基于 Windows 的客户端计算机的浏览器上，JavaScript 调用 **getDate** 根据客户端计算机上 Windows 中设置的时区返回日期。<br/><br/>
下面的示例以本地时间创建 **Date** 对象 <code>myLocalDate</code>，然后调用 **toUTCString** 将该日期转换为 UTC 格式的日期字符串。




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

虽然可以使用 JavaScript **Date** 对象获取基于 UTC 或客户端计算机时区的日期或时间，但是 **Date** 对象在一个方面存在限制 – 它不提供方法以针对任何其他特定时区返回日期或时间值。例如，如果客户端计算机设置为东部标准时间 (EST)，则没有 **Date** 方法可用于获得除 EST 或 UTC 之外的时间值，例如太平洋标准时间 (PST)。


## Outlook 外接程序的日期相关功能


当您使用适用于 Office 的 JavaScript API 处理在 Outlook 富客户端和 Outlook Web App 或适用于设备的 OWA 上运行的 Outlook 外接程序中的日期或时间值时，应考虑前面提及的 JavaScript 限制。


### Outlook 客户端的时区

为清楚起见，让我们先定义要讨论的时区。



|**时区**|**说明**|
|:-----|:-----|
|客户端计算机时区|这在客户端计算机的操作系统上设置。大多数浏览器使用客户端计算机时区来显示 JavaScript **Date** 对象的日期或时间值。<br/><br/>Outlook 富客户端使用此时区在用户界面中显示日期或时间值。 <br/><br/>例如，在运行 Windows 的客户端计算机上，Outlook 将使用 Windows 上设置的时区作为本地时区。在 Mac 上，如果用户更改客户端计算机上的时区，Outlook for Mac 会提示用户同时更新 Outlook 中的时区。|
|Exchange 管理中心 (EAC) 时区|当用户首次登录到 Outlook Web App 或适用于设备的 OWA 时，用户需设置此时区（和首选语言）。 <br/><br/>Outlook Web App 和适用于设备的 OWA 使用此时区在用户界面中显示日期或时间值。|
由于 Outlook 富客户端使用客户端计算机时区，而 Outlook Web App 和适用于设备的 OWA 用户界面使用 EAC 时区，因此，当在 Outlook 富客户端和 Outlook Web App 或适用于设备的 OWA 中运行时，针对同一邮箱安装的同一外接程序的本地时间可能会不同。作为 Outlook 外接程序开发人员，您应该正确输入和输出日期值，以便那些值始终与用户期望的相应客户端上的时区保持一致。


### 日期相关的 API

以下是支持日期相关功能的适用于 Office 的 JavaScript API 中的属性和方法。reference/outlook/Office.context.mailbox.item.md



**API 成员**|**时区表示形式**|**Outlook 富客户端的示例**|**Outlook Web App 或适用于设备的 OWA 中的示例**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|在 Outlook 富客户端中，此属性返回客户端计算机时区。在 Outlook Web App 和适用于设备的 OWA 中，此属性返回 EAC 时区。 |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) 和 [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|上述每个属性返回 JavaScript **Date** 对象。此 **Date** 值采用 UTC 格式，如以下示例所示 - 在 Outlook 富客户端、Outlook Web App 和适用于设备的 OWA 中，`myUTCDate` 具有相同的值。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>但是，在客户端计算机的时区，调用 `myDate.getDate` 返回一个 date 值，它与用来显示 Outlook 富客户端中日期时间值的时区一致，但可能不同于其用户界面中 Outlook Web App 和适用于设备的 OWA 使用的 EAC 时区。|如果此项的创建时间是 9am UTC：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果此项的修改时间是 11am UTC：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` 返回 6am EST。|如果此项的创建时间是 9am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果此项的修改时间是 11am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` 返回 6am EST。<br/><br/>请注意，如果您想要在用户界面中显示创建或修改时间，要首先将时间转换为 PST 以与用户界面的其余部分保持一致。
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|每个 _Start_ 和 _End_ 参数都需要一个 JavaScript **Date** 对象。该参数应采用 UTC 格式，不考虑 Outlook 富客户端、Outlook Web App 或适用于设备的 OWA 的用户界面中使用的时区。|如果约会窗体的开始和结束时间分别是 9am UTC 和 11am UTC，则应确保 `start` 和 `end` 参数都是 UTC 格式，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>|如果约会窗体的开始和结束时间分别是 9am UTC 和 11am UTC，则应确保 `start` 和 `end` 参数都是 UTC 格式，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>

## 日期相关应用场景的帮助程序方法


如前面几节所述，因为 Outlook Web App 或适用于设备的 OWA 中用户的“本地时间”在 Outlook 富客户端上可以不同，但 JavaScript **Date** 对象支持仅转换为客户端计算机时区或 UTC，适用于 Office 的 JavaScript API 提供了两个帮助程序方法：[Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) 和 [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md)。 <br/><br/>
在 Outlook 富客户端、Outlook Web App 和适用于设备的 OWA 中，对于下面两个日期相关场景，这些帮助程序方法以不同的方式处理需要处理的日期或时间，从而对于外接程序的不同客户端强制“写入一次”。


### 应用场景 A：显示项目创建时间或修改时间

如果要在用户界面中显示项目创建时间 (**Item.dateTimeCreated**) 或修改时间 (**Item.dateTimeModified**)，请首先使用 **convertToLocalClientTime** 转换这些属性提供的 **Date** 对象以获取采用正确本地时间的字典表示形式。然后显示字典日期的各个部分。下面是此应用场景的一个示例：


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook web app or OWA for Devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

请注意，**convertToLocalClientTime** 会考虑 Outlook 富客户端、Outlook Web App 和适用于设备的 OWA 之间的差异：


- 如果 **convertToLocalClientTime** 检测到当前的主机为富客户端，那么该方法使用同一客户端计算机时区（与富客户端用户界面的其余部分保持一致）将 **Date** 表示形式转换为字典表示形式。
    
- 如果 **convertToLocalClientTime** 检测到当前主机为 Outlook Web App 或适用于设备的 OWA，那么该方法将采用 UTC 格式的 **Date** 表示形式转换为采用 EAC 时区（与 Outlook Web App 或适用于设备的 OWA 用户界面的其余部分保持一致）的字典格式。
    

### 应用场景 B：在新的约会表单中显示开始日期和结束日期

如果您要获取作为输入不同组成部分的以本地时间形式表示的日期时间值，并且希望在约会窗体中将该字典输入值作为开始或结束时间提供，请首先使用 **convertToUtcClientTime** 帮助程序方法将字典值转换为采用 UTC 格式的 **Date** 对象。<br/><br/>在以下示例中，假定 `myLocalDictionaryStartDate` 和 `myLocalDictionaryEndDate` 是从用户获得的采用字典格式的日期时间值。这些值基于本地时间，具体取决于主机应用程序。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

结果值 `myUTCCorrectStartDate` 和 `myUTCCorrectEndDate` 采用 UTC 格式。然后将这些 **Date** 对象作为 _Mailbox.displayNewAppointmentForm_ 方法的 _Start_ 和 **End** 参数的自变量来显示新约会窗体。<br/><br/>
请注意，**convertToUtcClientTime** 会考虑 Outlook 富客户端、Outlook Web App 和适用于设备的 OWA 之间的差异：


- 如果 **convertToUtcClientTime** 检测到当前主机为 Outlook 富客户端，那么该方法只是将字典表示形式转换为 **Date** 对象。此 **Date** 对象采用 UTC 格式，正如 **displayNewAppointmentForm** 期望的那样。
    
- 如果 **convertToUtcClientTime** 检测到当前主机为 Outlook Web App 或适用于设备的 OWA，那么该方法将采用 EAC 时区表示的日期和时间值的字典格式转换为 **Date** 对象。此 **Date** 对象采用 UTC 格式，正如 **displayNewAppointmentForm** 预期的那样。
    

## 其他资源



- [部署和安装 Outlook 外接程序以进行测试](../outlook/testing-and-tips.md)
    


