
# 處理 Outlook 增益集中日期值的秘訣

適用於 Office 的 JavaScript API 針對大部分的儲存與日期和時間的擷取使用 JavaScript [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp) 物件。該 **Date** 物件提供方法，例如 [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp)、[getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp) 及 [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp)，其根據全球定位時間 (UTC) 時間傳回要求的日期或時間值。<br/><br/>
**Date** 物件也提供其他方法，例如 [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp)、[getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp) 及 [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp)，其根據「當地時間」傳回要求的日期或時間。<br/><br/>「當地時間」的概念主要取決於用戶端電腦上的瀏覽器和作業系統。例如，在大部分在 Windows 型用戶端電腦上執行的瀏覽器上，JavaScript 會呼叫 **getDate**，根據用戶端電腦上的 Windows 中設定的時區傳回日期。<br/><br/>
下列範例會以本地時間建立 **Date** 物件 <code>myLocalDate</code>，並呼叫 **toUTCString** 將該日期轉換成 UTC 日期字串。




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

雖然您可以使用 JavaScript  **Date** 物件依據 UTC 或用戶端電腦時區取得日期或時間值，**Date** 物件的一個方面有所限制 - 它不提供傳回任何其他特定時區的日期或時間值的方法。例如，如果您的用戶端電腦設定為美加東部標準時間 (EST)，則沒有 **Date** 方法可讓您取得 EST 或 UTC，例如太平洋標準時間 (PST)。


## Outlook 增益集的日期相關功能


當您使用適用於 Office 的 JavaScript API 在 Outlook 豐富型用戶端及 Outlook Web App 或裝置用 OWA 執行的 Outlook 增益集中處理日期或時間值時，先前所述的 JavaScript 限制具有隱含的意義。


### Outlook 用戶端的時區

為了清楚起見，我們定義有問題的時區。



|**時區**|**說明**|
|:-----|:-----|
|用戶端電腦時區|這個是在用戶端電腦的作業系統上設定。大部分的瀏覽器會使用用戶端電腦的時區來顯示 JavaScript 的日期或時間值 **Date** 物件。<br/><br/>Outlook 豐富型用戶端會使用這個時區在使用者介面中顯示日期或時間值。 <br/><br/>例如，在執行 Windows 的用戶端電腦上，Outlook 會使用 Windows 上設定的時區做為當地時區。在 Mac 上，如果使用者在用戶端電腦上變更時區，Outlook for Mac 也會提示使用者在 Outlook 中更新時區。|
|Exchange 系統管理中心 (EAC) 時區|當使用者第一次登入到 Outlook Web App 或裝置用 OWA 時，會設定這個時區值 (和偏好的語言)。 <br/><br/>Outlook Web App 和裝置用 OWA 會使用這個時區來顯示使用者介面中的日期或時間值。|
因為 Outlook 豐富型用戶端使用用戶端電腦的時區，且 Outlook Web App 和裝置用 OWA 的使用者介面使用 EAC 時區，當在 Outlook 豐富型用戶端及在 Outlook Web App 或裝置用 OWA 中執行時，相同信箱所安裝的相同增益集的當地時間可能會不同。身為 Outlook 增益集開發人員，您應該適當地輸入和輸出日期值，如此這些值會永遠與使用者所預期在對應用戶端上的時區一致。


### 日期相關的 API

以下是適用於 Office 的 JavaScript API 中的屬性及方法，其支援日期相關的 features.reference/outlook/Office.context.mailbox.item.md



**API 成員**|**時區表示**|**Outlook 豐富型用戶端中的範例**|**Outlook Web App 或裝置用 OWA 中的範例**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|在 Outlook 豐富型用戶端中，這個屬性會傳回用戶端電腦的時區。在 Outlook Web App 及裝置用 OWA 中，這個屬性會傳回 EAC 時區。 |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) 及 [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|這些屬性每一個都會傳回 JavaScript  **Date** 物件。此 **Date** 值是 UTC 更正，如下列範例中所顯示 - `myUTCDate` 在 Outlook rich client、Outlook Web App 及裝置用 OWA 中具有相同的值。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>不過，呼叫 `myDate.getDate` 會傳回用戶端電腦時區的日期值，其會與用來顯示 Outlook 豐富型用戶端介面中的日期時間值的時區一致，但可能與 Outlook Web App 及裝置用 OWA 在其使用者介面中使用的 EAC 時區不同。|如果項目在 9am UTC 建立：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` 傳回 4am EST.<br/><br/>如果項目在 11am UTC 修改：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` 傳回 6am EST。|如果項目建立時間為 9am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` 傳回 4am EST.<br/><br/>如果項目在 11am UTC 修改：<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` 傳回 6am EST。<br/><br/>請注意，如果您想要在使用者介面中顯示建立或修改時間，您要先將時間轉換成與其他使用者介面的一致的 PST。
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|每個 _Start_ 和 _End_ 參數需要 JavaScript **Date** 物件。這些引數應該是 UTC 更正，不論 Outlook 豐富型用戶端、Outlook Web App 或裝置用 OWA 的使用者介面中所使用的時區為何。|如果約會表單的開始和結束時間是 9am UTC 及 11am UTC，則您應該確保 `start` 和 `end` 引數是 UTC 更正，這表示︰<br/><br/><ul><li>`start.getUTCHours` 傳回 9am UTC</li><li>`end.getUTCHours` 傳回 11am UTC</li></ul>|如果約會表單的開始和結束時間是 9am UTC 及 11am UTC，則您應該確保 `start` 和 `end` 引數是 UTC 更正，這表示︰<br/><br/><ul><li>`start.getUTCHours` 傳回 9am UTC</li><li>`end.getUTCHours` 傳回 11am UTC</li></ul>

## 日期相關案例的協助程式方法


如先前章節中所述，因為使用者在 Outlook Web App 或裝置用 OWA 中的「當地時間」可能在 Outlook豐富型用戶端上會不同，但 JavaScript **Date** 物件僅支援轉換至用戶端電腦的時區或 UTC，適用於 Office 的 JavaScript API 提供兩個協助程式方法：[Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) 和 [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md)。 <br/><br/>
這些協助程式方法會在下列兩個日期相關案例中以不同方式負責處理任何處理日期或時間的需要，在 Outlook 豐富型用戶端、Outlook Web App 及裝置用 OWA，因此會針對增益集的不同用戶端強化「寫入一次」。


### 案例 A：顯示項目建立或修改時間

如果您要在使用者介面中顯示項目建立時間 (**Item.dateTimeCreated**) 或修改時間 (**Item.dateTimeModified**)，請先使用 **convertToLocalClientTime** 轉換這些屬性所提供的 **Date** 物件，以取得適當的當地時間的字典表示。然後顯示字典日期的部分。下列是這個案例的範例：


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

請注意，**convertToLocalClientTime** 負責 Outlook 豐富型用戶端與 Outlook Web App 或裝置用 OWA 之間的差異：


- 如果 **convertToLocalClientTime** 偵測到目前的主應用程式為豐富型用戶端，則方法會將 **Date** 表示轉換成同一個用戶端電腦時區的字典表示，與其餘的豐富型用戶端使用者介面一致。
    
- 如果 **convertToLocalClientTime** 偵測到目前的主應用程式為 Outlook Web App 或裝置用 OWA，則方法會將 UTC 修正 **Date** 表示轉換成 EAC 時區的字典表示，與其餘的 Outlook Web App 或裝置用 OWA 使用者介面一致。
    

### 案例 B：在新的約會表單中顯示開始和結束日期

如果要取得做為以當地時間中表示的日期時間值輸入的不同部分，並且要提供此字典輸入值做為約會表單中的開始或結束時間，請先使用 **convertToUtcClientTime** 協助程式方法將字典值轉換成 UTC 更正 **Date** 物件。<br/><br/>在下列範例中，假設 `myLocalDictionaryStartDate` 和 `myLocalDictionaryEndDate` 是您從使用者取得的字典格式的日期時間值。這些值是根據當地時間，取決於主應用程式。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

結果值 (`myUTCCorrectStartDate` 和 `myUTCCorrectEndDate`) 為 UTC 更正。然後傳送 **Date** 物件做為 _Mailbox.displayNewAppointmentForm_ 方法的 _Start_ 和 **End** 參數的引數，以顯示新的約會表單。<br/><br/>
請注意，**convertToUtcClientTime** 負責 Outlook 豐富型用戶端與 Outlook Web App 或裝置用 OWA 之間的差異：


- 如果 **convertToUtcClientTime** 偵測到目前的主應用程式為 Outlook 豐富型用戶端，方法只會將字典表示轉換成 **Date** 物件。此 **Date** 物件為 UTC 更正，如 **displayNewAppointmentForm** 所預期。
    
- 如果 **convertToUtcClientTime** 偵測到目前的主應用程式為 Outlook Web App 或裝置用 OWA，方法會將以 EAC 時區表示的日期和時間值的字典格式轉換成 **Date** 物件。此 **Date** 物件為 UTC 更正，如 **displayNewAppointmentForm** 所預期。
    

## 其他資源



- [部署和安裝 Outlook 增益集以進行測試](../outlook/testing-and-tips.md)
    


