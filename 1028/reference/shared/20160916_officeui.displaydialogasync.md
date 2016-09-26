# UI.displayDialogAsync 方法

在 Office 主應用程式中顯示對話方塊。 

## 需求

|主應用程式|導入在|上次變更於|
|:---------------|:--------|:----------|
|Word、Excel、PowerPoint|1.1|1.1|
|Outlook|信箱 1.4|信箱 1.4|

此方法只有在 DialogAPI [需求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)才可用。 若要指定 DialogAPI 需求集，請使用資訊清單中的下列項目。

```xml
 <Requirements> 
   <Sets DefaultMinVersion="1.1"> 
     <Set Name="DialogAPI"/> 
   </Sets> 
 </Requirements> 

```

若要在執行階段偵測這個 API，請使用下列程式碼。

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



### 支援的平台
目前在下列平台上支援 DialogAPI 需求集︰

  - Office for Windows Desktop 2016 (組建 16.0.6741.0000 或更新版本)
  - Office for IPad (組建 1.22 或更新版本)
  - Office for Mac (組建 15.20 或更新版本) 

將支援更多平台。 

## 語法

```js
office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##範例

如需使用 **displayDialogAsync** 方法的簡單範例，請參閱 GitHub 上的 [Office 增益功能對話方塊 API 範例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/)。

如需顯示驗證案例的範例，請參閱 GitHub 上的 [適用於 AngularJS 的 Office 增益集 Office 365 用戶端驗證](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)範例。

 
## 參數

| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|startAddress|字串|接受在對話方塊中開啟的初始 HTTPS(TLS) URL。 <ul><li>初始頁面必須位於與父系頁面相同的網域。 初始頁面載入之後，您可以移至其他網域。</li><li>任何呼叫 [office.context.ui.messageParent](officeui.messageparent.md) 的頁面必須位於與父系頁面相同的網域。</li></ul>|
|選項|物件|選擇性。接受 options 物件，以定義對話方塊的行為。|
|callback|object|接受 callback 方法，以處理建立對話方塊的嘗試。|
    
### 組態選項
下列組態選項可用於對話方塊。


| 屬性     | 類型	   |描述|
|:---------------|:--------|:----------|
|**width**|object|選用。 將對話方塊的寬度定義為目前顯示的百分比。 預設值為 80%。 最小解析為 250 像素。|
|**height**|object|選用。 將對話方塊的高度定義為目前顯示的百分比。 預設值為 80%。 最小解析為 150 像素。|
|**displayInIFrame**|物件|選用。 決定對話方塊是否應該在 Office Online 用戶端中的 IFrame 顯示。 桌面用戶端會忽略這項設定。 可能的值如下：<ul><li>False (預設值) - 對話方塊會顯示為新的瀏覽器視窗 (快顯視窗)。 針對無法在 IFrame 中顯示的驗證頁面建議使用此選項。 </li><li>True - 對話方塊會顯示為與 IFrame 浮動重疊。 這對於使用者經驗與效能是最佳選項。</li>|


## 回呼值
傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **displayDialogAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用途**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取 [Dialog](../../reference/shared/officeui.dialog.md) 物件。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義物件或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|


## 設計考量
下列設計考量適用於對話方塊：

- Office 增益集任何時候只能有一個對話方塊開啟。
- 每個對話方塊都可以由使用者移動或調整。
- 每個對話方塊在開啟時都會在螢幕上置中。
- 對話方塊會顯示在主應用程式的最上方，並且以其建立的順序顯示。

使用對話方塊來：

- 顯示驗證頁面以收集使用者認證。
- 顯示從 ShowTaspane 或 ExecuteAction 命令取得的錯誤/進度/輸入畫面。
- 暫時增加使用者可以用來完成工作的表面區域。

請勿使用對話方塊與文件互動。 改為使用工作窗格。 

如需您可以用來建立對話方塊的設計模式，請參閱 GitHub 上 Office 增益集 UX 設計模式存放庫中的[用戶端對話方塊](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md)。
