# OneNote JavaScript API 程式設計的概觀

OneNote 為 OneNote Online 增益集推出 JavaScript API。 您可以建立工作窗格增益集、內容增益集和增益集命令，與 OneNote 物件互動，並連接到 Web 服務或其他以網路為基礎的資源。

增益集是由兩個基本元件所組成︰

- **Web 應用程式**，由網頁以及任何所需的 JavaScript、CSS 或其他檔案所組成。這些檔案是裝載在 Web 伺服器或虛擬主機服務上，例如 Microsoft Azure。在 OneNote Online 中，Web 應用程式會顯示在瀏覽器控制項或 iframe 中。
    
- **XML 資訊清單**，會指定增益集網頁的 URL，和任何存取需求、設定及增益集的功能。這個檔案儲存在用戶端。OneNote 的增益集，使用與其他 Office 增益集相同的[資訊清單](https://dev.office.com/docs/add-ins/overview/add-in-manifests)格式。

**Office 增益集 = 資訊清單 + 網頁**

![Office 增益集是由資訊清單和網頁所組成](../../images/onenote-add-in.png)

### 使用 JavaScript API

增益集會使用主應用程式的執行階段內容來存取 JavaScript API。API 有兩層︰ 

- **豐富 API** 適用於 OneNote 特定的作業，透過 **Application** 物件來存取。
- **一般 API** 由所有 Office 應用程式共用，透過 **Document** 物件來存取。

#### 透過 *Application* 物件來存取豐富 API

使用 **Application** 物件來存取 OneNote 的物件，如 **Notebook**、**Section** 和 **Page**。使用豐富 API，您可以對 Proxy 物件執行批次作業。基本流程如下所示： 

1- 從內容取得應用程式執行個體。

2- 建立代表您想要使用的 OneNote 物件的 Proxy。藉由讀取和寫入其屬性和呼叫其方法，以同步方式與 Proxy 物件互動。 

3- 在 Proxy 上呼叫 **load**，以填入參數中所指定的屬性值。這個呼叫會加入至命令的佇列。 

   對 API 的方法呼叫 (例如 `context.application.getActiveSection().pages;`) 也會加入至佇列。
    
4- 按照加入佇列的順序，呼叫 **context.sync** 以執行所有佇列中的命令。這會同步處理您正在執行的指令碼和實際物件之間的狀態，並擷取已載入的 OneNote 物件屬性，以用於指令碼。您可以使用傳回的 Promise 物件來鏈結其他動作。

例如： 

```
    function getPagesInSection() {
        OneNote.run(function (context) {
            
            // Get the pages in the current section.
            var pages = context.application.getActiveSection().pages;
            
            // Queue a command to load the id and title for each page.            
            pages.load('id,title');
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    
                    // Read the id and title of each page. 
                    $.each(pages.items, function(index, page) {
                        var pageId = page.id;
                        var pageTitle = page.title;
                        console.log(pageTitle + ': ' + pageId); 
                    });
                })
                .catch(function (error) {
                    app.showNotification("Error: " + error);
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
        });
    }
```

在 [API 參考](../../reference/onenote/onenote-add-ins-javascript-reference.md)，中您可以找到支援的 OneNote 物件和作業。

### 透過 *Document* 物件來存取一般 API。

使用 **Document** 物件來存取一般 API，例如 [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) 和 [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) 方法。 

例如：  

```
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```
OneNote 增益集只支援下列的一般 API：

| API | 附註 |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) | 只有 **Office.CoercionType.Text** 和 **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142145.aspx) | 只有 **Office.CoercionType.Text**、**Office.CoercionType.Image** 和 **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(name);](https://msdn.microsoft.com/en-us/library/office/fp142180.aspx) | 只有內容增益集支援設定 | 
| [Office.context.document.settings.set(name, value);](https://msdn.microsoft.com/en-us/library/office/fp161063.aspx) | 只有內容增益集支援設定 | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

一般情況下，您只使用一般 API 來執行豐富 API 中所不支援的一些動作。若要深入瞭解如何使用一般的 API，請參閱 Office 增益集[文件](https://dev.office.com/docs/add-ins/overview/office-add-ins)和[參考](https://dev.office.com/reference/add-ins/javascript-api-for-office)。


<a name="om-diagram"></a>
## OneNote 物件模型圖 
下圖代表 OneNote JavaScript API 中目前可用的項目。

  ![OneNote 物件模型圖](../../images/onenote-om.png)


## 其他資源

- [建立第一個 OneNote 增益集](onenote-add-ins-getting-started.md)
- [OneNote JavaScript API 參考](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader 範例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 增益集平台概觀](https://dev.office.com/docs/add-ins/overview/office-add-ins)
