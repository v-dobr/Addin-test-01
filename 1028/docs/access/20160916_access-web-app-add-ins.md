
# 建立 Access Web App 的增益集



本文將說明如何使用 Visual Studio 2015 來開發以 Access Web App 目標的 Office 增益集。

>
  **附註：**如需使用 VBA 開發 Access 的解決方案的詳細資訊，請參閱 MSDN 上的 [Access](https://msdn.microsoft.com/en-us/library/fp179695.aspx)。

## 必要條件

若要建立以 Access Web App 為目標的 Office 增益集，您需要︰


- Visual Studio 2015

- 一個 SharePoint Online 網站 (許多 Office 365 訂閱中均包含)。這個站台必須有增益集目錄。如需詳細資訊，請參閱[在 SharePoint 上設定增益集目錄](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。


 >**附註：**Office 增益集可用於主控在 SharePoint Online 或 Office 365 上的 Access Web App。Access 2013 桌面應用程式不支援 Office 增益集。以 Access Web App 為目標的 Office 增益集可受 Office.js 1.1 版和以後版本支援。


## 在 Visual Studio 中建立專案


1.  開啟 Visual Studio，然後在功能表中，依序選擇 [檔案、[新增 和 [專案。 [新增專案 對話方塊將會開啟。

2. 在 [新增專案 對話方塊中，於左邊窗格中，依序導覽至 [已安裝、[範本、[Visual C#、[Office/SharePoint 和 [Office 增益集。

3. 在 [新增專案 對話方塊中，於中間窗格選擇 [Office 增益集。

4. 在對話方塊的底端輸入專案的名稱，然後選擇 [確定。 如此會開啟 [建立 Office 增益集 對話方塊。

5. 在 [建立 Office 增益集 對話方塊中，選擇 [內容，然後選擇 [下一步。

6. 在 [建立 Office 增益集 對話方塊的下一個畫面 中，選擇 [基本增益集 或 [文件視覺效果增益集，並確定已選取 [Access 的核取方塊。

7. 當完成時，選擇 [完成。 Visual Studio 會建立起始專案作為您工作的基礎。

8. 在**方案總管**中，選擇專案的 Web 專案 (**project_name>Web**)。 在屬性窗格中，尋找 **SSL URL** 項目。 看起來應該類似下面的內容︰`https://localhost:44314/`。 選取此 URL，並將它複製到剪貼簿。 您稍後會需要它。

9. 在**方案總管**中，以滑鼠右鍵按一下您專案的名稱。 在內容功能表中，選擇 [發佈。 這會開啟**發佈增益集**精靈。

10. 在**發佈增益集**精靈中，選取 [目前的設定檔旁的下拉式清單。 在此下拉式清單中，選擇 [新增。 如此會開啟 [發佈的 Office 和 SharePoint 增益集 對話方塊。

11. 在此對話方塊中，選擇 [建立新設定檔，為設定檔輸入可辨識的名稱，然後選擇 [完成。 [發佈 Office 和 SharePoint 增益集 對話方塊會關閉，讓您返回 [發佈增益集 精靈。

12. 在精靈中，選擇 [封裝增益集。 這會完成您的增益集，使得可以將它發佈行到 SharePoint 中的增益集目錄。

13. 在下一個頁面中，針對 [您的網站架設在哪裡?，請填入您的網站的 URL。 這可以是您在步驟 8 中複製的 **SSL URL** 值。 然後選擇 [完成。

14. 在**方案總管**中，以滑鼠右鍵按一下專案的資訊清單節點 (就在專案名稱下方)，然後選取 [在檔案總管中開啟資料夾。 記下這個檔案的路徑。 您稍後需要這個值。


 >**附註：**若未使用 Access Web App 部署增益集，您無法對增益集進行偵錯。


## 檢閱資訊清單和 Home.Html 檔案


1. 在 Visual Studio 專案中，開啟 **Home.html** 檔案，並尋找參考 office.js 指令碼程式庫的行。

```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```
 >**附註：**該指令碼標記參考 1.1 版 (和以上) 的 Office.js。Access 會使用 1.1 版中推出的 API 元素。

2. 開啟與您的專案關聯的資訊清單檔。這個檔案將會您的專案名稱命名，且具有副檔名 ".xml"。

3.  在資訊清單檔案中，尋找 **Hosts** 區段，並尋找 **Host** 項目。

```xml
  <Hosts> <Host Name="Database" /> </Hosts>
```
 >**附註：**這是列出可以使用增益集的應用程式的位置。 因為您已在 [建立 Office 增益集 對話方塊中選取 [Access，會列出**資料庫**。 如果您已納入 Excel，則也會有**活頁簿**的項目。

Office和 SharePoint 增益集都是以 Web 為基礎。增益集的程式碼必須在 Web 伺服器上主控。針對此範例，Web 伺服器為您的開發電腦。伺服器必須執行中才能主控要測試的增益集，這表示，當您在 SharePoint 中檢視並偵錯增益集時，Visual Studio 必須在執行增益集。

為了讓使用者尋找和使用增益集，需要在 SharePoint 中的增益集目錄中登錄該增益集。增益集目錄所需的資訊包含在資訊清單檔案中。

 >**附註：**您需要建立 Access Web App，才能主控 Office 增益集。


## 將增益集發佈到 SharePoint Online 目錄


1.  登入 SharePoint Online 或 Office 365，然後透過在頁面頂端的 Office 365 工具列中選擇 [管理 來前往 **SharePoint 系統管理中心**。

2. 於 [SharePoint 系統管理中心 頁面上，在左邊的連結列中選擇 [增益集。 這將帶您前往增益集檢視。

3. 在頁面的中央窗格中，選擇 [增益集目錄。 這將帶您前往 [目錄 頁面。

4. 在 [目錄 頁面上，選擇 [散發 Office 增益集。 這將帶您前往稱為 **Office 增益集**的目錄頁面，其中列出所有已安裝的 Office 增益集。

5. 在 [Office 增益集 頁面頂端，選擇 [新的增益集。 如此將會顯示 [新增文件 對話方塊。

6. 在 [新增文件 對話方塊中，選擇 [瀏覽，然後前往 Visual Studio 專案中資訊清單檔案的位置。 如果您稍早複製了資訊清單檔案的位址，您可以將它貼入這個對話方塊。

7. 選擇您的專案中的資訊清單檔案，然後選擇 [確定。 SharePoint 現在會將增益集新增至本機 SharePoint 程式庫。


 >**附註：**這個程序假設您已在 SharePoint 上建立測試網站。 如果尚未這麼做，您可以從 SharePoint 視窗頂端的 [站台 索引標籤執行這項操作。 您可以使用現有的 Access Web App (如果已有的話)。


## 建立 Access Web App 來主控您的增益集


1. 前往您的測試網站。 在左方的連結列中，選擇 [網站內容。 這將帶您前往測試網站的 [網站內容 頁面。

2. 在 [網站內容 頁面上，選擇 [新增增益集。 這將帶您前往 [網站內容 - 您的增益集 頁面。

3. 在 [網站內容 -您的增益集 頁面上，使用頁面頂端的搜尋列來搜尋 **Access 應用程式**。

4. 您現在應該可以看到 **Access 應用程式**的磚。

     >**附註**  請記住，這不是您的 Office 增益集，它是新的 Access Web App。 此 Access Web App 將主控您的 Office 增益集。
5. 選擇這個磚會帶出 [新增 Access 應用程式 對話方塊。 為您的 Accessapp 輸入唯一名稱，然後選擇 [建立。 SharePoint 建立您的應用程式可能需要一段時間。 完成後，您會看到您的 Accessapp 在 [網站內容 頁面列出，其旁邊有一個**新**標籤。

6. Accessapp 現在會要求您在 Microsoft Access 2013 桌面版本中開啟它，並將資料新增到它，才能在 SharePoint 中將它開啟並檢視。


## 將您的增益集新增至 Access Web App


1. 開啟一個 Access Web App。

2. 在 SharePoint 索引標籤上，選擇左上角的齒輪圖示。 此時會出現一個功能表。 選擇 [Office 增益集 功能表項目。 這會開啟 [Office 增益集 對話方塊。

3. 選擇 [我的組織 檢視，並等候 SharePoint 在對話方塊中填入您可以使用的 Office 增益集。

    對話方塊中其中一個增益集應該是您在先前程序註冊的 Office 增益集。 選擇該增益集以將它插入 Access Web App。 請記住，必須在 Visual Studio 執行的應用程式，才能偵測出來並在 Access Web App 頁面上顯示。


## 偵錯您的 Office 增益集

若要偵錯增益集，請在 Internet Explorer 中按 F12，或在瀏覽器的索引標籤中選擇齒輪圖示列 (不是 SharePoint 頁面的齒輪圖示)。這時會出現 Internet Explorer 11 提供的 F12 偵錯工具。如果您使用其他瀏覽器，請檢查您的瀏覽器文件，以判斷如何進入偵錯模式。

此時，您可以設定中斷點、逐步執行您的 JavaScript 程式碼、瀏覽 DOM，並修改程式碼以確認您的變更會出現在以 Access Web App 為目標的 Office 增益集中。如需詳細資訊，請參閱[使用 F12 開發人員工具](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29)。


## 後續步驟

下載範例 [Office 365：繫結和處理 Access Web App 中的資料](https://code.msdn.microsoft.com/officeapps/Office-365-Bind-and-4876274e)以進一步了解如何實作可操作 Access Web App 中資料的 Office 增益集。


## 其他資源



- [了解適用於增益集的 JavaScript API](../develop/understanding-the-javascript-api-for-office.md)

- [JavaScript API for Office](../../reference/javascript-api-for-office.md)

