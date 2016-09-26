
# 在 iPad 和 Mac 上偵錯 Office 增益集

您可以在 Windows 上使用 Visual Studio 來開發並偵錯增益集，但您無法使用它來偵錯 iPad 或 Mac 上的增益集。由於增益集是使用 HTML 和 Javascript 開發，因此設計為跨平台使用，但不同瀏覽器呈現 HTML 的方式可能有細微差異。本文說明如何偵錯 iPad 或 Mac 上執行的增益集。 

## 以 Vorlon.js 偵錯 

Vorlon.js 是網頁的偵錯工具，類似於 F12 工具，設計用來遠端工作，並可讓您跨不同的裝置偵錯網頁。如需詳細資訊，請參閱 [Vorlon 網站](http://www.vorlonjs.com)。  

若要安裝和設定 Vorlon： 

1.  安裝 [Node.js](https://nodejs.org) 和 [Git](https://git-scm.com/) (如果尚未安裝)。 

2.  搭配使用下列命令與 git 來安裝 Vorlon︰`git clone https://github.com/MicrosoftDX/Vorlonjs.git`

3.  安裝與 `npm install` 的相依性。

4.  增益集需要 HTTPS，所以它們使用的延伸和任何指令碼也必須是 HTTPS，包括 Vorlon 指令碼。 因此，您必須設定 Vorlon 使用 SSL，以便使用 Vorlon 與增益集。 在您安裝 Vorlon 的資料夾底下，移至 /Server 資料夾，並且編輯 config.json 檔案。 變更 **useSSL** 屬性為 **true**。 當您進入後，您也可以啟用 Office 增益集的外掛程式 (將其「啟用」屬性變更為 true)。 

5.  使用命令 `sudo vorlon` 執行 Vorlon 伺服器。 

6.  開啟瀏覽器視窗並移至 [http://localhost:1337](http://localhost:1337)，也就是 Vorlon 介面。 信任安全性憑證，系統應該會提示您執行。 您也可以在 /Server/cert 下的 Vorlon 資料夾中找到安全性憑證。 

7.  將下列指令碼標記新增至增益集的 Home.html 檔 (或主要 HTML 檔案) 的 `<head>` 區段︰
```    
<script src="https://localhost:1337/vorlon.js"></script>    
```  

現在，每當在裝置上開啟增益集，它就會顯示在 Vorlon的用戶端清單中 (位於 Vorlon 介面的左邊)。您可以從遠端反白顯示 DOM 元素、從遠端執行命令，以及其他動作等等。  

![顯示 Vorlon.js 介面的螢幕擷取畫面](../../images/vorlon_interface.png)

Office 外掛程式會新增 Office.js 的額外功能，例如瀏覽物件模型和執行 Office.js 呼叫。 
