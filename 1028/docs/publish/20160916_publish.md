
# 部署及發佈 Office 增益集


您可以使用以下幾種方法中的任一種方法來部署 Office 增益集，以供測試之用或散發給使用者︰

- [側載](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) - 能做為部署程序的一部分，測試增益集在 Windows、Office Online、iPad 或 Mac 上的執行狀況。
- [SharePoint 目錄](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) - 能做為開發程序的一部分，測試增益集或將增益集散發給組織中的使用者。
- [Office 365 系統管理中心預覽](https://support.office.com/en-ie/article/Deploy-Office-Add-Ins-in-Office-365-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE) - 能用來將增益集散發給組織中的使用者。
- [Office 市集] - 能用來將增益集公開散發給使用者。

可用的選項視您鎖定的 Office 主應用程式和建立的增益集類型而定。

### Word、Excel 及 PowerPoint 增益集的部署選項

| 擴充點            | 側載 | SharePoint 目錄 | Office 365 系統管理中心預覽 | Office 市集 |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| 內容         | X           | X                  | X                               | X            |
| 工作窗格       | X           | X                  | X                               | X            |
| 命令         | X           |                    | X                               | X            |

> **附註：**Office 2016 for Mac 不支援 SharePoint 目錄。 若要將 Office 增益集部署到 Mac 用戶端，您必須將它們提交到 [Office 市集]。    

### Outlook 增益集的部署選項

| 擴充點     | 側載 | Exchange Server | Office 市集 |
|:---------|:-----------:|:---------------:|:------------:|
| 郵件應用程式 | X           | X               | X            |
| 命令  | X           | X               | X            |

若要擴大增益集的可及範圍，請確認它能跨越平台運作。 Windows、Mac、Web、iOS 及 Android 支援 Office 增益集。 如需各平台支援之功能的概觀，請參閱 [Office 增益集主應用程式和平台可用性]。   

如需授權您的 Office 市集增益集的詳細資訊，請參閱[授權您的增益集](https://msdn.microsoft.com/EN-US/library/office/jj163257.aspx)。

如需使用者如何取得、插入及執行增益集的相關資訊，請參閱[開始試用 Office 增益集](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)。

## 其他資源

- [Office 增益集主應用程式和平台可用性]
- [部署和安裝 Outlook 增益集以進行測試](../outlook/testing-and-tips.md) 
- [將增益集和 Web 應用程式提交至 Office 市集][Office 市集]
- [Office 增益集的設計指導方針](../design/add-in-design)
- [建立有效的 Office 市集增益集](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [疑難排解 Office 增益集的使用者錯誤](../testing/testing-and-troubleshooting.md)

[Office 市集]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office 增益集主應用程式和平台可用性]: http://dev.office.com/add-in-availability
