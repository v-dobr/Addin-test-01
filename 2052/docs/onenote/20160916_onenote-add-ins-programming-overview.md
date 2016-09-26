# OneNote JavaScript API 编程概述

OneNote 引入了适用于 OneNote Online 外接程序的 JavaScript API。 你可以创建任务窗格外接程序、内容外接程序和与 OneNote 对象交互并连接至 Web 服务或其他基于 Web 的资源的外接程序命令。

外接程序包含两个基本组件：

- **Web 应用程序**包含网页和任何所需的 JavaScript、CSS 或其他文件。这些文件托管在 Web 服务器或 Web 托管服务上，例如 Microsoft Azure。在 OneNote Online 中，Web 应用程序在浏览器控件或 iframe 中显示。
    
- **XML 清单**指定外接程序网页的 URL 和适用于外接程序的任何访问要求、设置和功能。此文件存储在客户端上。OneNote 外接程序使用与其他 Office 外接程序相同的 [清单](https://dev.office.com/docs/add-ins/overview/add-in-manifests)格式。

**Office 外接程序 = 清单 + 网页**

![Office 外接程序包含清单和网页](../../images/onenote-add-in.png)

### 使用 JavaScript API

外接程序使用托管应用程序的运行时上下文以访问 JavaScript API。API 有两层： 

- 适用于 OneNote 特定操作的**丰富 API**，通过 **Application** 对象访问。
- 跨 Office 应用程序分享的**通用 API**，通过**Document** 对象访问。

#### 通过 *Application* 对象访问丰富 API。

使用 **Application** 对象以访问 OneNote 对象，例如 **Notebook**、**Section** 和 **Page**。通过丰富 API，您可在代理对象上运行批处理操作。基本流程类似如下： 

1- 从上下文中获取应用程序实例。

2- 创建你想要使用的表示 OneNote 对象的代理。 通过读取和写入代理对象的属性和调用其方法，你可以与其同步交互。 

3- 调用代理上的 **load** 以使用在参数中指定的属性值填充它。 此调用将添加至命令队列中。 

   对 API 的方法调用（如 `context.application.getActiveSection().pages;`）也将添加到队列中。
    
4- 调用 **context.sync** 以按它们已排队的顺序运行所有排队的命令。 这将同步你正在运行的脚本和真实对象之间的状态，并通过检索已加载的用于你的脚本的 OneNote 对象的属性实现。 您可以使用返回的 promise 对象以链接其他操作。

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

您可以在 [API 参考](../../reference/onenote/onenote-add-ins-javascript-reference.md) 中找到受支持的 OneNote 对象和操作。

### 通过 *Document* 对象访问通用 API

使用 **Document** 对象以访问通用 API，例如 [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) 和 [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) 方法。 

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
OneNote 外接程序仅支持以下通用 API：

| API | 注释 |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) | 仅限 **Office.CoercionType.Text** 和 **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142145.aspx) | 仅限 **Office.CoercionType.Text**、**Office.CoercionType.Image**  和 **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(name);](https://msdn.microsoft.com/en-us/library/office/fp142180.aspx) | 设置仅受内容外接程序支持 | 
| [Office.context.document.settings.set(name, value);](https://msdn.microsoft.com/en-us/library/office/fp161063.aspx) | 设置仅受内容外接程序支持 | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

一般情况下，您仅能使用通用 API 执行在丰富 API 中不支持的操作。若要了解有关使用通用 API 的详细信息，请参阅 Office 外接程序[文档](https://dev.office.com/docs/add-ins/overview/office-add-ins)和[引用](https://dev.office.com/reference/add-ins/javascript-api-for-office)。


<a name="om-diagram"></a>
## OneNote 对象模型图 
下图表示了 OneNote JavaScript API 中当前可用的内容。

  ![OneNote 对象模型图](../../images/onenote-om.png)


## 其他资源

- [生成你的第一个 OneNote 外接程序](onenote-add-ins-getting-started.md)
- [OneNote JavaScript API 参考](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 外接程序平台概述](https://dev.office.com/docs/add-ins/overview/office-add-ins)
