# 生成你的第一个 OneNote 外接程序

本文介绍生成可将一些文本添加到 OneNote 页面的简单任务窗格外接程序的步骤。

下图显示将创建的外接程序。

   ![构建自此演练的 OneNote 外接程序](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## 步骤 1：设置开发环境
1- 按照这些 [安装说明](https://dev.office.com/docs/add-ins/get-started/create-an-office-add-in-using-any-editor)(#安装说明) 安装 Yeoman Office 生成器及其系统必备组件。

   当您没有 Visual Studio 或您想使用技术而非纯 HTML、CSS 和 JavaScript 时，Yeoman Office 生成器可以轻松地创建外接程序项目。它还提供了用于测试的本地 Gulp Web 服务器的快捷访问。 

   >你可以有选择性地 [使用 Visual Studio](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio)(#使用-visual-studio) 以创建你的项目文件，但不会获得内置 Gulp 服务器支持。

<a name="create-project"></a>
## 步骤 2：创建外接程序项目 
1- 创建一个名为 *onenote 外接程序*的本地文件夹。

2- 打开 **cmd** 提示符，并导航到 **onenote 外接程序**文件夹。运行 `yo office` 命令，如下所示。

```
C:\your-local-path\onenote add-in\> yo office
```
>这些说明使用 Windows 命令提示符，但它们也同样适用于其他 shell 环境。 

3- 使用以下选项以创建项目。

| 选项 | 值 |
|:------|:------|
| 项目名称 | OneNote 外接程序 |
| 项目的根文件夹 | （接受默认值） |
| Office 项目类型 | 任务窗格外接程序 |
| 受支持的 Office 应用程序 | （选择 any--我们将稍后添加一个 OneNote 主机） |
| 要使用的技术 | HTML、CSS 和 JavaScript |

<a name="manifest"></a>
## 步骤 3：配置外接程序清单 
1- 在您的项目文件中打开 **manifest-onenote-add-in.xml**。添加以下行至 **Hosts** 部分。这将指定您的外接程序支持 OneNote 主机应用程序。

```
<Host Name="Notebook" />
```

请注意，已经为的 Gulp Web 服务器设置 **SourceLocation**。

```
<SourceLocation DefaultValue="https://localhost:8443/app/home/home.html"/>
```

<a name="develop"></a>
## 步骤 4：开发外接程序
您可以使用任何文本编辑器或 IDE 开发外接程序。如果您尚未尝试过 Visual Studio 代码，可以在 Linux、Mac OSX 和 Windows 上[免费下载](https://code.visualstudio.com/)。

1- 在 **app/home** 文件夹中打开 *home.html*。 

2- 编辑对 Office JavaScript API 和 [Office UI 结构](http://dev.office.com/fabric)样式及组件的引用。

   a.取消评论到 fabric.components.min.css 的链接。

   b.根据以下对 *beta* 版本的引用，替换对 Office.js 的脚本引用。

```
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

你的 Office 引用将如下所示。

```
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.min.css" rel="stylesheet">
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css" rel="stylesheet">
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

3- 用以下代码替换 `<body>` 元素。 这将添加一个使用 [Office UI Fabric 组件](http://dev.office.com/fabric/components)的文本区域和按钮。 **响应网格**布局源自 [Office UI Fabric 样式](http://dev.office.com/fabric/styles)集。 

```
<body class="ms-font-m">
    <div class="home flex-container">
        <div class="ms-Grid">
            <div class="ms-Grid-row ms-bgColor-themeDarker">
                <div class="ms-Grid-col">
                    <span class="ms-font-xl ms-fontColor-themeLighter ms-fontWeight-semibold">OneNote Add-in</span>
                </div>
            </div>
        </div>
        <br />
        <div class="ms-Grid">
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <label class="ms-Label">Enter content here</label>
                    <div class="ms-TextField ms-TextField--placeholder">
                        <textarea id="textBox" rows="5"></textarea>
                    </div>
                </div>
            </div>
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <div class="ms-font-m ms-fontColor-themeLight header--text">
                        <button class="ms-Button ms-Button--primary" id="addOutline">
                            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
                            <span class="ms-Button-label">Add outline</span>
                            <span class="ms-Button-description">Adds the content above to the current page.</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
```

4- 在 **app/home** 文件夹中打开 *home.js*编辑 **Office.initialize** 函数以添加一个点击事件至**添加 outline** 按钮，如下所示。 

```
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
    $(document).ready(function () {
        app.initialize();

        // Set up event handler for the UI.
        $('#addOutline').click(addOutlineToPage);
    });
};
```
 
5- 用以下 **addOutlineToPage** 方法替换 **getDataFromSelection** 方法。这将从文本区域中获取内容并将其添加至页面。

```
// Add the contents of the text area to the page.
function addOutlineToPage() {        
    OneNote.run(function (context) {
       var html = '<p>' + $('#textBox').val() + '</p>';

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to load the page with the title property.             
        page.load('title'); 

        // Add an outline with the specified HTML to the page.
        var outline = page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function() {
                console.log('Added outline to page ' + page.title);
            })
            .catch(function(error) {
                app.showNotification("Error: " + error); 
                console.log("Error: " + error); 
                if (error instanceof OfficeExtension.Error) { 
                    console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                } 
            }); 
        });
}
```

<a name="test"></a>
## 步骤 5：在 OneNote Online 上测试外接程序
1- 运行 Gulp Web 服务器。  

   a. 在 **onenote 外接程序**文件夹中打开 **cmd** 提示符。 

   b. 运行 `gulp serve-static` 命令，如以下所示。

```
C:\your-local-path\onenote add-in\> gulp serve-static
```

2- 安装 Gulp Web 服务器的自签名证书作为受信任的证书。对于用 Yeoman Office 生成器创建的外接程序项目，您只需在电脑上进行一次此操作。

   a.导航至托管的外接程序页面。默认情况下，这与您清单中的 URL 相同：

```
https://localhost:8443/app/home/home.html
```

   b. 安装证书作为受信任的证书。 有关详细信息，请参阅 [添加自签名的证书作为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/docs/trust-self-signed-cert.md)。

3- 在 OneNote Online 中，打开一个笔记本。

4-选择“**插入 > Office 外接程序**”。 这将打开 Office 外接程序对话框。
  - 如果使用消费者帐户登录，请选择“**我的外接程序**”选项卡，然后选择“**上载我的外接程序**”。
  - 如果使用工作或学校帐户登录，请选择“**我的组织**”选项卡，然后选择“**上载我的外接程序**”。 
  
  以下图像显示消费者笔记本的“**我的外接程序**”选项卡。

  ![显示“我的外接程序”选项卡的 Office 外接程序对话框](../../images/onenote-office-add-ins-dialog.png)
  
  >**注意**：若要启用“**Office 外接程序**”按钮，请在 OneNote 页内单击。

5- 在“上载外接程序”对话框中，浏览至项目文件中的 **manifest-onenote-add-in.xml**，然后选择“**上载**”。 测试时，你的清单文件可以在本地存储。

6- 该外接程序在 OneNote 页面旁的 iFrame 中打开。 在文本区域中输入一些文本，然后选择“**添加边框**”。 您输入的文本将添加至页面。 

## 故障排除和提示
- 您可以使用浏览器的开发者工具调试外接程序。当您在 Internet Explorer 或 Chrome 中使用 Gulp Web 服务器并进行调试时，您可以本地保存您的更改，然后仅刷新外接程序的 iFrame。

- 当您检查 OneNote 对象时，目前可用的属性显示实际值。需要加载的属性显示 *未定义*。展开 `_proto_` 节点以查看在对象上被定义但未加载的属性。

      ![Unloaded OneNote object in the debugger](../../images/onenote-debug.png)

- 如果您的外接程序使用任何 HTTP 资源，则需要启用浏览器中的混合内容。生产外接程序应当仅使用安全 HTTPS 资源。

## 其他资源

- [OneNote JavaScript API 编程概述](onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API 参考](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 外接程序平台概述](https://dev.office.com/docs/add-ins/overview/office-add-ins)
