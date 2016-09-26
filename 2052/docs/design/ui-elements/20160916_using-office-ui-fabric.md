
#在 Office 外接程序中使用 Office UI 结构

如果您要构建 Office 外接程序，我们建议您使用 [Office UI 结构](https://github.com/OfficeDev/Office-UI-Fabric)创建用户体验。以下步骤将向您演示使用结构的基础知识。  

##1.设置结构
将以下行添加到 HTML 的 head 部分，以引用 CDN 中的结构。

     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">


##2.使用结构图标和字体
使用图标变得非常简单。您只需使用“i”元素并参考相应的类即可。您可以通过更改字体大小控制图标的大小。

    <i class="ms-Icon ms-Icon--group" style="font-size:xx-large" aria-hidden="true"></i>


##3.使用简单组件的样式
结构提供了各种 UI 元素（如按钮和复选框）的样式。您只需引用适当的类来添加相应的样式即可，如下面的示例中所示。

    <button class="ms-Button" id="get-data-from-selection">
    <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
    <span class="ms-Button-label">Get Data from selection</span>
    <span class="ms-Button-description">Get Data from the document selection</span>
    </button>

##4.将组件与示例行为一起使用
Fabric 包括一些支持行为（例如在单击时会发生什么情况）的组件。 为了帮助你入门，Fabric 2.6.1 以 JQuery UI 插件的形式提供了一些**示例代码**供你使用。 你还可以使用运行所需的任何其他框架。 如果选择使用示例，请注意，代码不随 CDN 分发，因此你必须从 [Fabric GitHub 项目](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1) 的 2.6.1 版本下载，引用它，然后在代码中进行初始化。 

例如，要使用搜索框组件，请执行以下操作：

1. 从 [GitHub](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1/src/components/SearchBox) 下载搜索框组件。
2. 将以下引用添加到代码中：`<script src="SearchBox/Jquery.SearchBox.js"></script>`
3. 确保此行可在页面加载时执行，以初始化组件：`$(".ms-SearchBox").SearchBox();`。我们建议您将此行包含在外接程序的 `Office.Initialize` 块中。     

**注意：**如果不打算使用所有 Fabric 组件，可以通过选择改为托管每个组件的单个 CSS 文件来减少下载的资源。 可以从 [Fabric 2.6.1 GitHub 存储库](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1) 中的组件文件夹获取 CSS 文件。 


##后续步骤
如果您正在寻找介绍如何使用结构的端到端示例，我们已经为您准备好了。请参阅[Office 外接程序结构 UI 示例](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)。您还可以浏览交互式 [Office UI 结构](https://github.com/OfficeDev/Office-UI-Fabric)网站。

