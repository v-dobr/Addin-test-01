
# 在 iPad 和 Mac 上调试 Office 外接程序

您可以使用 Visual Studio 开发和调试 Windows 上的外接程序。但是，无法使用它调试 iPad 或 Mac 上的外接程序。由于外接程序使用 HTML 和 Javascript 开发，它们应旨在跨平台工作，但不同浏览器呈现您的 HTML 的方式可能存在细微差异。本文介绍如何调试在 iPad 或 Mac 上运行的外接程序。 

## 使用 Vorlon.js 进行调试 

Vorlon.js 是网页的调试程序，与 F12 工具类似，它设计为远程工作，让您可以跨不同设备调试网页。有关详细信息，请参阅 [Vorlon 网站](http://www.vorlonjs.com)。  

安装和设置 Vorlon： 

1.  如果尚未安装，请安装 [Node.js](https://nodejs.org)和 [Git](https://git-scm.com/)。 

2.  通过以下命令使用 Git 安装 Vorlon：`git clone https://github.com/MicrosoftDX/Vorlonjs.git`。

3.  通过 `npm install` 安装依赖项。

4.  外接程序要求使用 HTTPS，因此其使用的任何脚本扩展也必须是 HTTPS，包括 Vorlon 脚本。 因此，必须将 Vorlon 配置为使用 SSL，从而通过外接程序使用 Vorlon。 在安装 Vorlon 的文件夹下，转到 /Server 文件夹并编辑 config.json 文件。 将 **useSSL** 属性更改为 **true**。 此时，还可以为 Office 外接程序启用该插件（将“已启用”属性更改为 true）。 

5.  使用命令 `sudo vorlon` 运行 Vorlon 服务器。 

6.  打开浏览器窗口，然后转到 Vorlon 界面 [http://localhost:1337](http://localhost:1337)。 信任安全证书，应会提示你执行此操作。 还可以在 /Server/cert 下的 Vorlon 文件夹中找到该安全证书。 

7.  向外接程序的 home.html 文件（或主 HTML 文件）的 `<head>` 部分添加以下脚本标记：
```    
<script src="https://localhost:1337/vorlon.js"></script>    
```  

现在，不管您何时在设备上打开外接程序，都会显示在 Vorlon 的客户端列表中（在 Vorlon 界面的左边）。您可以远程突出显示 DOM 元素、远程执行命令等。  

![显示 Vorlon.js 界面的快照](../../images/vorlon_interface.png)

Office 插件为 Office.js 添加额外的功能，例如探索对象模型和执行 Office.js 调用。 
