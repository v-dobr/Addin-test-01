
# 创建 Access Web 应用的外接程序



本文介绍如何使用 Visual Studio 2015 开发面向 Access Web 应用的 Office 外接程序。

>
  **注意：**有关使用 VBA 开发 Access 解决方案的信息，请参阅 MSDN 上的 [Access](https://msdn.microsoft.com/en-us/library/fp179695.aspx)。

## 先决条件

若要创建面向 Access Web 应用程序的 Office 外接程序，需要具备以下条件：


- Visual Studio 2015

- SharePoint Online 站点（包括在多个 Office 365 订阅中）。此站点必须包括外接程序目录。有关详细信息，请参阅 [在 SharePoint 上设置外接程序目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。


 >**注释**  Office 外接程序将能与承载在 SharePoint Online 或 Office 365 上的 Access Web 应用程序一起使用。Access 2013 桌面应用程序不支持 Office 外接程序。Office.js 版本 1.1 及更高版本支持面向 Access Web 应用程序的 Office 外接程序。


## 在 Visual Studio 中创建项目


1.  打开 Visual Studio，然后在菜单中选择“**文件**”、“**新建**”、“**项目**”。 “**新建项目**”对话框将打开。

2. 在“**新建项目**”对话框的左侧窗格中，依次导航到“**已安装**”、“**模板**”、“**Visual C#**”、“**Office/SharePoint**”、“**Office 外接程序**”。

3. 在“**新建项目**”对话框的中心窗格中，选择“**Office 外接程序**”。

4. 在对话框底部输入项目名称，并选择“**确定**”。 将打开“**创建 Office 外接程序**”对话框。

5. 在“**创建 Office 外接程序**”对话框中，选择“**内容**”，然后选择“**下一步**”。

6. 在“**创建 Office 外接程序**”对话框的下一个屏幕中，选择“**基本外接程序**”或“**文档可视化外接程序**”，并确保已选中“**Access**”复选框。

7. 完成后，选择“**完成**”。 Visual Studio 将创建基于你的工作的入门版项目。

8. 在“**解决方案资源管理器**”中，选择项目的 Web 项目 (**project_name>Web**)。 在属性窗格中找到 **SSL URL** 条目。 这看上去有些类似于：`https://localhost:44314/`。 选择此 URL，并将其复制到剪贴板。 你将很快需要它。

9. 在“**解决方案资源管理器**”中右键单击项目的名称。 在上下文菜单中，选择“**发布**”。 将打开“**发布外接程序**”向导。

10. 在“**发布外接程序**”向导中，选择“**当前配置文件**”旁边的下拉列表。 在此下拉列表中选择“**新建**”。 将打开“**发布 Office 和 SharePoint 外接程序**”对话框。

11. 在此对话框中选择“**创建新配置文件**”，输入配置文件易于识别的名称，然后选择“**完成**”。 “**发布 Office 和 SharePoint 外接程序**”对话框将关闭，并返回到“**发布外接程序**”向导。

12. 在向导中，选择“**打包外接程序**”。 这将完成你的外接程序，以便可以发布到 SharePoint 中的外接程序目录。

13. 在下一个页面中，对于“**你的网站托管在哪里?**”，输入托管你的网站主机的 URL。 它可以是你在步骤 8 中复制的“**SSL URL**”值。 然后选择“**完成**”。

14. 在“**解决方案资源管理器**”中，右键单击项目的清单节点（在项目名称正下方），并选择“**打开文件资源管理器中的文件夹**”。 记下此文件的路径。 稍后你将需要此值。


 >**注意**  你必须使用 Access Web  App 部署外接程序，才能对其进行调试。


## 查看清单和 Home.Html 文件


1. 在 Visual Studio 项目中打开“**Home.html**”文件，并查找引用 office.js 脚本库的行。

```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```
 >**请注意**，脚本标记文件引用的是 Office.js 版本 1.1（及更高版本）。Access 将使用版本 1.1 中引入的 API 元素。

2. 打开与项目关联的清单文件。此文件将在对项目命名后进行命名，并具有扩展名".xml"。

3.  在清单文件中，查找“**主机**”部分并查找“**主机**”条目。

```xml
  <Hosts> <Host Name="Database" /> </Hosts>
```
 >**注意** 应用程序可以在此处使用列出的外接程序。 因为你已在“**创建 Office 外接程序**”对话框中选择“**Access**”，因此列出了“**数据库**”。 如果包括了 Excel，则还有“**工作簿**”相关条目。

Office 和 SharePoint 的外接程序基于 Web。加载项代码必须承载在 Web 服务器上。此示例中，Web 服务器是您的开发计算机。必须运行服务器以供加载项测试使用，这意味着当您在 SharePoint 中查看并调试加载项的同时，Visual Studio 必须运行此加载项。

对于要查找并使用加载项的用户，需要在 SharePoint 中通过加载项目录注册加载项。加载项目录所需的信息包含在清单文件中。

 >**注释**  您将需要创建Access Web 应用程序以承载 Office 外接程序。


## 将加载项发布到 SharePoint Online 目录


1.  登录到 SharePoint Online 或 Office 365，然后通过在页面顶部的 Office 365 工具栏中选择“**管理**”，转到“**SharePoint 管理中心**”。

2. 在“**SharePoint 管理中心**”页的左侧链接栏中，选择“**外接程序**”。 将转到外接程序视图。

3. 在页面的中心窗格中，选择“**外接程序目录**”。 将转到“**目录**”页。

4. 在“**目录**”页面上，选择“**分发 Office 外接程序**”。 将转到名为“**Office 外接程序**”的目录页，此页列出了所有已安装的 Office 外接程序。

5. “**Office 外接程序**”页面顶部，选择“**新建外接程序**”。 将显示“**添加文档**”对话框。

6. 在“**添加文档**”对话框中，选择“**浏览**”，然后转到 Visual Studio 项目中清单文件的位置。 如果之前已复制清单文件的地址，则可以将其粘贴到此目录中。

7. 选择项目中的清单文件，并选择“**确定**”。 SharePoint 现会将你的外接程序添加到本地 SharePoint 库。


 >**注意**  此过程假定你已在 SharePoint 上创建了一个测试站点。 如果还未创建，你可以从 SharePoint 窗口顶部的“**站点**”选项卡进行创建。 你可以使用现有的 Access Web App（如果有）。


## 创建 Access Web App 以托管外接程序


1. 转到测试站点。 在左侧的链接栏中，选择“**网站内容**”。 这将转到到测试站点的“**网站内容**”页面。

2. 在“**网站内容**”页面上选择“**添加外接程序**”。 将转到“**网站内容 – 你的外接程序**”页。

3. 在“**网站内容 – 你的外接程序**”页中，使用页面顶部的搜索栏搜索“**Access 应用程序**”。

4. 现在应该能看到“**Access 应用程序**”的磁贴。

     >**注意**  请记得这并非你的 Office 外接程序，而是新的 Access Web App。 这个 Access Web App 将托管你的 Office 外接程序。
5. 选择此磁贴将显示“**添加 Access 应用程序**”对话框。 输入 Access应用程序的唯一名称，并选择“**创建**”。 SharePoint 创建应用程序将需要一段时间。 完成后，你将看到“**网站内容**”页中列出的 Access 应用程序，旁边带有“**新建**”标签。

6. 现在需要在 Microsoft Access 2013 的桌面版本中打开 Access应用程序，并向其添加数据，然后再在 SharePoint 中打开和查看。


## 将加载项添加到 Access Web 应用程序


1. 打开 Access Web App。

2. 在 SharePoint 选项卡栏中，选择左上角的齿轮图标。 将会显示菜单。 选择“**Office 外接程序**”菜单项。 将打开“**Office 外接程序**”对话框。

3. 选择“**我的组织**”视图并等待 SharePoint 将可用的 Office 外接程序填入对话框。

    对话框中的外接程序之一应为上一步骤中所注册的 Office 外接程序。 选择该外接程序，以将其插入到 Access Web App。 请记得此应用必须在 Visual Studio 中运行，以使其在 Access Web App 页上检测得到并显示。


## 调试 Office 外接程序

若要调试加载项，请在 Internet Explorer 中，按 F12 或选择浏览器选项卡栏中的齿轮图标（不是 SharePoint 页上的齿轮图标）。将显示 Internet Explorer 11 提供的 F12 调试工具。如果您使用的是其他浏览器，请检查浏览器文档以确认输入调试模式的方式。

此时，您可以设置断点、逐步调试 JavaScript 代码、浏览 DOM 和修改代码以确定将更改显示在面向 Access Web 应用程序的 Office 外接程序中。请参阅 [使用 F12 开发人员工具](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29)以了解详细信息。


## 后续步骤

下载示例 [Office 365：在 Access Web 应用程序中绑定和操作数据](https://code.msdn.microsoft.com/officeapps/Office-365-Bind-and-4876274e)以详细了解如何实现在 Access Web 应用程序中操作数据的 Office 外接程序。


## 其他资源



- [了解加载项的 JavaScript API](../develop/understanding-the-javascript-api-for-office.md)

- [适用于 Office 的 JavaScript API](../../reference/javascript-api-for-office.md)

