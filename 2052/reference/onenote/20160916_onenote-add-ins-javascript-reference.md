# OneNote 外接程序 JavaScript API 参考

*适用于：OneNote Online*

下面的链接显示了 API 中可用的高级 OneNote 对象。每个对象页面链接包含对象可用的属性、关系和方法的描述。如需了解详细信息，请浏览下面的链接。 
    
- [应用程序](application.md)：用来访问所有全局可寻址的 OneNote 对象的顶级对象，例如活动笔记本和活动分区。

- [笔记本](notebook.md)：一个笔记本。笔记本包含分区组合和分区。

   - [NotebookCollection](notebookcollection.md)：笔记本的集合。

- [SectionGroup](sectiongroup.md)：一个分区组。分区组包含分区组和分区。

   - [SectionGroupCollection](sectiongroupcollection.md)：分区组的集合。

- [Section](section.md)：一个分区。分区包含页面。

   - [SectionCollection](sectioncollection.md)：分区的集合。

- [Page](page.md)：一个页面。页面包含 PageContent 对象。

   - [PageCollection](pagecollection.md)：页面的集合。

- [PageContent](pagecontent.md)：页面上包含内容类型的顶级地区，例如 Outline 或 Image。可在页面上为 PageContent 对象分配一个位置。

   - [PageContentCollection](pagecontentcollection.md)：PageContent 对象的集合，表示页面的内容。

- [Outline](outline.md)：Paragraph 对象的容器。Outline 是 PageContent 对象的直接子级。

- [Image](image.md)：Image 对象。Image 可以是 PageContent 对象或 Paragraph 的直接子级。

- [Paragraph](paragraph.md)：页面上可见内容的容器。Paragraph 是 Outline 的直接子级。

  - [ParagraphCollection](paragraphcollection.md)：Outline 中 Paragraph 对象的集合。

- [RichText](richtext.md)：RichText 对象。

- [表格](table.md)：TableRow 对象的容器。

- [TableRow](tablerow.md)：TableCell 对象的容器。

  - [TableRowCollection](tablerowcollection.md)：表中 TableRow 对象的集合。
 
- [TableCell](tablecell.md)：段落对象的容器。

  - [TableCellCollection](tablecellcollection.md)：TableRow 中 TableCell 对象的集合。
        
## 其他资源

- [OneNote JavaScript API 编程概述](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [生成你的第一个 OneNote 外接程序](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 外接程序平台概述](https://dev.office.com/docs/add-ins/overview/office-add-ins)
