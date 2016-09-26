# OneNote 增益集 JavaScript API 參考

**適用版本：OneNote Online*

以下連結會顯示 API 中可用的高階 OneNote 物件。每個物件頁面連結都會說明物件可用的屬性、關聯性和方法。請瀏覽下列連結，了解詳細資訊。 
    
- [](application.md)：最上層物件，用以存取所有的全域可定址 OneNote 物件，例如使用中的筆記本中和使用中的節。

- [筆記本](notebook.md)：筆記本。筆記本包含節群組和節。

   - [NotebookCollection](notebookcollection.md):筆記本的集合。

- [SectionGroup](sectiongroup.md):節群組。節群組包含節群組和節。

   - [SectionGroupCollection](sectiongroupcollection.md):節群組的集合。

- [Section](section.md):章節。節包含頁面。

   - [SectionCollection](sectioncollection.md):節的集合。

- [Page](page.md):頁面。頁面包含 PageContent 物件。

   - [PageCollection](pagecollection.md):頁面的集合。

- [PageContent](pagecontent.md):頁面上的最上層區域，包含例如 Outline 或 Image 內容類型。可以指派位置給 PageContent 物件。

   - [PageContentCollection](pagecontentcollection.md):PageContent 物件的集合，代表頁面的內容。

- [Outline](outline.md):Paragraph 物件的容器。大綱是 PageContent 物件的直接子項。

- [Image](image.md):影像物件。影像可以是 PageContent 物件或段落的直接子項。

- [Paragraph](paragraph.md):在頁面上可見內容的容器。段落是大綱的直接子項。

  - [ParagraphCollection](paragraphcollection.md):在大綱中的 Paragraph 物件集合。

- [RichText](richtext.md)RichText 物件。

- [Table](table.md)TableRow 物件的容器。

- [TableRow](tablerow.md)TableCell 物件的容器。

  - [TableRowCollection](tablerowcollection.md)Table 中的 TableRow 物件的集合。
 
- [TableCell](tablecell.md)Paragraph 物件的容器。

  - [TableCellCollection](tablecellcollection.md)TableRow 中的 TableCell 物件的集合。
        
## 其他資源

- [OneNote JavaScript API 程式設計的概觀](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [建立第一個 OneNote 增益集](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Rubric Grader 範例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 增益集平台概觀](https://dev.office.com/docs/add-ins/overview/office-add-ins)
