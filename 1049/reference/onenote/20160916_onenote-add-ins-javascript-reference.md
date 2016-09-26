# Справочник по API JavaScript для надстроек OneNote

*Область применения: OneNote Online*

По ссылкам ниже расположены статьи, в которых рассказывается об объектах OneNote высокого уровня, доступных в API. В этих статьях описаны свойства, связи и методы, доступные для соответствующих объектов. Чтобы получить дополнительные сведения, перейдите по этим ссылкам. 
    
- [Application](application.md): объект верхнего уровня, используемый для доступа ко всем глобально адресуемым объектам OneNote, например активной записной книжке и активному разделу.

- [Notebook](notebook.md): записная книжка. Записные книжки содержат группы разделов и разделы.

   - [NotebookCollection](notebookcollection.md): представляет коллекцию записных книжек.

- [SectionGroup](sectiongroup.md): группа разделов. Группы разделов содержат разделы и группы разделов.

   - [SectionGroupCollection](sectiongroupcollection.md): коллекция групп разделов.

- [Section](section.md): раздел. Разделы содержат страницы.

   - [SectionCollection](sectioncollection.md): коллекция разделов.

- [Page](page.md): страница. Страницы содержат объекты PageContent.

   - [PageCollection](pagecollection.md): коллекция страниц.

- [PageContent](pagecontent.md): область верхнего уровня на странице, содержащая контент, например типов Outline или Image. Объекту PageContent можно назначить позицию на странице.

   - [PageContentCollection](pagecontentcollection.md): коллекция объектов PageContent, представляющая содержимое страницы.

- [Outline](outline.md): контейнер для объектов Paragraph. Объект Outline — прямой потомок объекта PageContent.

- [Image](image.md): объект Image. Объект Image может быть прямым потомком объекта PageContent или объекта Paragraph.

- [Paragraph](paragraph.md): Контейнер для содержимого, отображаемого на странице. Объект Paragraph — прямой потомок объекта Outline.

  - [ParagraphCollection](paragraphcollection.md): коллекция объектов Paragraph в объекте Outline.

- [RichText](richtext.md): объект RichText.

- [Table](table.md): контейнер для объектов TableRow.

- [TableRow](tablerow.md): контейнер для объектов TableCell.

  - [TableRowCollection](tablerowcollection.md): Коллекция объектов TableRow в объекте Table.
 
- [TableCell](tablecell.md): контейнер для объектов Paragraph.

  - [TableCellCollection](tablecellcollection.md) коллекция объектов TableCell в объекте TableRow.
        
## Дополнительные ресурсы

- [Обзор создания кода с помощью API JavaScript для OneNote](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [Создание первой надстройки OneNote](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
