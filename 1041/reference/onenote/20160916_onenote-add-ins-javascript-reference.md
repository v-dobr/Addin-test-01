# OneNote アドインの JavaScript API リファレンス

*適用対象:OneNote Online*

以下のリンクは、API で使用できる高レベルのOneNote オブジェクトを示しています。オブジェクトのページの各リンクには、オブジェクトで使用できるプロパティ、リレーションシップ、メソッドの説明が含まれています。詳しくは、以下のリンクをご確認ください。 
    
- [Application](application.md):グローバルにアドレス可能な OneNote オブジェクト (アクティブなノートブック、アクティブなセクションなど) すべてへのアクセスに使用する最上位のオブジェクトです。

- [Notebook](notebook.md):ノートブックです。ノートブックには、セクション グループとセクションが含まれます。

   - [NotebookCollection](notebookcollection.md):ノートブックのコレクションです。

- [SectionGroup](sectiongroup.md):セクション グループです。セクション グループには、セクション グループとセクションが含まれます。

   - [SectionGroupCollection](sectiongroupcollection.md):セクション グループのコレクションです。

- [Section](section.md):セクションです。セクションには、ページが含まれます。

   - [SectionCollection](sectioncollection.md):セクションのコレクションです。

- [Page](page.md):ページです。ページには、PageContent オブジェクトが含まれます。

   - [PageCollection](pagecollection.md):ページのコレクションです。

- [PageContent](pagecontent.md):Outline や Image などのコンテンツの種類を含むページの最上位の領域です。PageContent オブジェクトは、ページ上の位置を指定できます。

   - [PageContentCollection](pagecontentcollection.md):PageContent オブジェクトのコレクションで、ページのコンテンツを表します。

- [Outline](outline.md):Paragraph オブジェクトのコンテナーです。Outline は、PageContent オブジェクトの直接の子です。

- [Image](image.md):Image オブジェクトです。Image は、PageContent オブジェクトまたは Paragraph の直接の子にすることができます。

- [Paragraph](paragraph.md):ページに表示されるコンテンツのコンテナーです。Paragraph は、Outline の直接の子です。

  - [ParagraphCollection](paragraphcollection.md):Outline 内の Paragraph オブジェクトのコレクションです。

- [RichText](richtext.md):RichText オブジェクトです。

- [Table](table.md):TableRow オブジェクトのコンテナーです。

- [TableRow](tablerow.md):TableCell オブジェクトのコンテナーです。

  - [TableRowCollection](tablerowcollection.md):Table 内の TableRow オブジェクトのコレクションです。
 
- [TableCell](tablecell.md):Paragraph オブジェクトのコンテナーです。

  - [TableCellCollection](tablecellcollection.md):TableRow 内の TableCell オブジェクトのコレクションです。
        
## その他のリソース

- [OneNote の JavaScript API のプログラミングの概要](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [最初の OneNote 用アドインをビルドする](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](https://dev.office.com/docs/add-ins/overview/office-add-ins)
