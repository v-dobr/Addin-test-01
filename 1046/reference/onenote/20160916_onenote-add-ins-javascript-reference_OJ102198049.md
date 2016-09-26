# Referência de API JavaScript de suplementos do OneNote

*Aplica-se a: OneNote Online*

Os links a seguir mostram os objetos de nível superior do OneNote disponíveis na API. Os link de página dos objetos contêm uma descrição das respectivas propriedades, relações e métodos disponíveis. Acesse os links abaixo para saber mais. 
    
- [Application](application.md): o objeto de nível superior usado para acessar todos os objetos do OneNote globalmente endereçados, como o bloco de anotações ativo e a sessão ativa.

- [Notebook](notebook.md): um bloco de anotações. blocos de anotações contêm grupos de seções e seções.

   - [NotebookCollection](notebookcollection.md): uma coleção de blocos de anotações.

- [SectionGroup](sectiongroup.md): um grupo de seções. Os grupos de seções contêm seções e grupos de seções.

   - [SectionGroupCollection](sectiongroupcollection.md): uma coleção de grupos de seção.

- [Section](section.md): uma seção. As seções contêm páginas.

   - [SectionCollection](sectioncollection.md): uma coleção de seções.

- [Page](page.md): uma página. As páginas contêm objetos PageContent.

   - [PageCollection](pagecollection.md): uma coleção de páginas.

- [PageContent](pagecontent.md): uma região de nível superior em uma página que contém os tipos de conteúdo como estrutura de tópicos ou imagem. Um objeto PageContent pode ser atribuído a uma posição na página.

   - [PageContentCollection](pagecontentcollection.md): uma coleção de objetos PageContent, que representam os conteúdos da página.

- [Outline](outline.md): um contêiner para objetos Paragraph. Uma estrutura de tópicos é um filho direto do objeto PageContent.

- [Image](image.md): um objeto Image. Um Image pode ser um filho direto de um objeto PageContent ou Paragraph.

- [Paragraph](paragraph.md): um contêiner para o conteúdo visível em uma página. Um parágrafo é um filho direto de uma estrutura de tópicos.

  - [ParagraphCollection](paragraphcollection.md): uma coleção de objetos Paragraph em uma estrutura de tópicos.

- [RichText](richtext.md): um objeto RichText.

- [Table](table.md): um contêiner para objetos TableRow.

- [TableRow](tablerow.md): um contêiner para objetos TableCell.

  - [TableRowCollection](tablerowcollection.md): um conjunto de objetos TableRow em uma Table.
 
- [TableCell](tablecell.md): um contêiner para objetos Paragraph.

  - [TableCellCollection](tablecellcollection.md): um conjunto de objetos TableCell em uma TableRow.
        
## Recursos adicionais

- [Visão geral da programação da API JavaScript do OneNote](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [Crie seu primeiro suplemento do OneNote](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma Suplementos do Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
