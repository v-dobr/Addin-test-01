# Référence de l’API JavaScript des compléments OneNote

*S’applique à : OneNote Online*

Les liens ci-dessous renvoient aux objets OneNote de niveau supérieur disponibles dans l’API. Chaque lien de page d’objet contient une description des propriétés, des relations et des méthodes disponibles sur l’objet. Explorez les liens ci-dessous pour en savoir plus. 
    
- [Application](application.md) : Objet de niveau supérieur utilisé pour accéder à tous les objets OneNote globalement adressables, tels que le bloc-notes actif et la section active.

- [Bloc-notes](notebook.md) : Bloc-notes. Les blocs-notes contiennent des groupes de sections et des sections.

   - [NotebookCollection](notebookcollection.md) : Collection de blocs-notes.

- [SectionGroup](sectiongroup.md) : Groupe de sections. Les groupes de sections contiennent des sections et des groupes de sections.

   - [SectionGroupCollection](sectiongroupcollection.md) : Collection de groupes de sections.

- [Section](section.md) : Section. Les sections contiennent des pages.

   - [SectionCollection](sectioncollection.md) : Collection de sections.

- [Page](page.md) : Page. Les pages contiennent des objets PageContent.

   - [PageCollection](pagecollection.md) : Collection de pages.

- [PageContent](pagecontent.md) : Zone de niveau supérieur sur une page qui contient des types de contenu tels que des plans ou des images. Un objet PageContent peut être affecté à une position sur la page.

   - [PageContentCollection](pagecontentcollection.md) : Collection d’objets PageContent qui représente le contenu d’une page.

- [Outline](outline.md) : Conteneur pour les objets Paragraph. Un plan est un enfant direct d’un objet PageContent.

- [Image](image.md) : Objet Image. Une image peut être un enfant direct d’un objet Paragraph ou PageContent.

- [Paragraph](paragraph.md) : Conteneur pour le contenu visible d’une page. Un paragraphe est un enfant direct d’un plan.

  - [ParagraphCollection](paragraphcollection.md) : Collection d’objets Paragraph dans un plan.

- [Richtext](richtext.md) : Objet RichText.

- [Table](table.md) : Conteneur pour les objets TableRow.

- [TableRow](tablerow.md) : Conteneur pour les objets TableCell.

  - [TableRowCollection](tablerowcollection.md) : Collection d’objets TableRow dans un tableau.
 
- [TableCell](tablecell.md) : Conteneur pour les objets Paragraph.

  - [TableCellCollection](tablecellcollection.md) : Collection d’objets TableCell dans un élément TableRow.
        
## Ressources supplémentaires

- [Vue d’ensemble de la programmation de l’API JavaScript de OneNote](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [Créer votre premier complément OneNote](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Exemple de grille d’évaluation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d’ensemble de la plateforme des compléments pour Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
