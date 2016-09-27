# JavaScript-API-Referenz zu OneNote-Add-Ins

*Gilt für: OneNote Online*

Unter den im Folgenden aufgeführten Links werden die in der API zur Verfügung stehenden OneNote-Objekte auf hoher Ebene Excel-Objekte erläutert. Auf jeder Seite sind eine Beschreibung der Eigenschaften, das Verhältnis und verfügbare Methoden des Objekts enthalten. Erkunden Sie diese Links, um mehr zu erfahren. 
    
- [Application](application.md): Das Objekt der obersten Ebene, das für den Zugriff auf alle global adressierbaren OneNote-Objekte verwendet wird, z. B. das aktive Notizbuch und den aktiven Abschnitt.

- [Notebook](notebook.md): Ein Notizbuch. Notizbücher können Abschnittsgruppen und Abschnitte enthalten.

   - [NotebookCollection](notebookcollection.md): Eine Auflistung von Notizbüchern.

- [SectionGroup](sectiongroup.md): Eine Abschnittsgruppe. Abschnittsgruppen enthalten Abschnittsgruppen und Abschnitte.

   - [SectionGroupCollection](sectiongroupcollection.md): Eine Auflistung von Abschnittsgruppen.

- [Section](section.md): Ein Abschnitt. Abschnitte enthalten Seiten.

   - [SectionCollection](sectioncollection.md): Eine Auflistung von Abschnitten.

- [Page](page.md): Eine Seite. Seiten enthalten PageContent-Objekte.

   - [PageCollection](pagecollection.md): Eine Auflistung von Seiten.

- [PageContent](pagecontent.md): Ein Bereich auf oberster Ebene auf einer Seite, der Inhaltstypen enthält, z. B. Outline oder Image. Ein PageContent-Objekt kann einer Position auf der Seite zugewiesen werden.

   - [PageContentCollection](pagecontentcollection.md): Eine Auflistung von PageContent-Objekten, die die Inhalte einer Seite darstellt.

- [Outline](outline.md): Ein Container für Paragraph-Objekte. Eine Gliederung ist ein direkt untergeordnetes Element eines PageContent-Objekts.

- [Image](image.md): Ein Bildobjekt. Ein Bild kann ein direkt untergeordnetes Element eines PageContent-Objekts oder eines Paragraph-Objekts sein.

- [Paragraph](paragraph.md): Ein Container für den sichtbaren Inhalt auf einer Seite. Ein Absatz ist ein direkt untergeordnetes Element einer Gliederung.

  - [ParagraphCollection](paragraphcollection.md): Eine Auflistung von Paragraph-Objekten in einer Gliederung.

- [RichText](richtext.md): Ein RichText-Objekt.

- [Table](table.md): Ein Container für TableRow-Objekte.

- [TableRow](tablerow.md): Ein Container für TableCell-Objekte.

  - [TableRowCollection](tablerowcollection.md): Eine Auflistung von TableRow-Objekten in einer Tabelle.
 
- [TableCell](tablecell.md): Ein Container für Paragraph-Objekte.

  - [TableCellCollection](tablecellcollection.md): Eine Auflistung von TableCell-Objekten in einer TableRow.
        
## Weitere Ressourcen

- [Übersicht über die JavaScript-API-Programmierung für OneNote](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [Erstellen Ihres ersten OneNote-Add-Ins](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Rubric Grader-Beispiel](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office-Add-Ins-Plattformübersicht](https://dev.office.com/docs/add-ins/overview/office-add-ins)
