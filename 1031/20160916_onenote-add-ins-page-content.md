# Arbeiten mit OneNote-Seiteninhalt 

In der JavaScript-API von OneNote-Add-Ins wird Seiteninhalt durch das folgende Objektmodell dargestellt.

  ![OneNote-Objektmodelldiagramm](../../images/OneNoteOM-page.png)

- Eine Page-Objekt enthält eine Auflistung von PageContent-Objekten.
- Ein PageContent-Objekt enthält den Inhaltstyp „Outline“, „Image“ oder „Other“.
- Ein Outline-Objekt enthält eine Auflistung von Paragraph-Objekten.
- Ein Paragraph-Objekt enthält den Inhaltstyp „RichtText“, „Image“, „Table“ oder „Other“.

Verwenden Sie eine der folgenden Methoden, um eine leere OneNote-Seite zu erstellen:

- [Section.addPage](../../reference/onenote/section.md#addpagetitle-string)
- [Page.insertPageAsSibling](../../reference/onenote/page.md#insertpageassiblinglocation-string-title-string)

Verwenden Sie dann folgende Methoden in den folgenden Objekten, um mit dem Seiteninhalt zu arbeiten, z. B. Page.addOutline und Outline.appendHtml 

- [Seite](../../reference/onenote/page.md)
- [Gliederung](../../reference/onenote/outline.md)
- [Absatz](../../reference/onenote/paragraph.md)

Inhalt und Struktur einer OneNote-Seite werden durch HTML-Code dargestellt. Für das Erstellen oder Aktualisieren von Seiteninhalt wird nur eine Teilmenge des HTML-Codes unterstützt, wie im Folgenden beschrieben.

## Unterstützter HTML-Code

Die JavaScript-API des OneNote-Add-Ins unterstützt den folgenden HTML-Code für das Erstellen und Aktualisieren von Seiteninhalten:

- `<html>`, `<body>`, `<div>`, `<span>`, `<br/>` 
- `<p>`
- `<img>`
- `<a>`
- `<ul>`, `<ol>`, `<li>` 
- `<table>`, `<tr>`, `<td>`
- `<h1>` ... `<h6>`
- `<b>`, `<em>`, `<strong>`, `<i>`, `<u>`, `<del>`, `<sup>`, `<sub>`, `<cite>`

## Zugriff auf Seiteninhalte

Sie können nur über `Page#load` auf den *Seiteninhalt* für die derzeit aktive Seite zugreifen. Rufen Sie zum Ändern der aktiven Seite `navigateToPage($page)` auf.

Metadaten, z. B. Titel, können weiterhin für eine beliebige Seite abgefragt werden.

## Weitere Ressourcen

- [Übersicht über die JavaScript-API-Programmierung für OneNote](onenote-add-ins-programming-overview.md)
- [JavaScript-API-Referenz für OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader-Beispiel](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office-Add-Ins-Plattformübersicht](https://dev.office.com/docs/add-ins/overview/office-add-ins)
