# Übersicht über die JavaScript-API-Programmierung für OneNote

OneNote führt eine JavaScript-API für OneNote Online-Add-Ins ein. Sie können Aufgabenbereich-Add-Ins, Inhalts-Add-Ins und Add-In-Befehle erstellen, die mit OneNote-Objekten interagieren und eine Verbindung zu Webdiensten oder anderen webbasierten Ressourcen herstellen.

Add-Ins bestehen aus zwei grundlegenden Komponenten:

- Einer **Webanwendung**, die aus einer Webseite und anderen erforderlichen JavaScript-, CSS- oder anderen Dateien besteht. Diese Dateien werden auf einem Webserver oder auf einem Webhostdienst gehostet, z. B. Microsoft Azure. In OneNote Online wird die Webanwendung in einem Browsersteuerelement oder Iframe angezeigt.
    
- Einem **XML-Manifest**, das die URL der Webseite des Add-Ins sowie Zugriffsanforderungen, Einstellungen und Funktionen für das Add-In angibt. Diese Datei wird auf dem Client gespeichert. OneNote-Add-Ins verwenden dasselbe [Manifest](https://dev.office.com/docs/add-ins/overview/add-in-manifests)format wie andere Office-Add-Ins.

**Office Add-In = Manifest + Webseite**

![Ein Office-Add-In besteht aus einem Manifest und einer Webseite](../../images/onenote-add-in.png)

### Verwenden der JavaScript-API

Add-Ins verwenden den Laufzeitkontext der Hostanwendung, um auf die JavaScript-API zuzugreifen. Die API besteht aus zwei Ebenen: 

- Einer **umfangreiche API** für OneNote-spezifische Vorgänge, auf die über das **Application**-Objekt zugegriffen wird.
- Einer **allgemeinen API**, die über Office-Anwendungen hinweg freigegeben ist und über die über das **Document**-Objekt zugegriffen wird.

#### Zugreifen auf die umfangreiche API über das *Application*-Objekt

Verwenden Sie das **Application**-Objekt, um auf OneNote-Objekte zuzugreifen, z. B. **Notebook**, **Section** und **Page**. Mit umfangreichen APIs können Sie auf Proxyobjekten Batchvorgänge ausführen. Der grundlegende Fluss sieht ungefähr folgendermaßen aus: 

1. Rufen Sie die Anwendungsinstanz aus dem Kontext auf.

2. Erstellen Sie einen Proxy, der das OneNote-Objekt darstellt, mit dem Sie arbeiten möchten. Sie interagieren synchron mit Proxyobjekten, indem Sie deren Eigenschaften lesen und ihre Methoden aufrufen. 

3. Rufen Sie **load** im Proxy auf, um diesen mit den im Parameter angegebenen Eigenschaften zu füllen. Dieser Aufruf wird der Warteschlange von Befehlen hinzugefügt. 

   Methodenaufrufe der API (z. B. `context.application.getActiveSection().pages;`) werden ebenfalls der Warteschlange hinzugefügt.
    
4. Rufen Sie **context.sync** auf, um alle in die Warteschlange eingereihten Befehle in der Reihenfolge auszuführen, in der sie in die Warteschlange gestellt wurden. Dadurch wird der Zustand zwischen dem ausgeführten Skript und den realen Objekten und durch Abrufen von Eigenschaften von geladenen OneNote-Objekten für die Verwendung in Ihrem Skript synchronisiert. Sie können das zurückgegebene promise-Objekt zum Verketten zusätzlicher Aktionen verwenden.

Beispiel: 

```
    function getPagesInSection() {
        OneNote.run(function (context) {
            
            // Get the pages in the current section.
            var pages = context.application.getActiveSection().pages;
            
            // Queue a command to load the id and title for each page.            
            pages.load('id,title');
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    
                    // Read the id and title of each page. 
                    $.each(pages.items, function(index, page) {
                        var pageId = page.id;
                        var pageTitle = page.title;
                        console.log(pageTitle + ': ' + pageId); 
                    });
                })
                .catch(function (error) {
                    app.showNotification("Error: " + error);
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
        });
    }
```

Die unterstützten OneNote-Objekte und -Vorgänge finden Sie in der [API-Referenz](../../reference/onenote/onenote-add-ins-javascript-reference.md).

### Zugreifen auf die allgemeine API über das *Document*-Objekt

Verwenden Sie das **Document**-Objekt, um auf die allgemeine API zuzugreifen, z. b. die Methoden ["getselecteddataasync"](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) und [SetSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync). 

Beispiel:  

```
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```
OneNote-Add-Ins unterstützen nur die folgenden allgemeinen APIs:

| API | Hinweise |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) | Nur **Office.CoercionType.Text** und **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142145.aspx) | Nur **Office.CoercionType.Text**, **Office.CoercionType.Image** und **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(name);](https://msdn.microsoft.com/en-us/library/office/fp142180.aspx) | Einstellungen werden nur von Inhalts-Add-Ins unterstützt. | 
| [Office.context.document.settings.set(name, value);](https://msdn.microsoft.com/en-us/library/office/fp161063.aspx) | Einstellungen werden nur von Inhalts-Add-Ins unterstützt. | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

Im Allgemeinen verwenden Sie die allgemeine API nur, um etwas zu tun, was von der umfangreichen API nicht unterstützt wird. Weitere Informationen zur Verwendung der allgemeinen API finden Sie in der [Dokumentation](https://dev.office.com/docs/add-ins/overview/office-add-ins) und [Referenz](https://dev.office.com/reference/add-ins/javascript-api-for-office) der Office-Add-Ins.


<a name="om-diagram"></a>
## OneNote-Objektmodelldiagramm 
Im folgenden Diagramm ist dargestellt, was derzeit in der JavaScript-API für OneNote verfügbar ist.

  ![OneNote-Objektmodelldiagramm](../../images/onenote-om.png)


## Zusätzliche Ressourcen

- [Erstellen Ihres ersten OneNote-Add-Ins](onenote-add-ins-getting-started.md)
- [JavaScript-API-Referenz für OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader-Beispiel](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office-Add-Ins-Plattformübersicht](https://dev.office.com/docs/add-ins/overview/office-add-ins)
