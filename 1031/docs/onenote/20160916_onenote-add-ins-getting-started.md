# Erstellen Ihres ersten OneNote-Add-Ins

Dieser Artikel führt Sie durch das Erstellen eines einfachen Aufgabenbereich-Add-Ins, mit dem einer OneNote-Seite Text hinzugefügt wird.

In der folgenden Abbildung ist das Add-In dargestellt, das Sie erstellen.

   ![Das OneNote-Add-In, das anhand dieser exemplarischen Vorgehensweise erstellt wurde](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## Schritt 1: Einrichten der Entwicklungsumgebung
1. Installieren Sie den Yeoman Office-Generator und dessen Komponenten, indem Sie die folgenden [Installationsanweisungen](https://dev.office.com/docs/add-ins/get-started/create-an-office-add-in-using-any-editor) befolgen.

   Der Yeoman Office-Generator erleichtert das Erstellen von Add-In-Projekten, wenn Sie nicht über Visual Studio verfügen oder andere Technologien als reines HTML, CSS und JavaScript verwenden möchten. Er bietet auch schnellen Zugriff auf einen lokalen Gulp-Webserver für Testzwecke. 

   >Sie können optional [Visual Studio verwenden](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio), um die Projektdateien zu erstellen, dann erhalten Sie aber nicht den integrierten Gulp-Serversupport.

<a name="create-project"></a>
## Schritt 2: Erstellen des Add-In-Projekts 
1. Erstellen Sie einen lokalen Ordner mit dem Namen *onenote add-in*.

2 – Öffnen Sie eine **Cmd**-Eingabeaufforderung, und navigieren Sie zum Ordner  **onenote add-in**. Führen Sie den `yo office`-Befehl wie unten dargestellt aus.

```
C:\your-local-path\onenote add-in\> yo office
```
>Diese Anweisungen verwenden die Windows-Eingabeaufforderung, gelten jedoch auch für andere Shell-Umgebungen. 

3 – Verwenden Sie zum Erstellen des Projekts die folgenden Optionen.

| Option | Wert |
|:------|:------|
| Projektname | OneNote-Add-In |
| Stammordner des Projekts | (Übernehmen Sie den Standardwert) |
| Office-Projekttyp | Aufgabenbereich-Add-In |
| Unterstützte Office-Anwendungen | (Treffen Sie eine Auswahl – wir werden später einen OneNote-Host hinzufügen) |
| Zu verwendende Technologie | HTML, CSS und JavaScript |

<a name="manifest"></a>
## Schritt 3: Konfigurieren des Add-In-Manifests 
1 – Öffnen Sie **manifest-onenote-add-in.xml** in Ihren Projektdateien. Fügen Sie dem Abschnitt **Hosts** die folgende Zeile hinzu. Dadurch wird angegeben, dass Ihr Add-In die OneNote-Hostanwendung unterstützt.

```
<Host Name="Notebook" />
```

Beachten Sie, dass die **SourceLocation** bereits für den Gulp-Webserver eingerichtet ist.

```
<SourceLocation DefaultValue="https://localhost:8443/app/home/home.html"/>
```

<a name="develop"></a>
## Schritt 4: Entwickeln des Add-Ins
Sie können das Add-in mitt einem beliebigen Text-Editor oder IDE entwickeln. Wenn Sie Visual Studio Code noch nicht ausprobiert haben, können Sie es [hier kostenlos herunterladen](https://code.visualstudio.com/) (für Linux, Windows und Mac OSX).

1 – Öffnen Sie die Datei **home.html** im Ordner *app/home*. 

2 – Bearbeiten Sie die Verweise auf die Formatvorlagen und Komponenten der Office-JavaScript-API und der [Office-UI-Fabric](http://dev.office.com/fabric).

   a. Kommentieren Sie den Link auf „fabric.components.min.css“ aus.

   b. Ersetzen Sie den Skriptverweis auf „Office.js“ durch den folgenden Verweis auf die *Beta*version.

```
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

Die Office-Verweise sehen wie folgt aus.

```
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.min.css" rel="stylesheet">
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css" rel="stylesheet">
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

3. Ersetzen Sie das `<body>`-Element durch den folgenden Code. Dadurch werden ein Textbereich und eine Schaltfläche mithilfe von [Office-UI-Fabric-Komponenten](http://dev.office.com/fabric/components) hinzugefügt. Das **dynamische Rasterlayout** stammt aus dem Satz von [Office-UI-Fabric-Formatvorlagen](http://dev.office.com/fabric/styles). 

```
<body class="ms-font-m">
    <div class="home flex-container">
        <div class="ms-Grid">
            <div class="ms-Grid-row ms-bgColor-themeDarker">
                <div class="ms-Grid-col">
                    <span class="ms-font-xl ms-fontColor-themeLighter ms-fontWeight-semibold">OneNote Add-in</span>
                </div>
            </div>
        </div>
        <br />
        <div class="ms-Grid">
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <label class="ms-Label">Enter content here</label>
                    <div class="ms-TextField ms-TextField--placeholder">
                        <textarea id="textBox" rows="5"></textarea>
                    </div>
                </div>
            </div>
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <div class="ms-font-m ms-fontColor-themeLight header--text">
                        <button class="ms-Button ms-Button--primary" id="addOutline">
                            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
                            <span class="ms-Button-label">Add outline</span>
                            <span class="ms-Button-description">Adds the content above to the current page.</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
```

4 – Öffnen Sie die Datei **home.js** im Ordner *app/home*. Bearbeiten Sie die Funktion **Office.initialize**, um ein Klickereignis zur Schaltfläche **Gliederung hinzufügen** wie folgt hinzuzufügen. 

```
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
    $(document).ready(function () {
        app.initialize();

        // Set up event handler for the UI.
        $('#addOutline').click(addOutlineToPage);
    });
};
```
 
5 – Ersetzen Sie die **getDataFromSelection**-Methode durch die folgende **addOutlineToPage**-Methode. Dadurch wird der Inhalt aus dem Textbereich abgerufen und zu der Seite hinzugefügt.

```
// Add the contents of the text area to the page.
function addOutlineToPage() {        
    OneNote.run(function (context) {
       var html = '<p>' + $('#textBox').val() + '</p>';

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to load the page with the title property.             
        page.load('title'); 

        // Add an outline with the specified HTML to the page.
        var outline = page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function() {
                console.log('Added outline to page ' + page.title);
            })
            .catch(function(error) {
                app.showNotification("Error: " + error); 
                console.log("Error: " + error); 
                if (error instanceof OfficeExtension.Error) { 
                    console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                } 
            }); 
        });
}
```

<a name="test"></a>
## Schritt 5: Testen des Add-Ins in OneNote Online
1 – Führen Sie den Gulp-Webserver aus.  

   a. Öffnen Sie eine **md**-Eingabeaufforderung, im Ordner **onenote add-in**. 

   b. Führen Sie den `gulp serve-static`-Befehl wie unten dargestellt aus.

```
C:\your-local-path\onenote add-in\> gulp serve-static
```

2 - Installieren Sie das selbstsignierte Zertifikat des Gulp-Webservers als vertrauenswürdiges Zertifikat. Sie müssen dies nur ein Mal auf Ihrem Computer für mit dem Yeoman Office-Generator erstellte Add-In-Projekte ausführen.

   a. Navigieren Sie zu der gehosteten Add-In-Seite. Standardmäßig ist dies die gleiche URL wie im Manifest:

```
https://localhost:8443/app/home/home.html
```

   b. Installieren Sie das Zertifikat als vertrauenswürdiges Zertifikat. Weitere Informationen finden Sie unter [Hinzufügen von selbstsignierten Zertifikaten als vertrauenswürdiges Stammzertifikat](https://github.com/OfficeDev/generator-office/blob/master/docs/trust-self-signed-cert.md).

3. Öffnen Sie in OneNote Online ein Notizbuch.

4. Wählen Sie **Einfügen > Office-Add-Ins**. Daraufhin wird das Dialogfeld „Office-Add-ins“ geöffnet.
  - Wenn Sie mit Ihrem Consumer-Konto angemeldet sind, wählen Sie die Registerkarte **Meine Add-Ins** aus, und wählen Sie dann **Mein Add-In hochladen** aus.
  - Wenn Sie mit Ihrem Geschäfts- oder Schulkonto angemeldet sind, wählen Sie die Registerkarte **Meine Organisation** aus, und wählen Sie dann **Mein Add-In hochladen** aus. 
  
  In der folgenden Abbildung ist die Registerkarte **Meine Add-Ins** für Consumer-Notizbücher dargestellt.

  ![Das Dialogfeld „Office-Add-Ins“ mit der Registerkarte „Meine Add-Ins“](../../images/onenote-office-add-ins-dialog.png)
  
  >**Hinweis**: Um die Schaltfläche **Office-Add-Ins** zu aktivieren, klicken Sie auf eine OneNote-Seite.

5. Navigieren Sie im Dialogfeld „Mein Add-In hochladen“ zu **manifest-onenote-add-in.xml** in den Projektdateien, und wählen Sie **Hochladen**. Beim Testen kann die Manifestdatei lokal gespeichert werden.

6. Das Add-In wird in einem iFrame neben der OneNote-Seite geöffnet. Geben Sie Text in den Textbereich ein, und wählen Sie dann **Gliederung hinzufügen** aus. Der von Ihnen eingegebene Text wird der Seite hinzugefügt. 

## Problembehandlung und Tipps
- Sie können das Add-In mithilfe der Entwicklertools Ihres Browsers debuggen. Wenn Sie den Gulp-Webserver verwenden und in Internet Explorer oder Chrome debuggen, können Sie Ihre Änderungen lokal speichern und dann einfach das iFrame des Add-Ins aktualisieren.

- Wenn Sie ein OneNote-Objekt prüfen, verwenden die Eigenschaften, die derzeit für die Verwendung verfügbar sind, tatsächliche Werte. Für Eigenschaften, die geladen werden müssen, wird *nicht definiert* angezeigt. Erweitern Sie den `_proto_`-Knoten, um die Eigenschaften anzuzeigen, die im Objekt definiert sind, aber noch nicht geladen wurden.

      ![Unloaded OneNote object in the debugger](../../images/onenote-debug.png)

- Wenn Ihr Add-In HTTP-Ressourcen verwendet, müssen Sie im Browser gemischten Inhalt aktivieren. Produktions-Add-Ins sollten nur sichere HTTPS-Ressourcen verwenden.

## Zusätzliche Ressourcen

- [Übersicht über die JavaScript-API-Programmierung für OneNote](onenote-add-ins-programming-overview.md)
- [JavaScript-API-Referenz für OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader-Beispiel](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office-Add-Ins-Plattformübersicht](https://dev.office.com/docs/add-ins/overview/office-add-ins)
