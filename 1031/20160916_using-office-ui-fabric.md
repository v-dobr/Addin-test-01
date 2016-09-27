
#Verwenden der Office-UI-Fabric in Office-Add-Ins

Wenn Sie ein Office-Add-In erstellen, empfehlen wir Ihnen die [Office-UI-Fabric](https://github.com/OfficeDev/Office-UI-Fabric) zum Erstellen der Benutzeroberfläche. Die folgenden Schritte führen Sie schrittweise durch die Grundlagen für die Verwendung von Fabric.  

##1. Einrichten von Fabric
Fügen Sie dem Anfangsabschnitt Ihres HTML-Codes folgende Zeilen hinzu, um aus dem CDN auf die Fabric zu verweisen.

     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">


##2. Verwenden von Fabric-Symbolen und -Schriftarten
Das Verwenden von Symbolen ist einfach. Dazu müssen Sie nur das "i"-Element verwenden und auf die entsprechenden Klassen verweisen. Sie können die Größe des Symbols steuern, indem Sie den Schriftgrad ändern.

    <i class="ms-Icon ms-Icon--group" style="font-size:xx-large" aria-hidden="true"></i>


##3. Verwenden von Formatvorlagen für einfache Komponenten
Fabric verfügt über Formatvorlagen für verschiedene Benutzeroberflächenelemente, z. B. Schaltflächen und Kontrollkästchen. Zum Hinzufügen der entsprechenden Formatvorlage müssen Sie nur auf die entsprechenden Klassen verweisen, wie im folgenden Beispiel gezeigt.

    <button class="ms-Button" id="get-data-from-selection">
    <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
    <span class="ms-Button-label">Get Data from selection</span>
    <span class="ms-Button-description">Get Data from the document selection</span>
    </button>

##4. Verwenden von Komponenten mit Beispielverhalten
Fabric umfasst einige Komponenten, die Verhaltensweisen unterstützen (z. B. was beim Klicken passiert). Für den Einstieg umfasst Fabric 2.6.1 enthält einige **Codebeispiele** in Form von JQuery UI-Plug-Ins, die Sie verwenden können. Sie können auch ein beliebiges anderes Framework verwenden, um die Elemente zu verbinden. Wenn Sie diese Beispiele verwenden möchten, beachten Sie, dass der Code nicht als Teil des CDN verteilt wird, deshalb müssen Sie ihn aus der Version 2.6.1 des [Fabric-GitHub-Projekts](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1) herunterladen, darauf verweisen, und in Ihrem Code darauf verweisen. 

Gehen Sie beispielsweise folgendermaßen vor, um die SearchBox-Komponente zu verwenden:

1. Laden Sie die SearchBox-Komponente aus [GitHub](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1/src/components/SearchBox) herunter.
2. Fügen Sie Ihrem Code folgenden Verweis hinzu: `<script src="SearchBox/Jquery.SearchBox.js"></script>`
3. Initialisieren Sie die Komponente, indem Sie sicherstellen, dass diese Zeile beim Laden der Seite ausgeführt wird: `$(".ms-SearchBox").SearchBox();`. Es wird empfohlen, die Zeile im `Office.Initialize`-Block Ihres Add-Ins einzubinden.     

**Hinweis:** Wenn Sie nicht alle Fabric-Komponenten verwenden möchten, können Sie die Größe der Ressourcen verringern, die Sie herunterladen, indem Sie stattdessen die einzelnen CSS-Dateien für die einzelnen Komponenten hosten. Sie erhalten die CSS-Dateien von den Komponentenordnern im [Fabric 2.6.1-GitHub-Repository](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1). 


##Nächste Schritte
Wenn Sie End-to-End-Beispiele suchen, in denen die Verwendung von Fabric veranschaulicht wird, können wir Ihnen weiterhelfen. Informationen finden Sie im [Beispiel für die Office-Add-In-Fabric-UI](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample). Sie können auch die interaktive [Office-UI-Fabric](https://github.com/OfficeDev/Office-UI-Fabric)-Website besuchen.

