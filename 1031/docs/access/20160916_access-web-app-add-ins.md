
# Erstellen von Add-Ins für Access-Web-Apps



In diesem Artikel erfahren Sie, wie Sie Visual Studio 2015 zum Entwickeln eines Office-Add-Ins verwenden, das auf Access Web Apps abzielt.

>
  **Hinweis:** Informationen zur Entwicklung von Lösungen für Access unter Verwendung von VBA finden Sie unter [Access](https://msdn.microsoft.com/en-us/library/fp179695.aspx) auf MSDN.

## Voraussetzungen

Zum Erstellen eines Office-Add-In für Access Web Apps benötigen Sie Folgendes:


- Visual Studio 2015

- Eine SharePoint Online-Website (in vielen Office 365-Abonnements enthalten). Diese Website muss über einen Add-In-Katalog verfügen. Weitere Informationen finden Sie unter [Einrichten eines Add-In-Katalogs unter SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


 >**Hinweis**  Office-Add-Ins funktionieren mit Access Web Apps, die auf SharePoint Online oder Office 365 gehostet sind. Die Access 2013-Desktopanwendung unterstützt keine Office-Add-Ins. Office-Add-Ins, die auf Access Web Apps abzielen, werden von Version 1.1 und höher von Office.js unterstützt.


## Erstellen eines Projekts in Visual Studio


1.  Öffnen Sie Visual Studio, und wählen Sie im Menü nacheinander **Datei**,  **Neu**,  **Projekt**. Das Dialogfeld  **Neues Projekt** wird angezeigt.

2. Navigieren Sie im Dialogfeld  **Neues Projekt** im linken Fensterbereich zu **Installiert**,  **Vorlagen**,  **Visual C#**,  **Office/SharePoint**,  **Office-Add-Ins**.

3. Wählen Sie im Dialogfeld  **Neues Projekt** im mittleren Bereich **Office-Add-In** aus.

4. Geben Sie unten im Dialogfeld einen Namen für das Projekt ein, und wählen Sie  **OK**. Dadurch wird das Dialogfeld  **Office-Add-in erstellen** geöffnet.

5. Wählen Sie im Dialogfeld  **Office-Add-In erstellen** zuerst **Inhalt** und dann **Weiter** aus.

6. Wählen Sie auf der nächsten Seite des Dialogfelds  **Office-Add-in erstellen** entweder **Grundlegendes Add-in** oder **Dokumentvisualisierungs-Add-in** aus, und stellen Sie sicher, dass das Kontrollkästchen **Access** aktiviert ist.

7. Abschließend wählen Sie  **Fertig stellen**. Visual Studio erstellt ein Startprojekt für Sie, auf dem Sie aufbauen können.

8. Wählen Sie im  **Projektmappen-Explorer** das Webprojekt aus ( **project_name>Web**). Suchen Sie im Eigenschaftenbereich den Eintrag für  **SSL-URL**. Dieser sollte etwa so aussehen:  `https://localhost:44314/`. Markieren Sie diese URL, und kopieren Sie sie in die Zwischenablage. Sie benötigen sie in Kürze.

9. Klicken Sie im  **Projektmappen-Explorer** mit der rechten Maustaste auf den Namen Ihres Projekts. Wählen Sie im Kontextmenü den Befehl  **Veröffentlichen**. Dadurch wird der Assistent  **Add-in veröffentlichen** geöffnet.

10. Wählen Sie im Assistenten  **Add-in veröffentlichen** die Dropdownliste neben **Aktuelles Profil** aus. Wählen Sie in dieser Dropdownliste  **neu** aus. Dadurch wird das Dialogfeld  **Office- und SharePoint-Add-ins veröffentlichen** geöffnet.

11. Wählen Sie in diesem Dialogfeld  **Neues Profil erstellen**, geben Sie einen wiedererkennbaren Namen für das Profil ein, und wählen Sie dann  **Fertig stellen**. Das Dialogfeld  **Office- und SharePoint-Add-ins veröffentlichen** wird geschlossen, und Sie kehren zum Assistenten **Add-in veröffentlichen** zurück.

12. Wählen Sie im Assistenten die Option  **Add-in verpacken** aus. Dadurch wird das Add-in abgeschlossen, sodass es in einem Add-in-Katalog in SharePoint veröffentlicht werden kann.

13. Geben Sie auf der nächsten Seite unter **Wo wird Ihre Website gehostet?** die URL für den Host Ihrer Website ein. Dies kann der **SSL-URL**-Wert sein, den Sie im Schritt 8 kopiert haben. Wählen Sie dann  **Fertig stellen**.

14. Klicken Sie im  **Projektmappen-Explorer** unter dem Projektnamen mit der rechten Maustaste auf den Manifestknoten des Projekts und anschließend auf **Ordner im Projektmappen-Explorer öffnen**. Notieren Sie sich den Pfad dieser Datei. Sie benötigen diesen Wert später.


 >**Hinweis**  Sie können das Add-In nicht debuggen, ohne es mit einer Access Web App bereitzustellen.


## Überprüfen des Manifests und der Datei "Home.html"


1. Öffnen Sie in Ihrem Visual Studio-Projekt die Datei  **Home.html**, und suchen Sie die Zeilen, die auf die Skriptbibliothek "office.js" verweisen.

```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```
 >**Beachten Sie**, dass das Skript-Tag auf die Version 1.1 (und höher) von „Office.js“ verweist. Access verwendet API-Elemente, die in der Version 1.1 eingeführt wurden.

2. Öffnen Sie die Ihrem Projekt zugeordnete Manifestdatei. Die Datei trägt den Namen des Projekts mit der Erweiterung ".xml".

3.  Suchen Sie in der Manifestdatei den Abschnitt **Hosts** und darin einen Eintrag **Host**.

```xml
  <Hosts> <Host Name="Database" /> </Hosts>
```
 >**Hinweis** Dort sind die Anwendungen aufgeführt, die das Add-In verwenden können. Da Sie im **Dialogfeld Office-Add-in** erstellen die Option  **Access** ausgewählt haben, ist **Database** aufgeführt. Wenn Sie Excel eingeschlossen haben, ist auch ein Eintrag für **Workbook** vorhanden.

Office- und SharePoint-Add-Ins sind webbasiert. Der Code für das Add-In muss auf einem Webserver gehostet werden. Im vorliegenden Beispiel dient Ihr Entwicklungscomputer als Webserver. Der Server muss in Betrieb sein, um das Add-I)n zum Testen zur Verfügung zu stellen. In diesem Fall heißt das, dass das Add-In in Visual Studio ausgeführt werden muss, wenn Sie das Add-In in SharePoint überprüfen und debuggen möchten.

Damit Benutzer das Add-In finden und verwenden können, muss das Add-In in einem Add-in-Katalog in SharePoint registriert sein. Die Informationen, die für den Add-in-Katalog benötigt werden, sind in der Manifestdatei enthalten.

 >**Hinweis**  Sie müssen eine Access Web App erstellen, die Ihre Office-Add-In hostet.


## Veröffentlichen des Add-Ins in einem SharePoint Online-Katalog


1.  Melden Sie sich in SharePoint Online oder Office 365 an, und navigieren Sie dann zum **SharePoint Admin Center**, indem Sie in der Office 365-Symbolleiste oben auf der Seite  **Admin** auswählen.

2. Wählen Sie auf der Linkleiste links auf der Seite  **SharePoint Admin Center** die Option **Add-ins** aus. Dadurch gelangen Sie zur Add-in-Ansicht.

3. Wählen Sie im mittleren Bereich der Seite  **Add-in-Katalog** aus. Dadurch gelangen Sie auf die Seite **Katalog**.

4. Wählen Sie auf der Seite  **Katalog** die Option **Office-Add-ins** aus. Dadurch gelangen Sie auf eine Verzeichnisseite namens **Office-Add-ins**. Dort sind alle installierten Office-Add-Ins aufgeführt.

5. Wählen Sie oben auf der Seite  **Office-Add-ins** die Option **Neues Add-in** aus. Daraufhin wird das Dialogfeld **Dokument hinzufügen** angezeigt.

6. Wählen Sie im Dialogfeld  **Dokument hinzufügen** die Option **Durchsuchen** aus, und navigieren Sie dann zum Speicherort der Manifestdatei in Ihrem Visual Studio-Projekt. Wenn Sie die Adresse der Manifestdatei zuvor kopiert haben, können Sie sie nun im Dialogfeld einfügen.

7. Wählen Sie die Manifestdatei im Projekt und dann  **OK** aus. Das Add-In wird nun von SharePoint der lokalen SharePoint-Bibliothek hinzugefügt.


 >**Hinweis**  Bei diesem Verfahren wird vorausgesetzt, dass Sie bereits eine Testwebsite in SharePoint erstellt haben. Andernfalls können Sie dies auf der Registerkarte  **Websites** oben im SharePoint-Fenster nachholen. Sie können eine vorhandene Access Web Apps verwenden, wenn Sie über eine verfügen.


## Erstellen einer Access Web App zum Hosten des Add-Ins


1. Navigieren Sie zur Testwebsite. Wählen Sie in der Linkleiste auf der linken Seite  **Websiteinhalte** aus. Dadurch gelangen Sie auf die Seite **Websiteinhalte** der Testwebsite.

2. Wählen Sie auf der Seite  **Websiteinhalte** die Option **Add-in hinzufügen** aus. Dadurch gelangen Sie auf die Seite **Websiteinhalte – Ihre Add-ins**.

3. Suchen Sie auf der Seite  **Websiteinhalte – Ihre Add-ins** mithilfe der Suchleiste oben auf der Seite nach **Access-App**.

4. Es sollte jetzt eine Kachel für  **Access App** angezeigt werden.

     >**Hinweis** Bedenken Sie, dass dies nicht Ihr Office-Add-In, sondern eine neue Access-Web-App ist. Diese Access-Web-App hostet Ihr Office-Add-In.
5. Durch das Auswählen wird das Dialogfeld  **Access App hinzufügen** angezeigt. Geben Sie einen eindeutigen Namen für Ihre Access-App ein, und wählen Sie **Erstellen** aus. Es dauert möglicherweise eine Weile, bis SharePoint Ihre App erstellt hat. Wenn der Vorgang abgeschlossen ist, wird Ihre Access-App auf der Seite **Websiteinhaltes** mit der Beschriftung **Neu** angezeigt.

6. Sie müssen die AccessApp nun in der Desktopversion von Microsoft Access 2013 öffnen und Daten hinzufügen. Erst dann kann sie in SharePoint geöffnet und angezeigt werden.


## Hinzufügen des Add-Ins zu einer Access Web Apps


1. Öffnen Sie eine Access Web Apps.

2. Wählen Sie auf der SharePoint-Registerkartenleiste in der oberen linken Ecke das Zahnradsymbol aus. Ein Menü wird angezeigt. Wählen Sie das Menüelement  **Office-Add-ins** aus. Dadurch wird das Dialogfeld **Office-Add-ins** geöffnet.

3. Wählen Sie die Ansicht  **Meine Organisation**, und warten Sie einen Moment, bis das Dialogfeld von SharePoint mit den für Sie verfügbaren Office-Add-Ins gefüllt wurde.

    Eines der Add-Ins im Dialogfeld sollte das Office-Add-In sein, das Sie in einem vorherigen Verfahren registriert haben. Wählen Sie dieses Add-In aus, um es in Ihre Access-Web-Apps einzufügen. Bedenken Sie, dass die App in Visual Studio ausgeführt werden muss, damit Sie auf der Seite der Access-Web-Apps erkannt und angezeigt wird.


## Debuggen des Add-Ins für Office

Drücken Sie zum Debuggen Ihres Add-Ins in Internet Explorer F12, oder wählen Sie das Zahnradsymbol in der Registerkartenleiste des Browsers (nicht das Zahnradsymbol auf der SharePoint-Seite). Dadurch werden die F12-Debuggingtools von Internet Explorer 11 angezeigt. Wenn Sie einen anderen Browser verwenden, lesen Sie in der Browserdokumentation nach, wie der Debuggingmodus gestartet wird.

An dieser Stelle können Sie Haltepunkte setzen, Ihren JavaScript-Code schrittweise durchlaufen, das DOM erkunden und den Code ändern, um zu überprüfen, ob die Änderungen in der Office-Add-In für Access Web Apps sichtbar werden. Weitere Informationen finden Sie unter [Using the F12 developer tools](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29).


## Nächste Schritte

Laden Sie das Beispiel [Office 365: Binden und Bearbeiten von Daten in einer Access-Webanwendung](https://code.msdn.microsoft.com/officeapps/Office-365-Bind-and-4876274e) herunter, um weitere Informationen zum Implementieren eines Office-Add-In zu erhalten, das Daten in einer Access Web App bearbeitet.


## Weitere Ressourcen



- [Grundlegendes zur JavaScript-API für Add-ins](../develop/understanding-the-javascript-api-for-office.md)

- [JavaScript-API für Office](../../reference/javascript-api-for-office.md)

