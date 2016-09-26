
# Debuggen von Office-Add-Ins auf dem iPad und einem Mac-Computer

Sie können Visual Studio für das Entwickeln und Debuggen von Add-Ins unter Windows verwenden. Sie können jedoch nicht zum Debuggen von Add-Ins auf dem iPad oder Mac verwenden. Da Add-Ins ausschließlich in HTML und JavaScript entwickelt werden, sollten sie plattformübergreifend funktionieren, aber möglicherweise gibt es subtile Unterschiede darin, wie verschiedene Browser den HTML-Code rendern. In diesem Thema wird beschrieben, wie Sie Add-Ins debuggen, die auf dem iPad oder Mac ausgeführt werden. 

## Debuggen mit Vorlon.js 

Vorlon.js ist ein Debugger für Webseiten, ähnlich wie die F12-Tools, der auf das Remotearbeiten ausgelegt ist und Ihnen das Debuggen von Webseiten über verschiedene Geräte ermöglicht. Weitere Informationen finden Sie auf der [Vorlon.js-Website](http://www.vorlonjs.com).  

So installieren Sie Vorlon und richten es ein: 

1.  Installieren Sie [Node.js](https://nodejs.org) und [Git](https://git-scm.com/), falls noch nicht erfolgt. 

2.  Installieren Sie Vorlon mit git mit dem folgenden Befehl: `git clone https://github.com/MicrosoftDX/Vorlonjs.git`

3.  Installieren von Abhängigkeiten mit `npm install`.

4.  Add-Ins erfordern HTTPS, alle Skripts, die diese verwenden, müssen daher ebenfalls HTTPS sein, einschließlich des Vorlon-Skripts. Daher müssen Sie Vorlon so konfigurieren, dass es SSL verwendet, damit Vorlon mit Add-Ins verwendet werden kann. Gehen Sie unter dem Ordner, in dem Sie Vorlon installiert haben, zu dem Ordner „/Server“, und bearbeiten Sie die Datei „config.json“. Ändern Sie die **useSSL**-Eigenschaft in **true**. Sie können auch gleich das Plug-In für Office-Add-Ins aktivieren (ändern Sie die enabled-Eigenschaft in „true“). 

5.  Führen Sie den Vorlon-Server mit dem Befehl `sudo vorlon` aus. 

6.  Öffnen Sie ein Browserfenster, und wechseln Sie zu [http://localhost:1337](http://localhost:1337), der Vorlon-Schnittstelle. Vertrauen Sie dem Sicherheitszertifikat (Sie werden dazu aufgefordert). Sie können das Sicherheitszertifikat auch im Ordner „Vorlon“ unter „/Server/cert“ finden. 

7.  Fügen Sie das folgende Skripttag zum Abschnitt `<head>` der Datei „home.html“ (oder zur Haupt-HTML-Datei) Ihres Add-Ins hinzu:
```    
<script src="https://localhost:1337/vorlon.js"></script>    
```  

Sobald Sie das Add-In nun auf einem Gerät öffnen, wird es in der Liste der Clients in Vorlon (auf der linken Seite der Vorlon-Benutzeroberfläche) angezeigt. Sie können DOM-Elemente remote hervorheben, remote Befehle ausführen und vieles mehr.  

![Screenshot, der die Vorlon.js-Benutzeroberfläche anzeigt](../../images/vorlon_interface.png)

Das Office-Plug-In fügt zusätzliche Funktionen für Office.js hinzu, z. B. Erforschen des Objektmodells und Ausführen von Office.js-Aufrufen. 
