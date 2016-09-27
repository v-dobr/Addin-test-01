# Autorisieren von externen Diensten in Ihrem Office-Add-In

Mit beliebten Onlinediensten, wie Office 365, Google, Facebook, LinkedIn, SalesForce und GitHub, können Entwickler Benutzern Zugriff auf ihre Konten in anderen Anwendungen gewähren. Damit erhalten Sie die Möglichkeit, diese Dienste in Ihr Office-Add-In einzuschließen. 

Das Standardframework in der Branche zum Aktivieren des Webanwendungszugriffs auf einen Onlinedienst heißt OAuth 2.0. In den meisten Fällen müssen Sie nicht wissen, wie das Framework im Detail arbeitet, um es in Ihrem Add-In verwenden zu können. Es gibt viele Bibliotheken, in denen die Details für Sie abstrahiert werden.

Eine grundlegende Vorstellung von OAuth besteht darin, dass eine Anwendung ein Sicherheitsprinzipal für sich selbst ist, genau wie ein Benutzer oder eine Gruppe, mit eigener Identität und eigenem Berechtigungssatz. In den am häufigsten verwendeten Szenarien sendet das Add-In, wenn der Benutzer in dem Office-Add-In eine Aktion ausführt, für die der Onlinedienst erforderlich ist, dem Dienst eine Anforderung für einen bestimmten Satz von Berechtigungen im Benutzerkonto. Der Dienst fordert dann den Benutzer auf, dem Add-In die jeweiligen Berechtigungen zu gewähren. Nachdem die Berechtigungen erteilt wurden, sendet der Dienst dem Add-In ein kleines codiertes *Zugriffstoken*. Das Add-In kann den Dienst verwenden, indem das Token in alle Anforderungen an die Dienst-APIs eingeschlossen wird. Aber das Add-In kann nur innerhalb der Berechtigungen fungieren, die der Benutzer erteilt hat. Das Token läuft auch nach einer bestimmten Zeit ab.

Unterschiedliche OAuth-Muster, die als *Flüsse* oder *Erteilungstypen* bezeichnet werden, sind für unterschiedliche Szenarien vorgesehen. Nachfolgend finden Sie die beiden wichtigsten:

- **Impliziter Fluss**: Die Kommunikation zwischen dem Add-In und dem Onlinedienst wird mit dem clientseitigen JavaScript implementiert.
- **Autorisierungscodefluss**: Die Kommunikation erfolgt von *Server zu Server* zwischen der Webanwendung Ihres Add-Ins und dem Onlinedienst. Sie wird also mit serverseitigem Code implementiert.

Der Zweck der Flüsse besteht darin, die Identität und Autorisierung der Anwendung zu schützen. Im Autorisierungscodefluss erhalten Sie einen *geheimen Clientschlüssel*, der verborgen gehalten werden muss. Eine SPA (Single Page Application) hat keine Möglichkeit, den Schlüssel zu schützen, es wird daher empfohlen, dass Sie in SPAs den impliziten Fluss verwenden. 

Sie sollten mit den weiteren Vor- und Nachteilen der beiden Flüsse vertraut sein. Die offiziellen Definitionen unter [Autorisierungscode](https://tools.ietf.org/html/rfc6749#section-1.3.1) und [Implizit](https://tools.ietf.org/html/rfc6749#section-1.3.2) sind ein guter Ausgangspunkt. 

>**Hinweis:** Sie haben auch die Möglichkeit, einem Zwischendienst die gesamte Autorisierung für Sie zu überlassen und das Zugriffstoken an das Add-In zu übergeben. Weitere Informationen finden Sie im Abschnitt *Zwischendienste* weiter unten in diesem Artikel.

## Verwenden des impliziten Flusses in Office-Add-Ins
Die beste Möglichkeit herauszufinden, ob der Onlinedienst den impliziten Fluss unterstützt, besteht darin, in der Dokumentation nachzusehen.

Für Dienste, die diesen unterstützen, bieten wir eine JavaScript-Bibliothek, in der alle Details enthalten sind:

[Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

Der Ordner „\demo“ des Repositorys enthält ein Beispiel-Add-In, das die Bibliothek verwendet, um auf einige beliebte Dienste zuzugreifen, z. B. Google, Facebook und Office 365.

Siehe auch der Abschnitt **Bibliotheken** weiter unten in diesem Artikel.

## Verwenden des Autorisierungscodeflusses in Office-Add-Ins

Es gibt einige Beispiel-Add-Ins, die den Autorisierungscodefluss verwenden:

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

Es stehen viele Bibliotheken zur Implementierung des Autorisierungscodeflusses in verschiedenen Sprachen und Frameworks zur Verfügung. Details finden Sie im Abschnitt **Bibliotheken** weiter unten in diesem Artikel.

### Relay/Proxy-Funktionen

Sie können den Autorisierungscodefluss auch mit einer Webanwendung ohne Server verwenden, indem Sie die Werte *Client-ID* und *geheimer Clientschlüssel* in einer einfachen Funktion speichern, die in einem Dienst, z. B. [Azure-Funktionen](https://azure.microsoft.com/en-us/services/functions) oder [Amazon Lambda](https://aws.amazon.com/lambda), gehostet wird.
Die Funktion tauscht einen bestimmten Code für ein geeignetes *Zugriffstoken* aus und leitet es weiter an den Client. Die Sicherheit dieses Ansatzes ist davon abhängig, wie gut der Zugriff auf die Funktion geschützt ist.

Um diese Methode verwenden zu können, zeigt das Add-In eine Benutzeroberfläche/ein Popup an, um den Anmeldebildschirm für den Onlinedienst anzuzeigen (Google, Facebook usw.). Wenn der Benutzer angemeldet ist und dem Add-In die Berechtigung für seine Ressourcen im Onlinedienst gewährt, erhält der Entwickler einen Code, der dann an die Onlinefunktion gesendet werden kann. Die unter **Zwischendienste** in diesem Artikel beschriebenen Dienste verwenden einen ähnlichen Fluss. 

## Bibliotheken

Bibliotheken sind für viele Sprachen und Plattformen und für beide Flüsse verfügbar. Einige dienen einem allgemeinen Zweck, andere richten sich an spezifische Onlinedienste. 

**Office 365 und andere Dienste, die Azure Active Directory als Autorisierungsanbieter verwenden**: [Azure Active Directory-Authentifizierungsbibliotheken](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/). Eine Vorschau steht auch für die [Microsoft-Authentifizierungsbibliothek](https://www.nuget.org/packages/Microsoft.Identity.Client) zur Verfügung.

**Google**: Durchsuchen Sie [GitHub.com/Google](https://github.com/google) nach „Auth“ oder dem Namen Ihrer Sprache. Die meisten der relevanten Repositorys heißen `google-auth-library-[name of language]`.

**Facebook**: Durchsuchen Sie [Facebook für Entwickler](https://developers.facebook.com) nach "Bibliothek" oder "sdk". 

**OAuth 2.0 (allgemein)**: Eine Seite mit Links zu Bibliotheken für mehr als ein Dutzend Sprachen wird von der IETF-OAuth-Arbeitsgruppe an der folgenden Stelle verwaltet: [OAuth-Code](http://oauth.net/code/). Beachten Sie, dass einige dieser Bibliotheken der Implementierung eines OAuth-kompatiblen Diensts dienen. Die Bibliotheken, die für Sie als Add-In-Entwickler von Interesse sind, heißen auf dieser Seite *Client*bibliotheken, weil Ihr Webserver ein Client des OAuth-kompatiblen Diensts ist.

## Zwischendienste

Das Add-In kann einen Zwischendienst verwenden, z. B. Auth0, der entweder Zugriffstoken für viele beliebte Onlinedienste bereitstellt oder den Prozess des Aktivierens der sozialen Anmeldung für Ihr Add-In vereinfacht, oder beides. Mit sehr wenig Code kann Ihr Add-In entweder mithilfe von clientseitigem oder serverseitigem Code eine Verbindung mit dem Zwischendienst herstellen und alle erforderlichen Token für den Onlinedienst zurücksenden. Der gesamte Code der Autorisierungsimplementierung befindet sich im Zwischendienst. 

Es gibt ein Beispiel, das Auth0 verwendt, um die soziale Anmeldung bei Facebook, Google und Microsoft-Konten zu aktivieren:

[Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)

## Was ist CORS?

CORS steht für [Cross Origin Resource Sharing](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS). Informationen darüber, wie Sie mit CORS innerhalb von Add-Ins arbeiten können, finden Sie unter [Behandeln von Richtlinieneinschränkungen aufgrund desselben Ursprungs in Office-Add-Ins](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations).
