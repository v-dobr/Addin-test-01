
# Tipps für den Umgang mit Datumswerten in Outlook-Add-Ins

Die JavaScript-API für Office verwendet das JavaScript-Objekt [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp) für die meisten Speicher- und Abrufvorgänge von Datums- und Uhrzeitangaben. Das Objekt **Date** bietet Methoden wie [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp) und [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp), die den angeforderten Datums- oder Zeitwert in Universal Coordinated Time (UTC) zurückgeben.<br/><br/>
Das Objekt **Date** bietet auch andere Methoden wie z. B. [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp) und [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp), die das angeforderte Datum bzw. die Uhrzeit in Ortszeit („local time“) zurückgeben.<br/><br/>Das Konzept der Ortszeit wird größtenteils von Browser und Betriebssystem des Clientcomputers bestimmt. So gibt beispielsweise in den meisten Browsern auf Windows-basierten Clientcomputern ein JavaScript-Aufruf von **getDate** ein Datum zurück, das auf der in Windows auf dem Clientcomputer festgelegten Zeitzone basiert.<br/><br/>
Im folgenden Beispiel wird ein **Date**-Objekt <code>myLocalDate</code> in Ortszeit erstellt und **toUTCString** aufgerufen, um das Datum in eine Datumszeichenfolge in UTC umzuwandeln.




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Sie können zwar das JavaScript-Objekt  **Date** verwenden, um einen Datums- oder Zeitwert auf der Grundlage von UTC oder der Zeitzone des Clientcomputers abzurufen, das **Date**-Objekt besitzt jedoch in einer Hinsicht eine Einschränkung: Es stellt keine Methoden zur Rückgabe von Datums- oder Zeitwerten für andere spezifische Zeitzonen bereit. Wenn für Ihren Clientcomputer beispielsweise Eastern Standard Time (EST) eingestellt ist, steht keine  **Date**-Methode zur Verfügung, den Stundenwert einer anderen Zeitzone als EST oder UTC, beispielsweise Pacific Standard Time (PST), abzurufen.


## Datumsspezifische Funktionen für Outlook-Add-In


Die vorgenannten JavaScript-Einschränkungen sind für Sie von Bedeutung, wenn Sie die JavaScript-API für Office zur Verarbeitung von Datums- oder Zeitwerten in Outlook-Add-Ins verwenden, die auf einem Outlook Rich Client und in Outlook Web App oder OWA für mobile Geräte ausgeführt werden.


### Zeitzonen für Outlook-Clients

Lassen Sie uns zur Verdeutlichung die fraglichen Zeitzonen definieren.



|**Zeitzone**|**Beschreibung**|
|:-----|:-----|
|Zeitzone des Clientcomputers|Diese wird im Betriebssystem des Clientcomputers festgelegt. Die meisten Browser verwenden die Zeitzone des Clientcomputers, um Datums- oder Zeitwerte des JavaScript-Objekts  **Date** anzuzeigen.<br/><br/>Ein Outlook Rich Client verwendet diese Zeitzone, um Datums- oder Zeitwerte auf der Benutzeroberfläche anzuzeigen. <br/><br/>Beispielsweise verwendet Outlook auf einem Clientcomputer unter Windows die in Windows festgelegte Zeitzone als lokale Zeitzone. Wenn der Benutzer die Zeitzone auf einem Mac-Clientcomputer ändert, fordert Outlook für Mac den Benutzer auf, die Zeitzone auch in Outlook zu aktualisieren.|
|EAC-Zeitzone (Exchange Admin Center)|Dieser Zeitzonenwert (und die bevorzugte Sprache) wird vom Benutzer festgelegt, wenn er sich erstmals bei Outlook Web App oder OWA für mobile Geräte anmeldet. <br/><br/>Outlook Web App und OWA für mobile Geräte verwenden diese Zeitzone, um Datums- oder Zeitwerte auf der Benutzeroberfläche anzuzeigen.|
Da ein Outlook Rich Client die Zeitzone des Clientcomputers verwendet und die Benutzeroberfläche von Outlook Web App und OWA für mobile Geräte die EAC-Zeitzone, kann die lokale Zeitzone für ein für das gleiche Postfach installierte Outlook-Add-In bei Ausführung auf einem Outlook Rich Client und in Outlook Web App oder OWA für mobile Geräte unterschiedlich sein. Als Entwickler von Mail-Add-ins sollten Sie Datumswerte entsprechend ein- und ausgeben, damit diese Werte immer konsistent mit der Zeitzone sind, die der Benutzer auf dem entsprechenden Client erwartet.


### Datumsspezifische API

Die folgenden Eigenschaften und Methoden in der JavaScript-API für Office unterstützen die datumsspezifischen Features.reference/outlook/Office.context.mailbox.item.md



**API-Element**|**Zeitzonendarstellung**|**Beispiel unter einem Outlook Rich Client**|**Beispiel in Outlook Web App oder OWA für mobile Geräte**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|Unter einem Outlook Rich Client gibt diese Eigenschaft die Zeitzone des Clientcomputers zurück. Unter Outlook Web App und OWA für mobile Geräte gibt diese Eigenschaft die Zeitzone EAC zurück. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) und [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Jede dieser Eigenschaften gibt ein JavaScript-Objekt **Date** zurück. Dieser **Date**-Wert ist UTC-gemäß, wie im folgenden Beispiel gezeigt – `myUTCDate` hat in einem Outlook Rich Client, in Outlook Web App und OWA für mobile Geräte denselben Wert.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Wird jedoch  `myDate.getDate` aufgerufen, so wird ein Datumswert in der Zeitzone des Clientcomputers zurückgegeben, der der Zeitzone entspricht, die zum Anzeigen von Datums- und Zeitwerten in der Benutzeroberfläche des Outlook Rich Client verwendet wird, sich aber von der EAC-Zeitzone der Benutzeroberflächen von Outlook Web App und OWA für mobile Geräte unterscheiden kann.|Wird das Element um 9 Uhr UTC erstellt, gibt<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` „4am EST“ zurück.<br/><br/>Bei Änderung des Elements um 11 Uhr UTC, gibt<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` „6am EST“ zurück.|Bei Erstellungszeit 9 Uhr UTC gibt<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` „4am EST“ zurück.<br/><br/>Bei Änderung des Elements um 11 Uhr UTC, gibt<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` „6am EST“ zurück.<br/><br/>Beachten Sie bei der Anzeige der Erstellungs- und Änderungszeiten auf der Benutzeroberfläche, dass Sie diese zunächst in PST konvertieren sollten, damit diese konsistent mit dem Rest der Benutzeroberfläche sind.
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|Sowohl der  _Start_- als auch der  _End_-Parameter benötigt ein JavaScript-Objekt  **Date**. Die Argumente sollten unabhängig von der auf der Benutzeroberfläche von einem Outlook Rich Client, von Outlook Web App oder OWA für mobile Geräte verwendeten Zeitzone UTC-richtig sein.|Wenn die Anfangs- und Endzeiten des Terminformulars 9 Uhr UTC und 11 Uhr UTC angeben, sollten Sie sicherstellen, dass das `start`- und `end`-Argument UTC-richtig sind, das bedeutet:<br/><br/><ul><li>`start.getUTCHours` gibt „9am UTC“ zurück</li><li>`end.getUTCHours` gibt „11am UTC“ zurück</li></ul>|Wenn die Anfangs- und Endzeiten des Terminformulars 9 Uhr UTC und 11 Uhr UTC angeben, sollten Sie sicherstellen, dass das `start`- und `end`-Argument UTC-richtig sind, das bedeutet:<br/><br/><ul><li>`start.getUTCHours` gibt „9am UTC“ zurück</li><li>`end.getUTCHours` gibt „11am UTC“ zurück</li></ul>

## Hilfsmethoden für datumsspezifische Szenarien


Wie in den vorhergehenden Abschnitten beschrieben, bietet die JavaScript-API für Office zwei Hilfsmethoden, da die Ortszeit für Benutzer von Outlook Web App oder OWA für mobile Geräte sich von der in einem Outlook Rich Client unterscheiden kann, das JavaScript-Objekt **Date** aber nur die Umwandlung in die Zeitzone des Clientcomputers oder UTC unterstützt: [Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) und [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md). <br/><br/>
Diese Hilfsmethoden können immer dann verwendet werden, wenn Datum oder Uhrzeit für die folgenden zwei datumsbezogenen Szenarien in einem Outlook Rich Client, in Outlook Web App und OWA für mobile Geräte unterschiedlich behandelt werden müssen.


### Szenario A: Anzeigen der Erstellungs- oder Änderungszeit des Elements

Wenn Sie die Erstellungszeit (**Item.dateTimeCreated**) oder die Änderungszeit (**Item.dateTimeModified**) des Elements auf der Benutzeroberfläche anzeigen, verwenden Sie zunächst  **convertToLocalClientTime**, um das von diesen Eigenschaften zum Abrufen einer Wörterbuchdarstellung bereitgestellte **Date**-Objekt in die entsprechende Ortszeit zu konvertieren. Das Folgende ist ein Beispiel dieses Szenarios:


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook web app or OWA for Devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Beachten Sie, dass  **convertToLocalClientTime** die Unterschiede zwischen einem Outlook Rich Client und Outlook Web App bzw. OWA für mobile Geräte verarbeitet:


- Wenn  **convertToLocalClientTime** feststellt, dass der aktuelle Host ein Rich Client ist, konvertiert die Methode die **Date**-Darstellung in eine Wörterbuchdarstellung in der gleichen Clientcomputer-Zeitzone, die konsistent mit der restlichen Benutzeroberfläche des Rich Client ist.
    
- Wenn  **convertToLocalClientTime** feststellt, dass der aktuelle Host Outlook Web App ist, konvertiert die Methode die UTC-richtige **Date**-Darstellung in ein Wörterbuchformat in der Zeitzone EAC, die konsistent mit dem Rest der Outlook Web App-Benutzeroberfläche ist.
    

### Szenario B: Anzeigen von Anfangs- und Enddatum in einem neuen Terminformular

Wenn Sie als Eingabe verschiedene Teile eines Datum-Uhrzeit-Werts in Ortszeit erhalten und diesen Wert als Start- oder Endzeit in einem Terminformular angeben möchten, verwenden Sie zuerst die Hilfsmethode  **convertToUtcClientTime** zum Umwandeln des Werts in ein UTC-gemäßes **Date**-Objekt.<br/><br/>Im folgenden Beispiel wird davon ausgegangen, dass  `myLocalDictionaryStartDate` und `myLocalDictionaryEndDate` Datum-/Uhrzeit-Werte im Wörterbuchformat sind, das Sie vom Benutzer erhalten haben. Diese Werte basieren auf der lokalen Zeit, die abhängig von der Host-Anwendung ist.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Die resultierenden Werte  `myUTCCorrectStartDate` und `myUTCCorrectEndDate` sind UTC-gemäß. Übergeben Sie diese  **Date**-Objekte dann als Argumente für die Parameter _Start_ und _End_ der Methode **Mailbox.displayNewAppointmentForm**, um das neue Terminformular anzuzeigen.<br/><br/>
Beachten Sie, dass  **convertToLocalClientTime** die Unterschiede zwischen einem Outlook Rich Client und Outlook Web App bzw. OWA für mobile Geräte verarbeitet:


- Wenn  **convertToUtcClientTime** feststellt, dass es sich bei dem aktuellen Host um einen Outlook Rich Client handelt, konvertiert die Methode die Wörterbuchdarstellung einfach in ein **Date**-Objekt. Dieses  **Date**-Objekt ist UTC-richtig, wie von  **displayNewAppointmentForm** erwartet.
    
- Wenn  **convertToUtcClientTime** feststellt, dass es sich bei dem aktuellen Host um Outlook Web App handelt, konvertiert die Methode das Wörterbuchformat der in der Zeitzone EAC angezeigten Datums- und Zeitwerte in ein **Date**-Objekt. Dieses  **Date**-Objekt ist UTC-richtig, wie von  **displayNewAppointmentForm** erwartet.
    

## Weitere Ressourcen



- [Bereitstellen und Installieren von Outlook-Add-Ins zu Testzwecken](../outlook/testing-and-tips.md)
    


