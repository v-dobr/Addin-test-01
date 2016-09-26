
# Änderungen in der JavaScript-API für Office
Die JavaScript-API für Office bietet neue und aktualisierte Objekte, Methoden, Eigenschaften, Ereignisse und Enumerationen, welche den Funktionsumfang Ihrer Office-Add-Ins erweitern. Verwenden Sie die Links unten, um die neuen und aktualisierten API-Mitglieder anzuzeigen.

Um Add-Ins mithilfe der neuen API-Mitglieder zu entwickeln, müssen Sie [die Dateien der JavaScript-API für Office in Ihrem Projekt aktualisieren](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

Alle API-Mitglieder, einschließlich derjenigen, die aus vorherigen Updates unverändert übernommen wurden, finden Sie unter [JavaScript API for Office](../reference/javascript-api-for-office.md).


## Neue und aktualisierte API

 **Neue und aktualisierte Objekte**


|**Object**|**Beschreibung**|**Version hinzugefügt oder aktualisiert**|
|:-----|:-----|:-----|
|[Element](../reference/outlook/Office.context.mailbox.item.md)|Updates und Ergänzungen für:<br><ul><li><p>Die Methoden <a href="../reference/outlook/Office.context.mailbox.item.md#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> und <a href="../reference/outlook/Office.context.mailbox.item.md#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a>, um das Abrufen der Auswahl des Benutzers und deren Überschreiben im Betreff und Text einer Nachricht oder eines Termins zu unterstützen.</p></li><li><p>Die Methoden <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyAllForm" target="_blank">displayReplyAllForm</a> und <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyForm" target="_blank">displayReplyForm</a>, um das Hinzufügen einer Anlage zum Antwortformular für einen Termin zu unterstützen.</p></li></ul>|Mailbox 1.2|
|[Element](../reference/outlook/Office.context.mailbox.item.md)|Mit Methoden und Feldern zum Erstellen von Verfassenmodus-Outlook-Add-Ins aktualisiert. |1.1|
|[Bindung](../reference/shared/binding.md)|Unterstützt jetzt Tabellenbindung in Inhalts-Add-ins für Access.|1.1|
|[Bindings](../reference/shared/bindings.bindings.md)|Unterstützt jetzt Tabellenbindung in Inhalts-Add-ins für Access.|1.1|
|[Text](../reference/outlook/Body.md)|Wurde hinzugefügt, um das Erstellen und Bearbeiten des Textkörpers einer Nachricht oder eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|[Dokument](../reference/shared/document.md)|Aktualisierungen und Ergänzungen: <ul><li><p>Unterstützung der Eigenschaften <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>, <a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a> und <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> in Inhalts-Add-ins für Access</p></li><li><p>Abrufen des Dokuments als PDF mit der Methode <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> in Add-ins für PowerPoint und Word</p></li><li><p>Abrufen von Dateieigenschaften mit der Methode <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> in Add-ins für Excel, PowerPoint und Word</p></li><li><p>Navigieren zu Speicherorten und Objekten im Dokument mit der Methode <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> in Add-ins für Excel und PowerPoint</p></li><li><p>Abrufen der ID, des Titels und des Indexes für ausgewählte Folien mit der Methode <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> (wenn Sie die neue <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a>-Enumeration angeben) in Add-ins für PowerPoint</p></li></ul>|1.1|
|[Ort](../reference/outlook/Location.md)|Wurde hinzugefügt, um das Festlegen des Orts eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|[Büro](../reference/shared/office.md)|Die Auswahlmethode unterstützt jetzt das Abrufen von Bindungen in Inhalts-Add-ins für Access.|1.1|
|[Empfänger](../reference/outlook/Recipients.md)|Wurde hinzugefügt, um das Abrufen und Festlegen der Empfänger einer Nachricht oder eines Termins im Verfassenmodus zu ermöglichen.|1.1|
|[Einstellungen](../reference/shared/document.settings.md)|Unterstützt jetzt das Erstellen benutzerdefinierter Einstellungen in Inhalts-Add-ins für Access.|1.1|
|[Betreff](../reference/outlook/Subject.md)|Wurde hinzugefügt, um das Abrufen und Festlegen des Betreffs einer Nachricht oder eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|[Zeit](../reference/outlook/Time.md)|Wurde hinzugefügt, um das Abrufen und Festlegen der Start- und Endzeit eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|



**Neue und aktualisierte Enumerationen**


|**Object**|**Beschreibung**|**Version**|
|:-----|:-----|:-----|
|[ActiveView](../reference/shared/activeview-enumeration.md)|Gibt den Status der aktiven Ansicht des Dokuments an, z. B. ob der Benutzer das Dokument bearbeiten kann.Wurde hinzugefügt, damit Add-ins für PowerPoint ermitteln können, ob die Benutzer die Präsentation ( **Diaschau**) anzeigen oder Folien bearbeiten. |1.1|
|[CoercionType](../reference/shared/coerciontype-enumeration.md)|Mit  **Office.CoercionType.SlideRange** aktualisiert, um das Abrufen des ausgewählten Folienbereichs mit der Methode **getSelectedDataAsync** in Add-ins für PowerPoint zu unterstützen.|1.1|
|[EventType](../reference/shared/eventtype-enumeration.md)|Mit dem neuen ActiveViewChanged-Ereignis aktualisiert.|1.1|
|[FileType](../reference/shared/filetype-enumeration.md)|Gibt jetzt Ausgaben in PDF-Format an.|1.1|
|[GoToType](../reference/shared/gototype-enumeration.md)|Wurde hinzugefügt, um die Stelle oder das Objekt im Dokument anzugeben, zu dem gewechselt werden soll.|1.1|

## Weitere Ressourcen


- [API und Schemaverweise für Office-Add-Ins](../reference/reference.md)
    
- [Office-Add-Ins](../docs/overview/office-add-ins.md)
    
