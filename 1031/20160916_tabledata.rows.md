
# TableData.rows-Eigenschaft
Ruft die Zeilen in einer Tabelle ab oder legt diese fest.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Verfügbar in [Anforderungssatz](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Hinzugefügt in**|1.1|

```
var myRows = tableBindingObj.rows;
```


## Rückgabewert

Gibt ein Array von Arrays zurück, das die Daten in der Tabelle enthält. Gibt ein leeres **array**`[]` zurück, wenn keine Zeilen vorhanden sind.


## Hinweise

Um Zeilen abzugeben, müssen Sie ein Array von Arrays angeben, das der Struktur der Tabelle entspricht. Um beispielsweise zwei Zeilen von **string**-Werten in einer zweispaltigen Tabelle anzugeben, legen Sie die **rows**-Eigenschaft auf ` [['a', 'b'], ['c', 'd']]` fest.

Wenn Sie **null** für die **rows**-Eigenschaft angeben (oder die Eigenschaft beim Konstruieren eines **TableData**-Objekts leer lassen), kommt es beim Ausführen des Codes zu folgenden Ergebnissen:


- Beim Einfügen einer neuen Tabelle wird eine leere Zeile eingefügt.
    
- Beim Überschreiben oder Aktualisieren einer vorhandenen Tabelle werden die vorhandenen Zeilen nicht geändert.
    

## Beispiel

In dem Beispiel unten wird eine Tabelle mit einer Spalte, einer Kopfzeile und drei weiteren Zeilen erstellt.


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}
```


## Supportdetails


Ein Häkchen (v) in der folgenden Matrix weist darauf hin, dass diese Methode in der entsprechenden Office-Hostanwendung unterstützt wird. Eine leere Zelle weist darauf hin, dass die Office-Hostanwendung diese Methode nicht unterstützt.

Weitere Informationen zu den Voraussetzungen der Office-Hostanwendung und des Servers finden Sie unter [Anforderungen zum Ausführen von Office-Add-Ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office für Windows Desktop**|**Office Online (im Browser)**|**Office für iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v|v|v|


|||
|:-----|:-----|
|**Verfügbar in Anforderungssätzen**|TableBindings|
|**Mindestberechtigungsstufe**|[Eingeschränkt](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in-Typen**|Inhalt, Aufgabenbereich|
|**Bibliothek**|Office.js|
|**Namespace**|Büro|

## Supportverlauf



****


|**Version**|**Änderungen**|
|:-----|:-----|
|1.1|Unterstützung für Word Online hinzugefügt.|
|1.1|Unterstützung für Excel und Word in Office für iPad hinzugefügt|
|1.0|Eingeführt|
