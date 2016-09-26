
# DefaultSettings-Element
Gibt den Standardspeicherort für die Quelle und andere Standardeinstellungen für Ihr Inhalts- oder Aufgabenbereich-Add-In an.

 **Add-In-Typ:** Inhalt, Aufgabenbereich


## Syntax:


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## Enthalten in:

[OfficeApp](../../reference/manifest/officeapp.md)


## Kann enthalten:



|**Element**|**Inhalt**|**E-Mails**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|x||x|
|[RequestedWidth](../../reference/manifest/requestedwidth.md)|x|||
|[RequestedHeight](../../reference/manifest/requestedheight.md)|x|||

## Bemerkungen

Der Quellspeicherort und andere Einstellungen im **DefaultSettings**-Element gelten nur für Inhalts- und Aufgabenbereich-Add-Ins. Bei E-Mail-Add-Ins geben Sie die Standardspeicherorte für Quelldateien und andere Standardeinstellungen im [FormSettings](../../reference/manifest/formsettings.md)-Element an.

