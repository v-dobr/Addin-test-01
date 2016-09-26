
# OfficeApp-Element
Das Stammelement im Manifest eines Office-Add-Ins.

 **Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail


## Syntax:


```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```


## Enthalten in:

 _Keine_


## Muss enthalten:



|**Element**|**Inhalt**|**E-Mails**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[ID](../../reference/manifest/id.md)|x|x|x|
|[Version](../../reference/manifest/version.md)|x|x|x|
|[ProviderName](../../reference/manifest/providername.md)|x|x|x|
|[DefaultLocale](../../reference/manifest/defaultlocale.md)|x|x|x|
|[DefaultSettings](../../reference/manifest/defaultsettings.md)|x|x|x|
|[DisplayName](../../reference/manifest/displayname.md)|x|x|x|
|[Beschreibung](../../reference/manifest/description.md)|x|x|x|
|[FormSettings](../../reference/manifest/formsettings.md)||x||
|[Berechtigungen](../../reference/manifest/permissions.md)|x||x|
|[Regel](../../reference/manifest/rule.md)||x||

## Kann enthalten:



|**Element**|**Inhalt**|**E-Mails**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](../../reference/manifest/alternateid.md)|x|x|x|
|[IconUrl](../../reference/manifest/iconurl.md)|x|x|x|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|x|x|x|
|[SupportUrl](../../reference/manifest/supporturl.md)|x|x|x|
|[AppDomains](../../reference/manifest/appdomains.md)|x|x|x|
|[Hosts](../../reference/manifest/hosts.md)|x|x|x|
|[Anforderungen](../../reference/manifest/requirements.md)|x|x|x|
|[AllowSnapshot](../../reference/manifest/allowsnapshot.md)|x|||
|[Berechtigungen](../../reference/manifest/permissions.md)||x||
|[DisableEntityHighlighting](../../reference/manifest/disableentityhighlighting.md)||x||
|[Dictionary](../../reference/manifest/dictionary.md)|||x|
|[VersionOverrides](../../reference/manifest/versionoverrides.md)|X|X|X|

## Attribute


|||
|:-----|:-----|
|xmlns|Definiert den Office-Add-In-Manifestnamespace und die Schemaversion. Dieses Attribut sollte immer auf `"http://schemas.microsoft.com/office/appforoffice/1.1"` festgelegt werden.|
|xmlns: xsi|Definiert die XML-Schemainstanz. Dieses Attribut sollte immer auf `"http://www.w3.org/2001/XMLSchema-instance"` festgelegt werden.|
|xsi:type|Definiert die Art des Office-Add-Ins. Dieses Attribut sollte immer auf `"ContentApp"`, `"MailApp"` oder `"TaskPaneApp"` festgelegt werden.|
