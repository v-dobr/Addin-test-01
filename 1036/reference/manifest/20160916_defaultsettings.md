
# Élément DefaultSettings
Spécifie l’emplacement source par défaut et d’autres paramètres par défaut pour votre complément de contenu ou de volet Office.

 **Type de complément :** Application de contenu et de volet Office


## Syntaxe :


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## Contenu dans :

[OfficeApp](../../reference/manifest/officeapp.md)


## Peut contenir :



|**Élément**|**Contenu**|**Application de messagerie**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|x||x|
|[RequestedWidth](../../reference/manifest/requestedwidth.md)|x|||
|[RequestedHeight](../../reference/manifest/requestedheight.md)|x|||

## Remarques

L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de volet Office et de contenu. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut pour les fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](../../reference/manifest/formsettings.md).

