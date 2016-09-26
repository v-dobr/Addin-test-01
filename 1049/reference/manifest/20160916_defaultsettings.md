
# Элемент DefaultSettings
Указывает исходное расположение по умолчанию и другие стандартные параметры для контентной надстройки или надстройки области задач.

 **Тип надстройки:** контентные надстройки и надстройки области задач.


## Синтаксис:


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## Элемент, в котором содержится:

[элемент OfficeApp](../../reference/manifest/officeapp.md)


## Может содержать:



|**Элемент**|**Содержимое**|**Почтовое**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|x||x|
|[RequestedWidth](../../reference/manifest/requestedwidth.md)|x|||
|[RequestedHeight](../../reference/manifest/requestedheight.md)|x|||

## Замечания

Исходное расположение и другие параметры в элементе **DefaultSettings** применяются только к надстройкам области задач и контентным надстройкам. В случае почтовых надстроек следует задавать расположения по умолчанию для исходных файлов и другие стандартные параметры с помощью элемента [FormSettings](../../reference/manifest/formsettings.md).

