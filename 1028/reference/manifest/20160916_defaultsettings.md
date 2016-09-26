
# DefaultSettings 項目
指定您的內容或工作窗格增益集的預設來源位置和其他的預設設定。

 **增益集類型︰**內容、工作窗格


## 語法：


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## 內含於：

[OfficeApp](../../reference/manifest/officeapp.md)


## 可以包含︰



|**元素**|**內容**|**郵件**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|x||x|
|[RequestedWidth](../../reference/manifest/requestedwidth.md)|x|||
|[RequestedHeight](../../reference/manifest/requestedheight.md)|x|||

## 備註

在 **DefaultSettings** 項目中的來源位置和其他設定，只會套用至內容和工作窗格增益集。您可以為郵件增益集，在 [FormSettings](../../reference/manifest/formsettings.md) 項目中，指定來源檔案和其他預設設定的預設位置。

