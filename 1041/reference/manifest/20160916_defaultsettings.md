
# DefaultSettings 要素
コンテンツ アドインまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ


## 構文:


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## 次に含まれる:

[OfficeApp](../../reference/manifest/officeapp.md)


## 含めることができるもの:



|**要素**|**コンテンツ**|**メール**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|x||x|
|[RequestedWidth](../../reference/manifest/requestedwidth.md)|x|||
|[RequestedHeight](../../reference/manifest/requestedheight.md)|x|||

## 解説

**DefaultSettings** 要素のソースの場所と他の設定が適用されるのは、コンテンツ アドインと作業ウィンドウ アドインのみです。メール アドインの場合は、ソース ファイルの既定の場所とその他の既定の設定を [FormSettings](../../reference/manifest/formsettings.md) 要素に指定します。

