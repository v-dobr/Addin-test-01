
# DefaultSettings 元素
指定内容或任务窗格外接程序的默认源位置和其他默认设置。

 **外接程序类型：**内容、任务窗格


## 语法：


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## 包含在：

[OfficeApp](../../reference/manifest/officeapp.md)


## 可以包含：



|**元素**|**内容**|**邮件**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|x||x|
|[RequestedWidth](../../reference/manifest/requestedwidth.md)|x|||
|[RequestedHeight](../../reference/manifest/requestedheight.md)|x|||

## 注解

**DefaultSettings** 元素中的源位置和其他设置仅应用于内容和任务窗格外接程序。对于邮件外接程序，您在 [FormSettings](../../reference/manifest/formsettings.md) 元素中指定源文件的默认位置和其他默认设置。

