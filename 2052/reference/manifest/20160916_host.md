
# Host 元素
指定应在其中激活外接程序的单个 Office 应用程序类型。

> **重要说明**：**Host** 元素的语法根据该元素是否在[基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。 但功能相同。  


## 基本清单

在基本清单（在 [OfficeApp](./officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。   

### 属性
| 属性     | 类型   | 必需 | 说明                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | 字符串 | 必需 | Office 主机应用程序的类型名称。 |


### 名称
指定此外接程序面向的主机类型。 值必须为以下值之一：

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### 示例
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

---

## VersionOverrides 节点
在 [VersionOverrides](./versionoverrides) 中定义时，主机类型由 `xsi:type` 属性决定。 

### 属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 描述这些设置所适应的 Office 主机。|

### 子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  是   |  定义受影响的外形规则。 |


### xsi:type
控制所包含的设置也适用的 Office 主机类别（Word、Excel、PowerPoint、Outlook 和 OneNote）。 值必须为以下值之一：

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## 主机示例 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
