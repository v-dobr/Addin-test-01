
# Host 元素
指定應該啟動增益集所在的個別 Office 應用程式類型。

> **重要事項**：**Host** 元素語法會根據元素是在[基本資訊清單](#basic-manifest)或 [VersionOverrides](#versionoverrides-node) 節點內定義而有所不同。 不過，功能都是一樣的。  


## 基本資訊清單

在基本資訊清單中定義時 ([OfficeApp](./officeapp.md) 底下)，主機類型是由 `Name` 屬性決定。   

### 屬性
| 屬性     | 類型   | 必要 | 說明                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [名稱](#name) | 字串 | 必要 | Office 主應用程式的類型名稱。 |


### 名稱
指定此增益集之目標的主機類型。 此值必須是下列任一項：

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### 範例
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

---

## VersionOverrides 節點
在 [VersionOverrides](./versionoverrides) 中定義時，主機類型是由 `xsi:type` 屬性決定。 

### 屬性

|  屬性  |  必要  |  說明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 說明套用這些設定的 Office 主應用程式。|

### 子元素

|  元素 |  必要  |  說明  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  是   |  定義受影響的表單係數。 |


### xsi:type
控制也套用包含的設定的 Office 主應用程式 (Word、Excel、PowerPoint、Outlook、OneNote)。 此值必須是下列任一項：

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## 主機範例 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
