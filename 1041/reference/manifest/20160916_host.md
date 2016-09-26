
# Host 要素
アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。

> **重要**:**Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。 ただし、機能は変わりません。  


## 基本のマニフェスト

基本のマニフェストで定義されている場合 ([OfficeApp](./officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。   

### 属性
| 属性     | 種類   | 必須 | 説明                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [名前](#name) | string | 必須 | Office ホスト アプリケーションの種類の名前。 |


### 名前
このアドインが対象にするホストの種類を指定します。 この値は、次のいずれかである必要があります。

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### 例
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

---

## VersionOverrides ノード
[VersionOverrides](./versionoverrides) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。 

### 属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  はい  | これらの設定を適用する Office ホストについて説明します。|

### 子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  はい   |  影響を受けるフォーム ファクターを定義します。 |


### xsi:type
含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。 この値は、次のいずれかである必要があります。

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## ホストの例 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
