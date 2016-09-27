
# Élément Host
Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.

> **Important** : La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node). Toutefois, la fonctionnalité est identique.  


## Manifeste de base

Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](./officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.   

### Attributs
| Attribut     | Type   | Requis | Description                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Nom](#name) | string | obligatoire | Nom du type d’application hôte Office. |


### Name
Spécifie le type d’hôte ciblé par ce complément. La valeur doit être l’une des suivantes :

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### Exemple
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

---

## Nœud VersionOverrides
Lorsqu’il est défini dans [VersionOverrides](./versionoverrides), le type d’hôte est déterminé par l’attribut `xsi:type`. 

### Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Oui  | Décrit l’hôte d’Office auquel ces paramètres s’appliquent.|

### Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  Oui   |  Définit le facteur de forme affecté. |


### xsi:type
Contrôle à quel hôte Office (Word, Excel, PowerPoint, Outlook, OneNote) s’appliquent également les paramètres contenus. La valeur doit être l’une des suivantes :

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## Exemple d’hôte 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
