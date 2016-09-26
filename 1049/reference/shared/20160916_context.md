
# Объект Context
Представляет среду выполнения надстройки и открывает доступ к ключевым объектам API.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```
Office.context
```


## Элементы

|||
|:-----|:-----|
|Имя|Описание|
|[commerceAllowed](../../reference/shared/office.context.commerceallowed.md)|Получает сведения о том, запущена ли надстройка на платформе, допускающей добавление ссылок на внешние системы платежей.|
|[contentLanguage](../../reference/shared/office.context.contentlanguage.md)|Получает региональную настройку (язык) для данных, хранимых в документе или элементе.|
|[displayLanguage](../../reference/shared/office.context.displaylanguage.md)|Получает региональную настройку (язык) для пользовательского интерфейса приложения размещения.|
|[document](../../reference/shared/office.context.document.md)|Получает объект, представляющий документ, с которым взаимодействует контентная надстройка или надстройка области задач.|
|[mailbox](../../reference/shared/office.context.mailbox.md)|Получает объект **mailbox**, который предоставляет доступ к элементам API, предназначенным для надстроек Outlook.|
|[officeTheme](../../reference/shared/office.context.officetheme.md)|Предоставляет доступ к свойствам цветов темы Office.|
|[Пользовательский интерфейс](../../reference/shared/officeui)|Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, и управления ими.|
|[roamingSettings](../../reference/shared/office.context.roamingsettings.md)|Получает объект, который представляет сохраненные настраиваемые параметры надстройки.|
|[touchEnabled](../../reference/shared/office.context.touchenabled.md)|Получает сведения о том, запущена ли надстройка в ведущем приложении Office, поддерживающем сенсорный ввод.|

## Заметки

Объект **Context** предоставляет доступ к ключевым объектам API JavaScript для Office.


## Сведения о поддержке



|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Надстройки области задач, надстройки Outlook, контентные надстройки|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлены свойства **commerceAllowed** и **touchEnabledAdded** (только в Excel, PowerPoint и Word в составе Office для iPad).|
|1.1|Добавлена поддержка надстроек для Excel и Word в Office для iPad.|
|1.1|Для элементов [contentLanguage](../../reference/shared/office.context.contentlanguage.md), [displayLanguage](../../reference/shared/office.context.displaylanguage.md) и [document](../../reference/shared/office.context.document.md) добавлена поддержка контентных надстроек Access.|
|1.0|Представлено|
