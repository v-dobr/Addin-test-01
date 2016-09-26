
# Изменения API JavaScript для Office
В интерфейс API JavaScript для Office периодически добавляются новые и обновленные объекты, методы, свойства, события и перечисления для расширения возможностей ваших Надстройки Office. Используйте следующие ссылки, чтобы ознакомиться с новыми и обновленными элементами API.

Для разработки надстроек с использованием новых элементов API вам потребуется [обновить файлы API JavaScript для Office в проекте](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

Сведения обо всех элементах API, в том числе о тех, которые не изменились по сравнению с предыдущими версиями, см. в статье [API JavaScript для Office](../reference/javascript-api-for-office.md).


## Новые и обновленные интерфейсы API

 **Новые и обновленные объекты**


|**Object**|**Описание**|**Объект**|
|:-----|:-----|:-----|
|[Элемент](../reference/outlook/Office.context.mailbox.item.md)|Добавленная или обновленная версия<br><ul><li><p>Методы <a href="../reference/outlook/Office.context.mailbox.item.md#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> и <a href="../reference/outlook/Office.context.mailbox.item.md#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> для поддержки считывания выделенного пользователем фрагмента и его замены в теме и тексте сообщения или встречи.</p></li><li><p>Методы <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyAllForm" target="_blank">displayReplyAllForm</a> и <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyForm" target="_blank">displayReplyForm</a> для поддержки добавления вложения в форму ответа для встречи.</p></li></ul>|Mailbox 1.2|
|[Элемент](../reference/outlook/Office.context.mailbox.item.md)|Обновлен так, чтобы включать методы и поля для создания надстроек Outlook, активирующихся в режиме создания. |1.1|
|[Binding](../reference/shared/binding.md)|Обновлен для поддержки привязки к таблице в контентных надстройках для Access.|1.1|
|[Bindings](../reference/shared/bindings.bindings.md)|Обновлен для поддержки привязки к таблице в контентных надстройках для Access.|1.1|
|[Body](../reference/outlook/Body.md)|Добавлен для поддержки создания и изменения текста сообщения или встречи в надстройках Outlook, активирующихся в режиме создания.|1.1|
|[Документ](../reference/shared/document.md)|Обновления и дополнения: <ul><li><p>Поддержка свойств <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>, <a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a> и <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> в контентных надстройках для Access.</p></li><li><p>Получение документа в виде PDF-файла с помощью метода <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> в надстройках для PowerPoint и Word.</p></li><li><p>Получение свойств файла с помощью метода <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> в надстройках для Excel, PowerPoint и Word.</p></li><li><p>Переход к расположениям и объектам в документе с помощью метода <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByldAsync</a> в надстройках для Excel и PowerPoint.</p></li><li><p>Получение идентификатора, заголовка и индекса выбранных слайдов с помощью метода <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> (при указании нового перечисления <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a>) в надстройках для PowerPoint.</p></li></ul>|1.1|
|[Расположение](../reference/outlook/Location.md)|Добавлен, чтобы стало возможным задание место встречи в надстройках Outlook, активирующихся в режиме создания.|1.1|
|[Office](../reference/shared/office.md)|Обновлен метод select для поддержки получения привязки в контентных надстройках для Access.|1.1|
|[Получатели](../reference/outlook/Recipients.md)|Добавлен для поддержки получения и установки получателей сообщения или встречи в приложениях режима создания.|1.1|
|[Параметры](../reference/shared/document.settings.md)|Обновлен для поддержки создания пользовательских настроек в контентных надстройках для Access.|1.1|
|[Тема](../reference/outlook/Subject.md)|Добавлен, чтобы стало возможным получение и задание темы сообщения или встречи в надстройках Outlook, активирующихся в режиме создания.|1.1|
|[Time](../reference/outlook/Time.md)|Добавлен, чтобы стало возможным получение и задание времени начала и окончания встречи в надстройках Outlook, активирующихся в режиме создания.|1.1|



**Новые и обновленные перечисления**


|**Object**|**Описание**|**Версия**|
|:-----|:-----|:-----|
|[ActiveView](../reference/shared/activeview-enumeration.md)|Указывает состояние активного представления документа, например возможность редактирования документа пользователем. Добавлен, чтобы надстройки для PowerPoint могли определить, просматривают ли пользователи презентацию ( **Показ слайдов**) или редактируют слайды. |1.1|
|[CoercionType](../reference/shared/coerciontype-enumeration.md)|Добавлен элемент  **Office.CoercionType.SlideRange** для поддержки получения выбранного диапазона слайдов с помощью метода **getSelectedDataAsync** в надстройках для PowerPoint.|1.1|
|[EventType](../reference/shared/eventtype-enumeration.md)|Добавлено новое событие ActiveViewChanged.|1.1|
|[FileType](../reference/shared/filetype-enumeration.md)|Добавлена возможность указания выходного файла в формате PDF.|1.1|
|[GoToType](../reference/shared/gototype-enumeration.md)|Добавлен для указания места или объекта в документе, к которому необходимо перейти.|1.1|

## Дополнительные ресурсы


- [Справка по API и схеме надстроек для Office](../reference/reference.md)
    
- [Надстройки Office](../docs/overview/office-add-ins.md)
    
