
# Советы по использованию значений дат в надстройках Outlook

В интерфейсе API JavaScript для Office для хранения и извлечения даты и времени используется преимущественно объект JavaScript [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp). Такой объект **Date** обеспечивает методы [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp) и [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp), которые возвращают запрос значения даты и времени в формате всемирного координированного времени (UTC).<br/><br/>
Объект **Date** обеспечивает также другие методы, например [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp) и [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp), которые возвращают запрос даты или времени по "местному времени".<br/><br/>Понятие "местного времени" в значительной мере определяется браузером и операционной системой на клиентском компьютере. Например, в большинстве браузеров, установленных на клиентских компьютерах под управлением Windows, при вызове метода JavaScript **getDate**, возвращается дата на основе часового пояса, установленного в операционной системе Windows на клиентском компьютере.<br/><br/>
В указанном ниже примере создается объект **Date** с параметром <code>myLocalDate</code> по местному времени с вызовом метода **toUTCString** для преобразования этой даты в строку даты во время в формате UTC.




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Объект JavaScript **Date** можно использовать для получения значения даты или времени на основе времени в формате UTC или местного времени клиентского компьютера, однако объект **Date** ограничен тем, что не поддерживает метод возвращения значения даты или времени для определенного часового пояса. Например, если клиентский компьютер настроен для использования восточного поясного времени (EST), нет ни одного метода **Date**, который позволит получить значение часа в часовом поясе, отличном от EST или UTC, например значение часа по тихоокеанскому времени (PST).


## Функции надстроек Outlook, связанные с датой


Описанное выше ограничение JavaScript создает определенные трудности для разработчика, использующего интерфейс API JavaScript для Office для обработки значений даты и времени в надстройках Outlook, которые работают в расширенном клиенте Outlook, Outlook Web App и Outlook Web App для устройств.


### Часовые пояса для клиентов Outlook

Во избежание недоразумений дадим определение часовым поясам.



|**Часовой пояс**|**Описание**|
|:-----|:-----|
|Часовой пояс клиентского компьютера|Устанавливается в операционной системе на клиентском компьютере. В большинстве браузеров для отображения значений даты и времени объекта JavaScript **Date** используется часовой пояс клиентского компьютера.<br/><br/>В расширенном клиенте Outlook используется этот часовой пояс для отображения значений даты и времени в пользовательском интерфейсе. <br/><br/>Например, на клиентском компьютере под управлением Windows в Outlook используется часовой пояс, установленный в операционной системе Windows в качестве местного часового пояса. На компьютерах Mac, если пользователь изменяет часовой пояс на клиентском компьютере, в приложении Outlook для Mac пользователю будет предложено также обновить часовой пояс в Outlook.|
|Часовой пояс Центра администрирования Exchange (EAC)|Пользователи устанавливают это значение часового пояса (и предпочитаемый язык), когда они в первый раз входят в систему в Outlook Web App или Outlook Web App для устройств. <br/><br/>В Outlook Web App и Outlook Web App для устройств этот часовой пояс используется для отображения значений даты и времени в пользовательском интерфейсе.|
Так как в расширенном клиенте Outlook используется часовой пояс клиентского компьютера, а в пользовательском интерфейсе Outlook Web App и Outlook Web App для устройств — часовой пояс Центра администрирования Exchange, местное время для одной и той же надстройки, установленной для одного и того же почтового ящика, может различаться в зависимости от того, где она запущена: в расширенном клиенте Outlook, Outlook Web App или Outlook Web App для устройств. Разработчику надстройки Outlook следует продумать ввод и вывод значений даты, чтобы эти значения всегда согласовывались с часовым поясом, который пользователь ожидает увидеть в соответствующем клиенте.


### Интерфейс API, связанный с датой

Ниже приведены свойства и методы в интерфейсе API JavaScript для Office, которые поддерживают функциональные возможности, связанные с датой.reference/outlook/Office.context.mailbox.item.md



**Элемент API**|**Представление часового пояса**|**Пример в расширенном клиенте Outlook**|**Пример в Outlook Web App или Outlook Web App для устройств**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|В расширенном клиенте Outlook это свойство возвращает часовой пояс клиентского компьютера. В Outlook Web App и Outlook Web App для устройств это свойство возвращает часовой пояс Центра администрирования Exchange. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) и [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Каждое из этих свойств возвращает объект JavaScript **Date**. Это значение **Date** правильное относительно UTC, как показано в следующем примере: для параметра `myUTCDate` указано одинаковое значение в расширенном клиенте Outlook, Outlook Web App и Outlook Web App для устройств.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Но при вызове метода `myDate.getDate` возвращается значение даты в часовом поясе клиентского компьютера, соответствующем часовому поясу, который используется для отображения значений времени и даты в интерфейсе расширенного клиента Outlook, но который может отличаться от часового пояса Центра администрирования Exchange в пользовательском интерфейсе Outlook Web App и Outlook Web App для устройств.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.<br/><br/>Обратите внимание, что если необходимо отобразить время создания или изменения в пользовательском интерфейсе, следует сначала преобразовать время в формат PST, чтобы оно соответствовало формату времени остального пользовательского интерфейса.
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|Для каждого из параметров _Start_ и _End_ требуется объект JavaScript **Date**. Аргументы должны быть правильными относительно UTC независимо от формата времени пользовательского интерфейса расширенного клиента Outlook, Outlook Web App или Outlook Web App для устройств.|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что:<br/><br/><ul><li>для метода `start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC;</li><li>для метода `end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC.</li></ul>|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>для метода `start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC;</li><li>для метода `end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC.</li></ul>

## Вспомогательные методы для сценариев, связанных с датами


Как описано в предыдущих разделах, поскольку "местное время" для пользователя в Outlook Web App или Outlook Web App для устройств может отличаться на расширенном клиенте Outlook, а объект JavaScript **Date** поддерживает преобразование только в часовой пояс клиентского компьютера или формат времени UTC, интерфейс API JavaScript для Office обеспечивает два вспомогательных метода: [Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) и [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md). <br/><br/>
Эти вспомогательные методы помогают избежать необходимости по-другому обрабатывать время или дату для двух указанных ниже сценариев в расширенном клиенте Outlook, Outlook Web App и Outlook Web App для устройств, таким образом укрепляя принцип "однократной записи" для разных клиентов надстройки.


### Сценарий A. Отображение времени создания или изменения элементов

При отображении времени создания (**Item.dateTimeCreated**) или времени изменения (**Item.dateTimeModified**) элемента в пользовательском интерфейсе метод **convertToLocalClientTime** используется в первый раз для преобразования объекта **Date**, предоставленного этими свойствами, с целью получения представления словаря в соответствующем местном времени. Затем отображаются части даты словаря. Ниже приведен пример этого сценария.


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook web app or OWA for Devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Обратите внимание, что в методе **convertToLocalClientTime** учтено различие между расширенным клиентом Outlook, Outlook Web App или Outlook Web App для устройств:


- Если метод **convertToLocalClientTime** воспринимает текущий узел как расширенный клиент, метод преобразует представление **Date** в представление словаря с использованием местного часового пояса клиентского компьютера, что согласуется с остальным пользовательским интерфейсом расширенного клиента.
    
- Если метод **convertToLocalClientTime** обнаруживает, что на текущем узле используется Outlook Web App или Outlook Web App для устройств, метод преобразует представление **Date**, сверенное с UTC, в формат словаря в часовом поясе Центра администрирования Exchange в соответствии с остальным пользовательским интерфейсом Outlook Web App или Outlook Web App для устройств.
    

### Сценарий Б. Отображение дат начала и окончания в новой форме встречи

При получении в качестве ввода разных частей значения времени и даты, представленных в формате местного времени, и предоставлении этого ввода значения словаря как времени начала или окончания в форме встречи, сначала используйте вспомогательный метод **convertToUtcClientTime**, чтобы преобразовать значение словаря в объект **Date**, правильный относительно UTC.<br/><br/>В указанном ниже примере предположим, что `myLocalDictionaryStartDate` и `myLocalDictionaryEndDate` — значения даты и времени в формате словаря, полученные от пользователя. Эти значения берут за основу местное время, зависящее от ведущего приложения.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

В результате получаются значения `myUTCCorrectStartDate` и `myUTCCorrectEndDate`, правильные относительно UTC. Затем передайте эти объекты **Date** как аргументы для параметров _Start_ и _End_ метода **Mailbox.displayNewAppointmentForm**, чтобы отобразить форму новой встречи.<br/><br/>
Обратите внимание, что в методе **convertToUtcClientTime** учитывается различие между расширенным клиентом Outlook, Outlook Web App или Outlook Web App для устройств:


- Если метод **convertToUtcClientTime** обнаруживает, что текущий узел является расширенным клиентом Outlook, метод просто преобразует представление словаря в объект **Date**. Этот объект **Date** является правильным относительно UTC, что и ожидается в методе **displayNewAppointmentForm**.
    
- Если метод **convertToUtcClientTime** обнаруживает, что текущим узлом является Outlook Web App или Outlook Web App для устройств, метод преобразует формат значений словаря даты и времени, выраженных в значениях часового пояса Центра администрирования Exchange, в объект **Date**. Этот объект **Date** является правильным с точки зрения UTC, что и ожидается в методе **displayNewAppointmentForm**.
    

## Дополнительные ресурсы



- [Развертывание и установка надстроек Outlook для тестирования](../outlook/testing-and-tips.md)
    


