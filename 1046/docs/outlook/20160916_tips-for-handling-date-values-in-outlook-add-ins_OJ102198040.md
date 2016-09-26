
# Dicas para lidar com valores de data em suplementos do Outlook

A API JavaScript para Office usa o objeto JavaScript [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp) para a maioria dos processos de armazenamento e recuperação de datas e horas. Esse objeto **Date** fornece métodos como [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp) e [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp), que retornam o valor de data ou hora selecionado de acordo com o UTC (Tempo Universal Coordenado).<br/><br/>
O objeto **Date** também fornece outros métodos como [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp) e [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp), que retornam a data ou a hora solicitada de acordo com a "hora local".<br/><br/>O conceito de "hora local" é basicamente determinado pelo navegador e pelo sistema operacional no computador cliente. Por exemplo, na maioria dos navegadores em execução em um computador cliente baseado no Windows, uma chamada JavaScript para **getDate** retorna uma data com base no fuso horário definido no Windows do computador cliente.<br/><br/>
O exemplo a seguir cria um objeto **Date**<code>myLocalDate</code> na hora local e chama **toUTCString** para converter essa data em uma cadeia de caracteres de data em UTC.




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Embora você possa usar o objeto JavaScript **Date** para obter um valor de data ou hora baseado em UTC ou no fuso horário do computador cliente, o objeto **Date** é limitado em um aspecto: ele não oferece métodos para retornar um valor de data ou hora para qualquer outro fuso horário específico. Por exemplo, se seu computador cliente estiver definido para EST (hora oficial do leste dos EUA), não existe método **Date** que permita que você obtenha o valor de hora diferente de EST ou UTC, como PST (hora oficial do Pacífico).


## Recursos relacionados a data para suplementos do Outlook


A limitação de JavaScript mencionada tem implicações quando você usa a API JavaScript para Office a fim de manipular valores de data ou hora em suplementos do Outlook que são executados em um cliente avançado do Outlook e no Outlook Web App ou no OWA para Dispositivos.


### Fusos horários para clientes do Outlook

Para maior clareza, vamos definir os fusos horários em questão.



|**Fuso horário**|**Descrição**|
|:-----|:-----|
|Fuso horário do computador cliente|Isso é definido no sistema operacional do computador cliente. A maioria dos navegadores usa o fuso horário do computador cliente para exibir os valores de data ou hora do objeto JavaScript **Date**.<br/><br/>Um cliente avançado do Outlook usa esse fuso horário para exibir os valores de data ou hora na interface do usuário. <br/><br/>Por exemplo, em um computador cliente executando o Windows, o Outlook usa o fuso horário definido no Windows como o fuso horário local. No Mac, se o usuário alterar o fuso horário no computador cliente, o Outlook para Mac solicita ao usuário que atualize o fuso horário no Outlook também.|
|Fuso horário do EAC (Centro de Administração do Exchange)|O usuário define o valor de fuso horário (e o idioma preferido) quando faz logon no Outlook Web App ou no OWA para Dispositivos pela primeira vez. <br/><br/>O Outlook Web App e o OWA para Dispositivos usam esse fuso horário para exibir os valores de data ou hora na interface do usuário.|
Como um cliente avançado do Outlook usa o fuso horário do computador cliente e a interface do usuário do Outlook Web App e do OWA para Dispositivos usa o fuso horário do EAC, a hora local para o mesmo suplemento instalado na mesma caixa de correio pode ser diferente durante a execução em um cliente avançado do Outlook e no Outlook Web App ou no OWA para Dispositivos. Como desenvolvedor de suplementos do Outlook, você deve fornecer valores de data de entrada e saída de forma que sejam sempre consistentes com o fuso horário que o usuário espera no cliente correspondente.


### API relacionada à data

A seguir estão as propriedades e métodos da API JavaScript para Office que dão suporte a features.reference/outlook/Office.context.mailbox.item.md relacionados à data



**Membro da API**|**Representação de fuso horário**|**Exemplo em um cliente avançado do Outlook**|**Exemplo no Outlook Web App ou no OWA para Dispositivos**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|Em um cliente avançado do Outlook, essa propriedade retorna o fuso horário do computador cliente. No Outlook Web App e no OWA para Dispositivos, essa propriedade retorna o fuso horário do EAC. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) e [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Cada uma dessas propriedades retorna um objeto JavaScript **Date**. Esse valor **Date** é corrigido para UTC, conforme mostrado no exemplo a seguir: `myUTCDate` tem o mesmo valor em um cliente avançado do Outlook, no Outlook Web App e no OWA para Dispositivos.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>No entanto, chamar `myDate.getDate` retorna um valor de data no fuso horário do computador cliente, que é consistente com o fuso horário usado para exibir valores de data e hora na interface do cliente avançado do Outlook, mas pode ser diferente do fuso horário do EAC que o Outlook Web App e o OWA para Dispositivos usam em suas interfaces do usuário.|Se o item é criado às 9h UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` é retornado às 4h EST.<br/><br/>Se o item for modificado às 11h UTC:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` é retornado às 6h EST.|Se a hora de criação do item for às 9h UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` é retornado às 4h EST.<br/><br/>Se o item for modificado às 11h UTC:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` é retornado às 6h EST.<br/><br/>Observe que se você quer exibir a hora de criação ou de alteração na interface do usuário, convém primeiro converter a hora em PST para ficar consistente com o resto da interface do usuário.
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|Cada um dos parâmetros _Start_ e _End_ requer um objeto JavaScript **Date**. Os argumentos devem ser corrigidos para UTC independentemente do fuso horário usado na interface do usuário de um cliente avançado do Outlook, do Outlook Web App ou do OWA para Dispositivos.|Se as horas de início e de término para o formulário de compromisso são 9h UTC e 11h UTC, você deve fazer com que os argumentos `start` e `end` estejam corretos em relação à UTC, o que significa que :<br/><br/><ul><li>`start.getUTCHours` é retornado às 9h UTC</li><li>`end.getUTCHours` é retornado às 11h UTC</li></ul>|Se as horas de início e de término para o formulário de compromisso são 9h UTC e 11h UTC, você deve fazer com que os argumentos `start` e `end` estejam corretos em relação à UTC, o que significa que :<br/><br/><ul><li>`start.getUTCHours` é retornado às 9h UTC</li><li>`end.getUTCHours` é retornado às 11h UTC</li></ul>

## Métodos auxiliares para cenários de data


Conforme descrito nas seções anteriores, como a “hora local” de um usuário no Outlook Web App ou no OWA para Dispositivos pode ser diferente em um cliente avançado do Outlook, mas o objeto JavaScript **Date** dá suporte ao objeto convertendo somente o fuso horário do computador cliente ou UTC, a API JavaScript para Office oferece dois métodos auxiliares: [Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) e [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md). <br/><br/>
Esses métodos auxiliares cuidam das necessidades de lidar com data ou hora diferentes nos dois cenários de datas a seguir, em um cliente avançado do Outlook, no Outlook Web App e no OWA para Dispositivos, impondo "gravação única" para clientes diferentes do seu suplemento.


### Cenário A: exibir a criação de item ou a hora da alteração

Se você estiver exibindo a hora de criação do item (**Item.dateTimeCreated**) ou a hora da modificação (**Item.dateTimeModified**) na interface do usuário, primeiro use **convertToLocalClientTime** para converter o objeto **Date** fornecido por essas propriedades a fim de obter uma representação de dicionário no horário local apropriado. Em seguida, exiba as partes da data do dicionário. A seguir, um exemplo desse cenário:


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

Observe que **convertToLocalClientTime** se encarrega da diferença entre um cliente avançado do Outlook e o Outlook Web App ou o OWA para Dispositivos:


- Se **convertToLocalClientTime** detecta que o host atual é um cliente avançado, o método converte a representação **Date** em uma representação de dicionário no mesmo fuso horário do computador cliente, consistente com o resto da interface do usuário do cliente avançado.
    
- Se **convertToLocalClientTime** detecta que o host atual é o Outlook Web App ou o OWA para Dispositivos, o método converte a representação corrigida para **Date** UTC em um formato de dicionário no fuso horário do EAC, consistente com o restante da interface do usuário do Outlook Web App ou do OWA para Dispositivos.
    

### Cenário B: exibir datas de início e de término em um formulário de novo compromisso

Se você está obtendo como entrada diferentes partes de um valor de data e hora representadas na hora local e quer fornecer esse valor de entrada de dicionário como uma hora de início ou de término em um formulário de compromisso, primeiro use o método auxiliar **convertToUtcClientTime** para converter o valor do dicionário em um objeto **Date** com correção para UTC.<br/><br/>No exemplo a seguir, assuma que `myLocalDictionaryStartDate` e `myLocalDictionaryEndDate` são valores de data e hora em formato de dicionário que você obteve do usuário. Esses valores se baseiam no horário local, dependendo do aplicativo host.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Os valores resultantes, `myUTCCorrectStartDate` e `myUTCCorrectEndDate`, são corrigidos para UTC. Passe esses objetos **Date** como argumentos para os parâmetros _Start_ e _End_ do método **Mailbox.displayNewAppointmentForm** para exibir o formulário do novo compromisso.<br/><br/>
Observe que **convertToUtcClientTime** se encarrega da diferença entre um cliente avançado do Outlook e o Outlook Web App ou o OWA para Dispositivos:


- Se **convertToUtcClientTime** detecta que o host atual é um cliente avançado do Outlook, o método simplesmente converte a representação de dicionário em um objeto **Date**. Esse objeto **Date** é corrigido para UTC, conforme o esperado por **displayNewAppointmentForm**.
    
- Se **convertToUtcClientTime** detecta que o host atual é o Outlook Web App ou o OWA para Dispositivos, o método converte o formato de dicionário dos valores de data e hora no fuso horário do EAC em um objeto **Date**. Esse objeto **Date** é corrigido para UTC, conforme o esperado por **displayNewAppointmentForm**.
    

## Recursos adicionais



- [Implantar e instalar suplementos do Outlook para teste](../outlook/testing-and-tips.md)
    


