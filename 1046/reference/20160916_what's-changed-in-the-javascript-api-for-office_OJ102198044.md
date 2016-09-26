
# O que mudou na API JavaScript para Office
A API JavaScript para Office é atualizada periodicamente com objetos, métodos, propriedades, eventos e enumerações novos e atualizados para estender a funcionalidade dos seus suplementos do Office. Use os links abaixo para ver os membros da API novos e atualizados.

Para desenvolver suplementos usando novos membros da API, você precisa [atualizar os arquivos da API JavaScript para Office em seu projeto](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

Para ver todos os membros da API, incluindo aqueles que não sofreram alterações nas atualizações anteriores, confira [API JavaScript para Office](../reference/javascript-api-for-office.md).


## APIs novas e atualizadas

 **Objetos novos e atualizados**


|**Objeto**|**Descrição**|**Versão adicionada ou atualizada**|
|:-----|:-----|:-----|
|[Item](../reference/outlook/Office.context.mailbox.item.md)|Atualizações e adições para:<br><ul><li><p>Os métodos <a href="../reference/outlook/Office.context.mailbox.item.md#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> e <a href="../reference/outlook/Office.context.mailbox.item.md#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> para oferecer suporte à obtenção da seleção do usuário e da substituição no assunto ou corpo de uma mensagem ou um compromisso.</p></li><li><p>Os métodos <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyAllForm" target="_blank">displayReplyAllForm</a> e <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyForm" target="_blank">displayReplyForm</a> para oferecer suporte à adição de um anexo ao formulário de resposta de um compromisso.</p></li></ul>|Mailbox 1.2|
|[Item](../reference/outlook/Office.context.mailbox.item.md)|Atualizado para incluir campos e métodos para criação de suplementos do Outlook no modo de composição. |1.1|
|[Associação](../reference/shared/binding.md)|Atualizado para dar suporte à associação de tabela em suplementos de conteúdo para o Access.|1.1|
|[Bindings](../reference/shared/bindings.bindings.md)|Atualizado para dar suporte à associação de tabela em suplementos de conteúdo para o Access.|1.1|
|[Corpo](../reference/outlook/Body.md)|Adicionado para permitir a criação e edição do corpo de uma mensagem ou compromisso em suplementos do Outlook no modo de composição.|1.1|
|[Documento](../reference/shared/document.md)|Atualizações e adições para: <ul><li><p>Suporte às propriedades <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>, <a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a> e <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> em suplementos de conteúdo para o Access.</p></li><li><p>Navegar até locais e objetos dentro do documento com o método <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> em suplementos para Excel e PowerPoint.</p></li><li><p>Obter as propriedades de arquivo com o método <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> em suplementos para Excel, PowerPoint e Word.</p></li><li><p>Navegar até locais e objetos dentro do documento com o método <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> em suplementos para Excel e PowerPoint.</p></li><li><p>Obter a identificação, o título e o índice dos slides selecionados com o método <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> (ao especificar a nova enumeração <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a>) em suplementos do PowerPoint.</p></li></ul>|1.1|
|[Local](../reference/outlook/Location.md)|Adicionado a fim de permitir a configuração do local de um compromisso em suplementos do Outlook no modo de composição.|1.1|
|[Office](../reference/shared/office.md)|Método de seleção atualizado para oferecer suporte à obtenção de associações em suplementos de conteúdo para Access.|1.1|
|[Destinatários](../reference/outlook/Recipients.md)|Adicionado para permitir a obtenção e configuração dos destinatários de uma mensagem ou de um compromisso no modo de composição.|1.1|
|[Configurações](../reference/shared/document.settings.md)|Atualizado para oferecer suporte à criação de configurações personalizadas em suplementos de conteúdo para o Access.|1.1|
|[Assunto](../reference/outlook/Subject.md)|Adicionado para permitir a obtenção e configuração do assunto de uma mensagem ou de um compromisso em suplementos do Outlook no modo de composição.|1.1|
|[Hora](../reference/outlook/Time.md)|Adicionado para permitir a obtenção e configuração das horas de início e de término de um compromisso em suplementos do Outlook no modo de composição.|1.1|



**Descrição**


|**Objeto**|**Descrição**|**Versão**|
|:-----|:-----|:-----|
|[ActiveView](../reference/shared/activeview-enumeration.md)|Adicionado para que os suplementos do PowerPoint possam determinar se os usuários estão exibindo a apresentação (**Apresentação de Slides**) ou editando slides. |1.1|
|[CoercionType](../reference/shared/coerciontype-enumeration.md)|Atualizado com **Office.CoercionType.SlideRange** para oferecer suporte à obtenção do intervalo de slides selecionado com o método **getSelectedDataAsync** em suplementos para PowerPoint.|1.1|
|[EventType](../reference/shared/eventtype-enumeration.md)|Atualizado para incluir o novo evento ActiveViewChanged.|1.1|
|[FileType](../reference/shared/filetype-enumeration.md)|Atualizado para especificar a saída no formato PDF.|1.1|
|[GoToType](../reference/shared/gototype-enumeration.md)|Adicionado para especificar o local ou objeto a ser acessado no documento.|1.1|

## Recursos adicionais


- [API de suplementos do Office e referências de esquema](../reference/reference.md)
    
- [Suplementos do Office](../docs/overview/office-add-ins.md)
    
