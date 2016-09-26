# Visão geral da programação da API JavaScript do Excel

Este artigo descreve como usar a API JavaScript do Excel para desenvolver suplementos para o Excel 2016. Ele apresenta os principais conceitos que são fundamentais para o uso das APIs, como RequestContext, objetos proxy JavaScript, sync(), Excel.run() e load(). Os exemplos de código no final do artigo mostram como aplicar os conceitos.

## RequestContext

O objeto RequestContext possibilita as solicitações para o aplicativo do Excel. Como o suplemento do Office e o aplicativo do Excel são executados em dois processos diferentes, o contexto de solicitação é necessário para obter acesso ao Excel e aos objetos relacionados, como planilhas e tabelas, por meio do suplemento. Um contexto de solicitação é criado conforme mostrado.

```js
var ctx = new Excel.RequestContext();
```

## Objetos proxy

Os objetos JavaScript do Excel declarados e usados em um suplemento são objetos proxy dos objetos reais de um documento do Excel. Todas as ações executadas em objetos proxy não são realizadas no Excel e o estado do documento do Excel não é realizado nos objetos proxy, até que o estado do documento seja sincronizado. O estado do documento é sincronizado quando context.sync() é executado. Confira abaixo.

Por exemplo, o objeto JavaScript local `selectedRange` é declarado para fazer referência ao intervalo selecionado. Você pode usá-lo para colocar a configuração das respectivas propriedades em fila e para invocar métodos. As ações são realizadas sobre esses objetos apenas quando o método sync() é executado.

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

O método sync() disponível no contexto de solicitação sincroniza o estado entre objetos proxy JavaScript e objetos reais no Excel, com a execução de instruções enfileiradas no contexto e com a recuperação de propriedades de objetos carregados do Office para uso no código.  Este método retorna uma promessa, que é resolvida quando o sistema conclui a sincronização.

## Excel.run(function(context) { batch })

O método Excel.run() executa um script em lotes que realiza ações no modelo de objeto do Excel. Os comandos em lotes incluem definições de objetos proxy JavaScript locais e métodos sync() que sincronizam o estado entre objetos locais e do Excel, e a resolução promessa. A vantagem do envio de solicitações em lotes com o método Excel.run() é que, quando a promessa é resolvida, todos os objetos de intervalo controlados que foram alocados durante a execução são automaticamente liberados.

O método de execução é realizado no método RequestContext e retorna uma promessa, geralmente, apenas o resultado de ctx.sync(). É possível executar a operação em lotes fora do Excel.run(). No entanto, todas as referências aos objetos de intervalo devem ser rastreadas e gerenciadas manualmente nesse cenário.

## load()

O método load() é usado para preencher os objetos proxy criados na camada JavaScript do suplemento. Ao tentar recuperar um objeto, por exemplo, uma planilha, um objeto proxy local é criado inicialmente na camada JavaScript. Você pode usar esse objeto para colocar a configuração das respectivas propriedades em fila e para invocar métodos. No entanto, você deve invocar inicialmente os métodos load() e sync() para as relações ou propriedades do objeto de leitura. O método load() realizado nas propriedades e relações que devem ser carregadas quando você chama o método sync().

_Sintaxe:_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
Em que:

* `properties` é a lista de propriedades e/ou nomes de relações a serem carregados, especificados como cadeias de caracteres delimitadas por vírgulas ou por uma matriz de nomes. Confira os métodos load() em cada objeto para saber mais.
* `loadOption` especifica um objeto que descreve as opções de selection, expansion, top e skip. Confira as [opções](../../reference/excel/loadoption.md) de carregamento do objeto para saber mais.

## Exemplo: Gravar valores de uma matriz em um objeto Range

O exemplo a seguir mostra como gravar valores de uma matriz em um objeto Range.

O método Excel.run() inclui um lote de instruções. Como parte deste lote, o sistema cria um objeto proxy que faz referência a um intervalo (A1:B2) na planilha ativa. O valor deste objeto de intervalo proxy é definido localmente. Para ler os valores novamente, a propriedade `text` do intervalo é instruída a ser carregada no objeto proxy. Todos esses comandos estão em fila e são executados quando o método ctx.sync() é chamado. O método sync() retorna uma promessa que pode ser usada para encadeamento com outras operações.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) {

    // Create a proxy object for the sheet
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    // Values to be updated
    var values = [
                 ["Type", "Estimate"],
                 ["Transportation", 1670]
                 ];
    // Create a proxy object for the range
    var range = sheet.getRange("A1:B2");

    // Assign array value to the proxy object's values property.
    range.values = values;

    // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
    return ctx.sync().then(function() {
            console.log("Done");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Exemplo: Copiar valores

O exemplo a seguir mostra como copiar os valores do intervalo A1:A2 a B1:B2 da planilha ativa, usando o método load() no objeto Range.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) {

    // Create a proxy object for the range
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");

    // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
    return ctx.sync().then(function() {
        // Assign the previously loaded values to the new range proxy object. The values will be updated once the following .then() function is invoked.
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = range.values;
    });
}).then(function() {
      console.log("done");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Seleção de propriedades e relações

Por padrão, o método object.load() seleciona todas as propriedades escalares e complexas do objeto que estão sendo carregadas. As relações não são carregadas por padrão; por exemplo, o formato é um objeto de relação do objeto Range. No entanto, convém marcar as propriedades e relações a serem carregadas explicitamente para melhorar o desempenho. Para fazer isso, especifique (no parâmetro `load()`) um subconjunto de propriedades e relações para inclusão na resposta. O método Load permite dois tipos de entradas:

* Nomes de propriedade e de relação como nomes de cadeia de caracteres separados por vírgula _ou_ como uma matriz de cadeias de caracteres que contém os nomes de propriedade ou de relação.
* Um objeto que descreve as opções de propriedade select, expand, top e skip. Confira as [opções](../../reference/excel/loadoption.md) de carregamento do objeto para saber mais.

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### Exemplo

A instrução de carregamento a seguir carrega todas as propriedades do objeto Range e, em seguida, expande o formato e o formato/preenchimento.

```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:B2";
    var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load(["address", "format/*", "format/fill", "entireRow" ]);
    return ctx.sync().then(function() {
        console.log (myRange.address); //ok
        console.log (myRange.format.wrapText); //ok
        console.log (myRange.format.fill.color); //ok
        //console.log (myRange.format.font.color); //not ok as it was not loaded

    });
}).then(function() {
      console.log("done");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

## Entrada com valor nulo

### Entrada com valor nulo em uma matriz 2D

Uma entrada `null` dentro de uma matriz bidimensional (para valores, formato de número ou fórmula) é ignorada na atualização da API. Nenhuma atualização será realizada para o destino pretendido quando a entrada `null` for enviada em valores, formato de número ou grades de fórmulas de valores.

Exemplo: para atualizar somente partes específicas do objeto Range, como alguns formatos de número de células, e para manter o formato de número existente em outras partes do objeto Range, defina o formato de número pretendido onde for necessário e envie `null` para as outras células.

Na solicitação de definição a seguir, somente algumas partes do Formato Numérico do Intervalo são definidas, mantendo ao mesmo tempo o Formato Numérico existente na parte restante (transmitindo valores nulos).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### Entrada nula de uma propriedade

`null` não é uma entrada válida única para toda a propriedade. Por exemplo, o modelo a seguir não é válido, uma vez que os valores inteiros não podem ser ignorados ou definidos como nulos.

```js
 range.values= null;

```

O exemplo a seguir também não é válido, porque nulo não é um valor de cor válido.

```js
 range.format.fill.color =  null;
```

### Resposta nula

Representação de propriedades de formatação que consiste em valores não uniformes que resultariam no retorno de um valor nulo na resposta.

Exemplo: Um intervalo pode consistir de uma ou mais células. Nos casos em que as células individuais incluídas no intervalo especificado não apresentam valores de formatação uniformes, a representação de nível do intervalo será indefinida.

```js
  "size" : null,
  "color" : null,
```

### Entrada e saída em branco

Os valores em branco nas solicitações de atualização são tratados como instrução para limpar ou redefinir a respectiva propriedade. Um valor em branco é representado por aspas duplas sem espaço entre elas. `""`

Exemplo:

* Para `values`, o valor do intervalo é removido. Isso equivale a limpar o conteúdo do aplicativo.

* Para `numberFormat`, o formato de número é definido como `General`.

* Para `formula` e `formulaLocale`, os valores de fórmula são excluídos.


Para as operações de leitura, espera-se receber valores em branco, caso o conteúdo das células esteja em branco. Quando a célula não inclui dados ou valores, a API retorna um valor em branco. Um valor em branco é representado por aspas duplas sem espaço entre elas. `""`.

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## Intervalo sem limite

### Leitura

Um endereço de intervalo não associado contém apenas os identificadores de coluna ou de linha e identificador não especificado de linha ou de coluna, respectivamente, como:

* `C:C`, `A:F`, `A:XFD` (inclui linhas não especificadas)
* `2:2`, `1:4`, `1:1048546` (inclui colunas não especificadas)

Quando a API faz uma solicitação para recuperar um intervalo não associado, por exemplo, `getRange('C:C')`, a resposta retornada inclui `null` para as propriedades no nível da célula, como `values`, `text`, `numberFormat`, `formula`, etc. Outras propriedades de intervalo, como `address`, `cellCount`, etc., refletem o intervalo não associado.

### Gravação

O sistema **não permite** a definição de propriedades de nível da célula (como valores, numberFormat, etc.) em intervalo não associado, uma vez que a solicitação de entrada pode ser muito extensa para ser manipulada.

Exemplo: O exemplo a seguir não é uma solicitação de atualização válida porque o intervalo solicitado não está associado.

```js
...
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
    range.values = 'Due Date';
...
```

Quando uma operação de atualização é tentada nesse intervalo, a API retorna um erro.


## Intervalo longo

Um intervalo longo significa um intervalo cujo tamanho é muito extenso para uma única chamada de API. Muitos fatores, como o número de células, os valores, os formatos de número e as fórmulas incluídas no intervalo, podem fazer com que a resposta seja tão extensa a ponto de se tornar inadequada para interação com a API. A API faz a melhor tentativa para retornar ou gravar os dados solicitados. No entanto, o tamanho extenso envolvido pode resultar em uma condição de erro da API devido à intensa utilização de recursos.

Para evitar isso, convém usar leitura ou gravação para Intervalo longo em vários tamanhos de intervalo menores.


## Cópia de entrada única

Para dar suporte a atualização de um intervalo com os formatos de número ou valores idênticos, ou para a aplicação de uma mesma fórmula em um intervalo, você deve usar a seguinte convenção na API de configuração. No Excel, esse comportamento é semelhante a inserir valores ou fórmulas em um intervalo no modo Ctrl+Enter.

A API vai procurar um *valor de célula única*, no entanto, se a dimensão do intervalo de destino não corresponder à dimensão do intervalo de entrada, ela aplicará a atualização ao intervalo inteiro no modo Ctrl+Enter com o valor ou fórmula fornecida na solicitação.

### Exemplos

A solicitação a seguir atualiza o intervalo escolhido com o texto "Data de conclusão". Observe que o intervalo tem 20 células, ao passo que a entrada fornecida tem apenas 1 valor de célula.

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.values = 'Due Date';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

A solicitação a seguir atualiza o intervalo escolhido com a data de '3/11/2015'.

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.numberFormat = 'm/d/yyyy';
    range.values = '3/11/2015';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
A solicitação a seguir atualiza o intervalo escolhido com uma fórmula que será aplicada em todo o intervalo no modo Ctrl+Enter.

```js
Excel.run(function (ctx) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.numberFormat = 'm/d/yyyy';
    range.values = '3/11/2015';
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


## Mensagens de erro

O sistema retorna erros usando um objeto Error composto por um código e uma mensagem. A tabela a seguir fornece uma lista de possíveis condições de erro que podem ocorrer.

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |O argumento é inválido, está ausente ou tem um formato incorreto.|
|InvalidRequest  |Não é possível processar a solicitação.|
|InvalidReference|Esta referência não é válida para a operação atual.|
|InvalidBinding  |Esta associação de objetos não é mais válida devido às atualizações anteriores.|
|InvalidSelection|A seleção atual é inválida para esta operação.|
|Unauthenticated |Informações de autenticação necessárias estão ausentes ou inválidas.|
|AccessDenied   |Você não pode realizar a operação solicitada.|
|ItemNotFound   |O recurso solicitado não existe.|
|ActivityLimitReached|O limite de atividades foi alcançado.|
|GeneralException|Ocorreu um erro interno ao processar a solicitação.|
|NotImplemented  |O recurso solicitado não foi implementado.|
|ServiceNotAvailable|O serviço não está disponível.|
|Conflito   |A solicitação não pôde ser processada devido a um conflito.|
|ItemAlreadyExists|O recurso que está sendo criado já existe.|
|UnsupportedOperation|Não há suporte para a operação que está sendo tentada.|
|RequestAborted|A solicitação foi anulada durante o tempo de execução.|
|ApiNotAvailable|A API solicitada não está disponível.|
|InsertDeleteConflict|A tentativa de operação de exclusão ou inserção resultou em um conflito.|
|InvalidOperation|A tentativa de operação é inválida no objeto.|

## Recursos adicionais

* [Criar o seu primeiro suplemento do Excel](build-your-first-excel-add-in.md)
* [Explorador de trecho de código](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Exemplos de código de Suplementos do Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Referência da API JavaScript de suplementos do Excel](excel-add-ins-javascript-api-reference.md)
