
# Propriedade TableData.rows
Obtém ou define as linhas na tabela.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Disponível no [Conjunto de requisitos](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Adicionado em**|1.1|

```
var myRows = tableBindingObj.rows;
```


## Valor retornado

Retorna uma matriz de matrizes que contém os dados na tabela. Retornará uma **matriz**`[]` vazia se não houver nenhuma linha.


## Comentários

Para especificar linhas, você deve especificar uma matriz de matrizes que corresponde à estrutura da tabela. Por exemplo, para especificar duas linhas de  valores de **string** em uma tabela com duas colunas, você define a propriedade **rows** como ` [['a', 'b'], ['c', 'd']]`.

Se você especificar **null** para a propriedade **rows** (ou deixar a propriedade vazia quando construir um objeto **TableData**), os resultados a seguir ocorrerão quando o código for executado:


- Se você inserir uma nova tabela, uma linha em branco será inserida.
    
- Se você substituir ou atualizar uma tabela existente, as linhas existentes não serão alteradas.
    

## Exemplo

O exemplo a seguir cria uma tabela de coluna única com um cabeçalho e três linhas.


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}
```


## Detalhes do suporte


Um Y maiúsculo na matriz a seguir indica que esse método tem suporte no aplicativo host do Office correspondente. Uma célula vazia indica que o aplicativo host do Office não dá suporte a esse método.

Para obter mais informações sobre os requisitos de servidor e aplicativo host do Office, consulte [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md) (Requisitos para a execução de Suplementos do Office).


||**Office for Windows desktop**|**Office Online (no navegador)**|**Office para iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|S|Y|S|
|**Word**|S|Y|S|


|||
|:-----|:-----|
|**Disponível nos conjuntos de requisitos**|TableBindings|
|**Nível de permissão mínimo**|[Restrito](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Tipos de suplemento**|Conteúdo, painel de tarefas|
|**Biblioteca**|Office.js|
|**Namespace**|Office|

## Histórico de suporte



****


|**Versão**|**Altera**|
|:-----|:-----|
|1.1|Adicionado suporte para Word Online.|
|1.1|Adicionado suporte para Excel e Word no Office para iPad|
|1.0|Introduzido|
