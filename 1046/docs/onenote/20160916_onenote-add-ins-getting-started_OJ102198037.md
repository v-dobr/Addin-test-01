# Crie seu primeiro suplemento do OneNote

Este artigo ajuda você a criar um suplemento simples de painel de tarefas que adiciona texto a uma página do OneNote.

A imagem a seguir mostra o suplemento que você criará.

   ![O suplemento do OneNote criado a partir deste passo a passo](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## Etapa 1: Configurar seu ambiente de desenvolvimento
1 – Instale o gerador Yeoman Office e seus pré-requisitos seguindo estas [instruções de instalação](https://dev.office.com/docs/add-ins/get-started/create-an-office-add-in-using-any-editor).

   O gerador Yeoman Office facilita a criação de projetos de suplemento quando você não tem o Visual Studio ou deseja usar tecnologias diferentes de JavaScript, CSS e HTML simples. Isto também fornece acesso rápido ao servidor Web local Gulp para teste. 

   >Você pode, opcionalmente, [usar o Visual Studio](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio) para criar os seus arquivos de projeto, mas não serão compatíveis com o servidor interno Gulp.

<a name="create-project"></a>
## Etapa 2: Criar o projeto do suplemento 
1 – Crie uma pasta local denominada *suplemento do onenote*.

2- Abra um prompt **cmd** e navegue até a pasta **suplemento do onenote **. Execute o comando `yo office`, conforme mostrado abaixo.

```
C:\your-local-path\onenote add-in\> yo office
```
>Estas instruções usam o prompt de comando, mas não são igualmente aplicáveis a outros ambientes de shell. 

3- Para criar um projeto, use as opções a seguir.

| Opção | Valor |
|:------|:------|
| Nome do projeto | Suplemento do OneNote |
| Pasta raiz do projeto | (aceitar o padrão) |
| Tipo de projeto do Office | Suplemento do Painel de Tarefas |
| Aplicativos do Office compatíveis | (escolha qualquer um, adicionaremos o host do OneNote posteriormente) |
| Tecnologia a ser usada | HTML, CSS e JavaScript |

<a name="manifest"></a>
## Etapa 3: Configurar o manifesto do suplemento 
1- Abra **manifest-onenote-add-in.xml** nos seus arquivos de projeto. Adicione a linha a seguir à seção **Hosts**. Isto especifica que o suplemento é compatível com o aplicativo de host do OneNote.

```
<Host Name="Notebook" />
```

Observe que **SourceLocation** já está configurado para o servidor Web Gulp.

```
<SourceLocation DefaultValue="https://localhost:8443/app/home/home.html"/>
```

<a name="develop"></a>
## Etapa 4: Desenvolver o suplemento
Você pode desenvolver o suplemento usando um editor de texto ou IDE. Se você não tiver experimentado o Visual Studio Code ainda, você pode [baixá-lo gratuitamente](https://code.visualstudio.com/) no Windows, no Mac OSX e no Linux.

1- Abra **home.html** na pasta *app/home*. 

2- Edite as preferências para o API JavaScript do Office e os estilos e componentes do [Office UI Fabric](http://dev.office.com/fabric).

   a. Retire o comentário do link para fabric.components.min.css.

   b. Substitua a referência de script do Office.js com a seguinte referência para a versão *beta*.

```
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

As suas referências do Office terão a seguinte aparência.

```
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.min.css" rel="stylesheet">
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css" rel="stylesheet">
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

3 – Substitua o elemento `<body>`pelo seguinte código. Isso adiciona uma área de texto e um botão usando [componentes do Office UI Fabric](http://dev.office.com/fabric/components). O layout **Grade Responsiva** vem do conjunto de [estilos do Office UI Fabric](http://dev.office.com/fabric/styles). 

```
<body class="ms-font-m">
    <div class="home flex-container">
        <div class="ms-Grid">
            <div class="ms-Grid-row ms-bgColor-themeDarker">
                <div class="ms-Grid-col">
                    <span class="ms-font-xl ms-fontColor-themeLighter ms-fontWeight-semibold">OneNote Add-in</span>
                </div>
            </div>
        </div>
        <br />
        <div class="ms-Grid">
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <label class="ms-Label">Enter content here</label>
                    <div class="ms-TextField ms-TextField--placeholder">
                        <textarea id="textBox" rows="5"></textarea>
                    </div>
                </div>
            </div>
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <div class="ms-font-m ms-fontColor-themeLight header--text">
                        <button class="ms-Button ms-Button--primary" id="addOutline">
                            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
                            <span class="ms-Button-label">Add outline</span>
                            <span class="ms-Button-description">Adds the content above to the current page.</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
```

4- Abra **home.js** na pasta *app/home*. Edite a função **Office.initialize** para adicionar um evento de clique para o botão **Adicionar estrutura de tópicos**, da seguinte maneira. 

```
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
    $(document).ready(function () {
        app.initialize();

        // Set up event handler for the UI.
        $('#addOutline').click(addOutlineToPage);
    });
};
```
 
5- Substitua o método **getDataFromSelection** pelo seguinte método **addOutlineToPage**. Isto obtém o conteúdo de uma área de texto e o adiciona à página.

```
// Add the contents of the text area to the page.
function addOutlineToPage() {        
    OneNote.run(function (context) {
       var html = '<p>' + $('#textBox').val() + '</p>';

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to load the page with the title property.             
        page.load('title'); 

        // Add an outline with the specified HTML to the page.
        var outline = page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function() {
                console.log('Added outline to page ' + page.title);
            })
            .catch(function(error) {
                app.showNotification("Error: " + error); 
                console.log("Error: " + error); 
                if (error instanceof OfficeExtension.Error) { 
                    console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                } 
            }); 
        });
}
```

<a name="test"></a>
## Etapa 5: Teste o suplemento no OneNote Online
1- Execute o servidor Web Gulp.  

   a. Abra um prompt **cmd** e navegue até a pasta **suplemento do onenote**. 

   b. Execute o comando `gulp serve-static`, conforme mostrado abaixo.

```
C:\your-local-path\onenote add-in\> gulp serve-static
```

2- Instale o certificado autoassinado do servidor Web Gulp como um certificado confiável. Você só precisa fazer isso uma vez no seu computador para projetos de suplemento criados com o gerador Yeoman Office.

   a. Navegue até a página de suplemento hospedada. Por padrão, é a mesma URL que está em seu manifesto:

```
https://localhost:8443/app/home/home.html
```

   b. Instale o certificado como um certificado confiável. Para saber mais, confira [Adicionar certificados autoassinados como certificado raiz de confiança](https://github.com/OfficeDev/generator-office/blob/master/docs/trust-self-signed-cert.md).

3 – No OneNote Online, abra um bloco de anotações.

4 – Escolha **Inserir > Suplementos do Office**. Isso abre a caixa de diálogo Suplementos do Office.
  - Se você estiver conectado com a sua conta de consumidor, escolha a guia **MEUS SUPLEMENTOS** e, em seguida, escolha  **Carregar Meu Suplemento**.
  - Se você estiver conectado com a sua conta corporativa ou de estudante, escolha a guia **MINHA ORGANIZAÇÃO** e, em seguida, escolha **Carregar Meu Suplemento**. 
  
  A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anotações do consumidor.

  ![A caixa de diálogo Suplementos do Office mostrando a guia MEUS SUPLEMENTOS](../../images/onenote-office-add-ins-dialog.png)
  
  >**OBSERVAÇÃO**: Para habilitar o botão **Suplementos do Office**, clique dentro da página do OneNote.

5 – Na caixa de diálogo Carregar suplemento, navegue até **manifest-onenote-add-in.xml** nos arquivos de projeto e, em seguida, escolha **Carregar**. Ao testar, o seu arquivo de manifesto pode ser armazenado localmente.

6 – O suplemento abre em um iFrame perto da página do OneNote. Insira algum texto na área correspondente e escolha **Adicionar estrutura de tópicos**. O texto inserido é adiciona à pagina. 

## Dicas e solução de problemas
- Você pode depurar o suplemento usando as ferramentas de desenvolvedor do seu navegador. Quando você estiver usando o servidor Web Gulp e depurando no Internet Explore ou no Chrome, você pode salvar as alterações localmente e apenas atualize o iFrame do suplemento.

- Quando você inspecionar um objeto do OneNote, as propriedades que estão atualmente disponíveis usam valores reais de exibição. As propriedades que precisam ser carregadas exibem *undefined*. Expanda o nó `_proto_` para ver as propriedades definidas no objeto, mas que ainda não foram carregadas.

      ![Unloaded OneNote object in the debugger](../../images/onenote-debug.png)

- Você precisa habilitar conteúdo misto no navegador, se o seu suplemento usar todos os recursos HTTP. Os suplementos de produção devem usar apenas recursos HTTPS seguros.

## Recursos adicionais

- [Visão geral da programação da API JavaScript do OneNote](onenote-add-ins-programming-overview.md)
- [Referência da API JavaScript do OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma Suplementos do Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
