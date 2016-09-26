
# Depurar suplementos do Office no iPad e no Mac

Você pode usar o Visual Studio para desenvolver e depurar suplementos no Windows, mas não pode usá-lo para depurar suplementos no iPad ou no Mac. Como os suplementos são desenvolvidos usando HTML e Javascript, são projetados para funcionar em várias plataformas, mas pode haver diferenças sutis em como cada navegador processa o HTML. Este artigo descreve como depurar suplementos em execução em um iPad ou em um Mac. 

## Depuração com Vorlon.js 

Vorlon.js é um depurador para páginas da Web, semelhante às ferramentas F12, que foi projetado para funcionar remotamente e permite depurar páginas da Web em dispositivos diferentes. Para saber mais, confira o [site do Vorlon](http://www.vorlonjs.com).  

Para instalar e configurar o Vorlon: 

1.  Instale [Node.js](https://nodejs.org) e [Git](https://git-scm.com/) se ainda não tiver feito isso. 

2.  Instale o Vorlon usando git com o seguinte comando: `git clone https://github.com/MicrosoftDX/Vorlonjs.git`

3.  Instale dependências com o `npm install`.

4.  Suplementos exigem HTTPS e, portanto, por extensão, todos os scripts que eles usarem deverão ser HTTPS também, incluindo o script Vorlon. Portanto, você precisará configurar Vorlon para usar SSL se quiser usar esse script com suplementos. Na pasta em que você instalou Vorlon, acesse a pasta /Server e edite o arquivo config.json. Mude a propriedade **useSSL** para **true**. Enquanto estiver nesse local, você também pode habilitar o plug-in para Suplementos do Office (mude a propriedade "enabled" para true). 

5.  Execute o servidor Vorlon com o comando `sudo vorlon`. 

6.  Abra uma janela do navegador e vá para [http://localhost:1337](http://localhost:1337), que é a interface do Vorlon. Confie no certificado de segurança. Você será solicitado a fazer isso. Você também pode encontrar o certificado de segurança na pasta Vorlon, em /Server/cert. 

7.  Adicione a seguinte marca de script à seção `<head>` do arquivo home.html (ou arquivo HTML principal) do seu suplemento:
```    
<script src="https://localhost:1337/vorlon.js"></script>    
```  

Agora, sempre que abrir o suplemento em um dispositivo, ele aparecerá na lista de clientes no Vorlon (no lado esquerdo da interface do Vorlon). Você pode realçar elementos DOM remotamente, executar comandos remotamente e muito mais.  

![Captura de tela que mostra a interface do Vorlon.js](../../images/vorlon_interface.png)

O plugin do Office adiciona recursos extras para Office.js, como explorar o modelo de objeto e executar chamadas de Office.js. 
