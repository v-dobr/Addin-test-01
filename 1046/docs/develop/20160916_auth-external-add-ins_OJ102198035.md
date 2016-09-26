# Autorizar serviços externos no seu suplemento do Office

Serviços online populares, incluindo o Office 365, o Google, o Facebook, o LinkedIn, o SalesForce e o GitHub, permitem que os desenvolvedores forneçam acesso para os usuários a suas contas em outros aplicativos. Isso dá a você a capacidade de incluir esses serviços no seu Suplemento do Office. 

A estrutura padrão do setor para habilitar o acesso de aplicativos Web a um serviço online é chamada de OAuth 2.0. Na maioria das situações, você não precisa saber os detalhes de como a estrutura funciona para usá-la no seu suplemento. Estão disponíveis muitas bibliotecas que abstraem os detalhes para você.

Uma ideia fundamental do OAuth é que um aplicativo pode ser uma entidade de segurança por si só, assim como um usuário ou um grupo, com sua própria identidade e conjunto de permissões. Nos cenários mais comuns, quando o usuário realiza uma ação no suplemento do Office que requer o serviço online, o suplemento envia ao serviço uma solicitação para um conjunto específico de permissões para a conta do usuário. Em seguida, o serviço solicita que o usuário conceda essas permissões ao suplemento. Após a concessão das permissões, o serviço envia ao suplemento um pequeno *token de acesso* codificado. O suplemento pode usar o serviço, incluindo o token, em todas as suas solicitações para as APIs do serviço. Porém, o suplemento só pode agir dentro das permissões concedidas a ele pelo usuário. O token também expira após um tempo especificado.

Vários padrões OAuth, chamados de *fluxos* ou *tipos de concessão*, foram projetados para diferentes cenários. Veja a seguir os dois mais importantes:

- **Fluxo Implícito**: A comunicação entre o suplemento e o serviço online é implementada com um JavaScript no lado do cliente.
- **Fluxo de Código de Autorização**: A comunicação é *de servidor para servidor* entre o aplicativo Web do seu suplemento e o serviço online. Portanto, a implementação é feita com código no lado do servidor.

A finalidade dos fluxos é proteger a identidade e a autorização do aplicativo. No fluxo de Código de Autorização, você recebe um *segredo de cliente* que precisa permanecer oculto. Como um Aplicativo de Página Única (SPA) não tem como proteger o segredo, nós recomendamos que você use o fluxo Implícito em SPAs. 

Você deve estar familiarizado com os outros prós e contras dos dois fluxos. As definições oficiais em [Código de Autorização](https://tools.ietf.org/html/rfc6749#section-1.3.1) e [Implícito](https://tools.ietf.org/html/rfc6749#section-1.3.2) são um bom ponto de partida. 

>**Observação:** Também existe a opção de designar um serviço intermediário para fazer toda a autorização para você e transmitir o token de acesso ao seu suplemento. Para obter detalhes, confira a seção *Serviços intermediários* mais adiante neste artigo.

## Usando o fluxo Implícito em suplementos do Office
A melhor maneira de descobrir se o serviço online dá suporte ao fluxo Implícito é consultar a documentação.

Para os serviços com suporte, fornecemos uma biblioteca JavaScript que faz todo o trabalho detalhado para você:

[Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

A pasta \demo do repositório contém um suplemento de exemplo que usa a biblioteca para acessar alguns serviços populares, entre eles o Google, o Facebook e o Office 365.

Consulte também a seção **Bibliotecas** mais adiante neste artigo.

## Usando o fluxo de Código de Autorização em suplementos do Office

Temos alguns suplementos de exemplo que usam o fluxo de Código de Autorização:

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

Muitas bibliotecas estão disponíveis para implementar o fluxo de Código de Autorização em várias linguagens e estruturas. Para obter detalhes, consulte a seção **Bibliotecas** mais adiante neste artigo.

### Funções de Retransmissão/Proxy

Você pode usar o fluxo de Código de Autorização até mesmo com um aplicativo Web sem servidor armazenando os valores de *ID do cliente* e *segredo cliente* em uma função simples que está hospedada em um serviço como [Funções do Azure](https://azure.microsoft.com/en-us/services/functions) ou o [Amazon Lambda](https://aws.amazon.com/lambda).
A função troca um código específico por um *token de acesso* apropriado e o transmite de volta para o cliente. A segurança dessa abordagem depende de quão bem o acesso à função é protegido.

Para usar essa técnica, o suplemento exibe uma interface do usuário/pop-up para mostrar a tela de logon do serviço online (Google, Facebook e assim por diante). Quando o usuário faz logon e concede permissão ao suplemento para seus recursos no serviço online, o desenvolvedor recebe um código que então pode ser enviado para a função online. Os serviços descritos em **Serviços intermediários** neste artigo usam um fluxo semelhante a esse. 

## Bibliotecas

Bibliotecas estão disponíveis para várias linguagens e plataformas, bem como para ambos os fluxos. Algumas são de uso geral, enquanto outras são para serviços online específicos. 

**Office 365 e outros serviços que usam o Azure Active Directory como provedor de autorização**: [Bibliotecas de autenticação do Azure Active Directory](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/). Também está disponível uma prévia da [Biblioteca de Autenticação da Microsoft](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google**: Pesquise "auth" ou o nome da sua linguagem no [GitHub.com/Google](https://github.com/google). A maioria dos repositórios relevantes se chama `google-auth-library-[name of language]`.

**Facebook**: Pesquise "library" ou "sdk" no [Facebook para Desenvolvedores](https://developers.facebook.com). 

**OAuth 2.0 Geral**: Uma página de links para bibliotecas de mais de uma dúzia de linguagens é mantida pelo IETF OAuth Working Group, em: [Código OAuth](http://oauth.net/code/). Observe que algumas dessas bibliotecas são para implementar um serviço compatível com OAuth. As bibliotecas que são interessantes para você como desenvolvedor se chamadas de bibliotecas de *cliente* nessa página, pois o seu servidor Web é um cliente do serviço compatível com OAuth.

## Serviços intermediários

Seu suplemento pode usar um serviço intermediário, como o Auth0, que fornece tokens de acesso para muitos serviços online populares e/ou simplifica o processo de habilitar o logon social para esse suplemento. Com muito pouco código, o suplemento pode usar qualquer script no lado do cliente ou código no lado do servidor para se conectar ao intermediário e retornar qualquer token necessário para o serviço online. Todo o código de implementação de autorização está no serviço intermediário. 

Temos um exemplo que usa Auth0 para habilitar o logon social com o Facebook, o Google e Contas da Microsoft:

[Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)

## O que é CORS?

CORS significa [Compartilhamento de recursos entre origens](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS). Para obter informações sobre como você pode usar o trabalho com o CORS em suplementos, consulte [Lidando com limitações de políticas de mesma origem em suplementos do Office](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations).
