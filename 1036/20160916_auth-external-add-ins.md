# Autoriser des services externes dans votre complément Office

Les services en ligne populaires, y compris Office 365, Google, Facebook, LinkedIn, SalesForce et GitHub, permettent aux développeurs d’accorder aux utilisateurs l’accès à leurs comptes dans d’autres applications. Vous avez ainsi la possibilité d’inclure ces services dans votre complément Office. 

L’infrastructure standard dans le secteur permettant d’activer l’accès d’une application web à un service en ligne est appelée OAuth 2.0. En règle générale, vous n’avez pas besoin de connaître les détails du fonctionnement de l’infrastructure pour pouvoir l’utiliser dans votre complément. Ces détails sont résumés pour vous dans de nombreuses bibliothèques disponibles.

L’un des fondements d’OAuth est qu’une application peut être un principal de sécurité en elle-même, de la même façon qu’un utilisateur ou un groupe, avec sa propre identité et son ensemble d’autorisations. Dans les scénarios les plus courants, lorsque l’utilisateur exécute une action dans le complément Office ayant besoin du service en ligne, le complément envoie une demande au service portant sur un ensemble spécifique d’autorisations pour le compte de l’utilisateur. Le service invite ensuite l’utilisateur à octroyer ces autorisations au complément. Une fois que les autorisations sont accordées, le service envoie un petit *jeton d’accès* codé au complément. Le complément peut utiliser le service en incluant le jeton dans toutes ses demandes aux API du service. Toutefois, le complément agit uniquement dans la limite des autorisations que l’utilisateur lui a accordées. En outre, le jeton expire après un certain délai.

Plusieurs modèles OAuth, appelés *flux* ou *types d’accès accordé*, sont conçus pour différents scénarios. Les deux principaux sont les suivants :

- **Flux implicite** : la communication entre le complément et le service en ligne est mise en œuvre avec JavaScript côté client.
- **Flux de code d’autorisation** : la communication est effectuée de *serveur à serveur* entre l’application web de votre complément et le service en ligne. Par conséquent, elle est mise en œuvre avec du code côté serveur.

L’objectif des flux consiste à sécuriser l’identité et l’autorisation de l’application. Dans le flux de code d’autorisation, une *clé secrète client* devant rester masquée vous est fournie. Comme une application monopage (SPA) ne permet pas de protéger la clé secrète, nous vous recommandons d’utiliser le flux implicite dans ce type d’application. 

Vous devez connaître les autres avantages et inconvénients des deux flux. Les définitions officielles de [Code d’autorisation](https://tools.ietf.org/html/rfc6749#section-1.3.1) et d’[Implicite](https://tools.ietf.org/html/rfc6749#section-1.3.2) constituent un excellent point de départ. 

>**Remarque :** vous avez aussi la possibilité de charger un service intermédiaire d’effectuer tout ce qui concerne les autorisations à votre place et de transmettre le jeton d’accès à votre complément. Pour plus d’informations, reportez-vous la section *Services intermédiaires* plus loin dans cet article.

## Utilisation du flux implicite dans des compléments Office
La meilleure façon de déterminer si le service en ligne prend en charge le flux implicite est de consulter la documentation.

Pour les services qui le prennent en charge, nous fournissons une bibliothèque JavaScript qui effectue tout le travail de détail à votre place :

[Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

Le dossier \demo du référentiel contient un exemple de complément qui utilise la bibliothèque pour accéder à certains services populaires, y compris Google, Facebook et Office 365.

Reportez-vous également à la section **Bibliothèques** plus loin dans cet article.

## Utilisation du flux de code d’autorisation dans les compléments Office

Les exemples de complément suivants utilisent le flux de code d’autorisation :

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

De nombreuses bibliothèques sont disponibles pour l’implémentation du flux de code d’autorisation dans différentes langues et infrastructures. Pour plus d’informations, reportez-vous à la section **Bibliothèques** plus loin dans cet article.

### Fonctions de relais/proxy

Vous pouvez utiliser le flux de code d’autorisation même avec une application web sans serveur en stockant les valeurs d’*ID client* et de *clé secrète client* dans une fonction simple, hébergée dans un service tel qu’[Azure Functions](https://azure.microsoft.com/en-us/services/functions) ou [Amazon Lambda](https://aws.amazon.com/lambda).
La fonction remplace un code donné par un *jeton d’accès* approprié et le transmet au client. La sécurité de cette approche dépend de la surveillance de l’accès à la fonction.

Pour utiliser cette technique, votre complément ouvre une interface utilisateur/un menu contextuel pour afficher l’écran de connexion au service en ligne (Google, Facebook, etc.). Lorsque l’utilisateur est connecté et accorde l’autorisation au complément d’accéder à ses ressources dans le service en ligne, le développeur reçoit un code qui peut être envoyé à la fonction en ligne. Les services décrits dans la section **Services intermédiaires** de cet article utilisent un flux semblable à celui-ci. 

## Bibliothèques

Les bibliothèques sont disponibles pour de nombreuses langues et plateformes, ainsi que pour les deux flux. Certaines d’entre elles sont destinées à un usage général, tandis que d’autres sont propres à des services en ligne spécifiques. 

**Office 365 et autres services utilisant Azure Active Directory en tant que fournisseur d’autorisation** : [bibliothèques d’authentification Azure Active Directory](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/). Un aperçu est également disponible pour la [bibliothèque d’authentification Microsoft](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google** : cherchez « auth » ou le nom de votre langue sur [GitHub.com/Google](https://github.com/google). La plupart des référentiels pertinents sont nommés `google-auth-library-[name of language]`.

**Facebook** : cherchez « bibliothèque » ou « sdk » sur le site [Facebook pour les développeurs](https://developers.facebook.com). 

**OAuth 2.0 général** : une page contenant des liens vers des bibliothèques pour plus d’une dizaine de langues est conservée par le groupe de travail OAuth de l’IETF sur une page relative au [code OAuth](http://oauth.net/code/). Notez que certaines de ces bibliothèques sont destinées à l’implémentation d’un service compatible OAuth. Les bibliothèques qui vous sont utiles en tant que développeur de compléments sont appelées bibliothèques *client* sur cette page car votre serveur web est un client du service compatible OAuth.

## Services intermédiaires

Votre complément peut utiliser un service intermédiaire tel qu’Auth0 qui fournit des jetons d’accès pour de nombreux services en ligne populaires ou qui simplifie la procédure de connexion aux réseaux sociaux pour votre complément, ou qui effectue ces deux opérations. Avec très peu de code, votre complément peut utiliser un script côté client ou du code côté serveur pour se connecter au service intermédiaire et renvoyer les jetons requis pour le service en ligne. L’ensemble du code de mise en œuvre des autorisations se trouve dans le service intermédiaire. 

L’exemple suivant utilise Auth0 pour activer la connexion aux réseaux sociaux avec Facebook, Google et les comptes Microsoft :

[Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)

## Que signifie l’acronyme CORS ?

CORS est l’acronyme de [Cross Origin Resource Sharing](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS) (partage des ressources d’origines croisées). Pour plus d’informations sur l’utilisation de CORS dans les compléments, reportez-vous à la rubrique relative à la [résolution des limites de stratégie d’origine identique dans les compléments Office](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations).
