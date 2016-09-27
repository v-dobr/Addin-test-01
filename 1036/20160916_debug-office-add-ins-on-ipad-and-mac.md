
# Débogage des compléments Office sur iPad et Mac

Vous pouvez utiliser Visual Studio pour le développement et le débogage des compléments sur Windows. Toutefois, vous ne pouvez pas l’utiliser pour déboguer les compléments sur iPad ou sur Mac. Dans la mesure où les compléments sont développés dans le code HTML et Javascript, ils devraient fonctionner sur différentes plateformes. Il peut toutefois exister de légères différences dans l’affichage du code HTML dans les différents navigateurs. Cette rubrique explique comment déboguer les compléments en exécution sur iPad ou sur Mac. 

## Débogage avec Vorlon.js 

Vorlon.js est un débogueur de pages web, semblable aux outils F12, conçu pour fonctionner à distance et pour vous permettre de déboguer des pages web sur différents appareils. Pour plus d’informations, accédez au [site web de Vorlon](http://www.vorlonjs.com).  

Pour installer et configurer Vorlon : 

1.  Installez [Node.js](https://nodejs.org) et [Git](https://git-scm.com/) si ce n’est pas déjà fait. 

2.  Installez Vorlon à l’aide de Git avec la commande suivante : `git clone https://github.com/MicrosoftDX/Vorlonjs.git`.

3.  Installez des dépendances avec `npm install`.

4.  Les compléments nécessitent HTTPS. Ainsi, par extension, les scripts qu’ils utilisent doivent également être HTTPS, y compris le script Vorlon. Par conséquent, vous devez configurer Vorlon de manière à ce qu’il utilise SSL afin de pouvoir utiliser Vorlon avec des compléments. Sous le dossier où vous avez installé Vorlon, accédez au dossier /Server et modifiez le fichier config.json. Définissez la propriété **useSSL** sur **True**. Profitez-en pour également activer le plug-in pour les compléments Office (définissez sa propriété « enabled » sur True). 

5.  Exécutez le serveur Vorlon avec la commande `sudo vorlon`. 

6.  Ouvrez une fenêtre de navigateur et accédez à [http://localhost:1337](http://localhost:1337), qui correspond à l’interface Vorlon. Approuvez le certificat de sécurité lorsque vous y serez invité. Vous pouvez également trouver le certificat de sécurité dans le dossier Vorlon sous /Server/cert. 

7.  Ajoutez la balise de script suivante à la section `<head>` du fichier home.html (ou fichier HTML principal) de votre complément :
```    
<script src="https://localhost:1337/vorlon.js"></script>    
```  

Désormais, chaque fois que vous ouvrez le complément sur un appareil, il apparaît dans la liste Clients dans Vorlon (sur le côté gauche de l’interface Vorlon). Vous pouvez surligner à distance des éléments DOM, exécuter à distance des commandes et bien plus encore.  

![Capture d’écran de l’interface Vorlon.js](../../images/vorlon_interface.png)

Le plug-in Office ajoute des fonctionnalités supplémentaires pour Office.js, telles que l’exploration du modèle objet et l’exécution d’appels Office.js. 
