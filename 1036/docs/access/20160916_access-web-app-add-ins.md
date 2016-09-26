
# Création de compléments pour les applications web Access



Cet article explique comment utiliser Visual Studio 2015 pour développer un Complément Office qui cible les applications web Access.

>
  **Remarque :** pour plus d’informations sur le développement de solutions pour Access à l’aide de VBA, consultez la rubrique [Access](https://msdn.microsoft.com/en-us/library/fp179695.aspx) sur MSDN.

## Conditions préalables

Pour créer une Complément Office qui cible applications web Access, vous avez besoin des éléments suivants :


- Visual Studio 2015

- Un site SharePoint Online (inclus dans de nombreux abonnements Office 365). Ce site doit disposer d’un catalogue de compléments. Pour plus d’informations, voir [Configurer un catalogue de compléments sur SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


 >**Remarque**   Les Compléments Office fonctionnent avec l’applications web Access hébergée sur SharePoint Online ou Office 365. L’application de bureau Access 2013 ne prend pas en charge les Compléments Office. Les Compléments Office qui ciblent l’applications web Access sont pris en charge par la version 1.1 et les versions ultérieures de Microsoft Office.js.


## Créer un projet dans Visual Studio


1.  Ouvrez Visual Studio et, dans le menu, choisissez **Fichier**, **Nouveau**, **Projet**. La boîte de dialogue **Nouveau projet** s’ouvre.

2. Dans la boîte de dialogue **Nouveau projet**, dans le volet de gauche, accédez à  **Installé**,  **Modèles**,  **Visual C#**,  **Office/SharePoint**,  **Compléments Office**.

3. Dans la boîte de dialogue **Nouveau projet**, dans le volet central, choisissez  **Complément Office**.

4. Au bas de la boîte de dialogue, saisissez le nom de votre projet et choisissez  **OK**. Ceci ouvre la boîte de dialogue  **Créer un complément Office**.

5. Dans la boîte de dialogue **Créer un complément Office**, choisissez  **Contenu** et cliquez sur **Suivant**.

6. Dans l’écran suivant de la boîte de dialogue  **Créer un complément Office**, choisissez  **Complément de base** ou **Complément de visualisation de document**, et assurez-vous que la case pour  **Access** est cochée.

7. Lorsque vous avez terminé, choisissez  **Terminer**. Visual Studio créera un projet de démarrage sur lequel baser votre travail.

8. Dans **l’Explorateur de solutions**, choisissez un projet web du projet (**nom_projet > Web**). Dans le volet de propriétés, recherchez l’entrée pour l’**URL SSL**. Cela doit avoir la forme suivante : `https://localhost:44314/`. Sélectionnez cette URL et copiez-la dans le Presse-papiers. Vous en aurez besoin dans peu de temps.

9. Cliquez avec le bouton droit sur le nom de votre projet dans l’**Explorateur de solutions**. Dans le menu contextuel, choisissez **Publier**. Ceci ouvre l’Assistant **Publier votre complément**.

10. Dans l’Assistant **Publier votre complément**, sélectionnez la liste déroulante en regard de **Profil actuel**. Dans cette liste déroulante, choisissez **nouveau**. Ceci ouvre la boîte de dialogue **Publier des compléments Office et SharePoint**.

11. Dans cette boîte de dialogue, choisissez  **Créer un profil**, saisissez un nom reconnaissable pour le profil, puis choisissez  **Terminer**. La boîte de dialogue  **Publier des compléments Office et SharePoint** se ferme, vous renvoyant à l’Assistant **Publier votre complément**.

12. Dans l’Assistant, choisissez  **Empaqueter le complément**. Ceci finalise votre complément pour qu’il puisse être publié dans un catalogue de compléments dans SharePoint.

13. Dans la page suivante, pour **Où votre site web est-il hébergé ?**, indiquez l’URL de l’hôte de votre site web. Il peut s’agir de la valeur  **URL SSL** copiée à l’étape 8. Ensuite, cliquez sur **Terminer**.

14. Dans l’ **Explorateur de solutions**, cliquez avec le bouton droit sur le nœud du manifeste du projet (immédiatement sous le nom du projet) et sélectionnez  **Ouvrir le dossier dans l’Explorateur de fichiers**. Notez le chemin de ce fichier. Vous aurez besoin de cette information ultérieurement.


 >**Remarque**  Pour déboguer le complément, vous devez le déployer avec une application web Access.


## Passer en revue le manifeste et le fichier Home.Html


1. Dans votre projet Visual Studio, ouvrez le fichier  **Home.html** et trouvez les lignes qui font référence à la bibliothèque de scripts d’Office.js.

```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```
 >**Notez** que la balise script fait référence à la version 1.1 (et aux versions ultérieures) d’Office.js. Access utilise des éléments d’API introduits dans la version 1.1.

2. Ouvrez le fichier manifeste associé à votre projet. Ce fichier sera nommé d’après le nom de votre projet, et aura l’extension « .xml ».

3.  Dans le fichier manifeste, trouvez la section **Hôtes** et recherchez une entrée **Hôte**.

```xml
  <Hosts> <Host Name="Database" /> </Hosts>
```
 >**Remarque** Les applications qui peuvent utiliser le complément sont répertoriées à cet endroit. Comme vous avez sélectionné **Access** dans la boîte de dialogue **Créer un complément Office**, l’élément **Base de données** est répertorié. Si vous avez inclus Excel, il existe également une entrée pour **Classeur**.

Les Compléments Office et SharePoint sont basées sur le web. Le code du complément doit être hébergé sur un serveur web. Pour cet exemple, le serveur web est votre ordinateur de développement. Le serveur doit être en cours d’exécution pour que le complément s’en serve pour les tests, ce qui signifie que Visual Studio doit exécuter le complément au moment où vous le visualisez et le déboguez dans SharePoint.

Pour qu’un utilisateur trouve et utilise le complément, celui-ci doit être inscrit auprès d’un catalogue de compléments dans SharePoint. Les informations dont le catalogue de compléments a besoin sont contenues dans le fichier manifeste.

 >**Remarque**  Vous aurez besoin de créer une application web Access pour héberger votre Complément Office.


## Publier votre complément dans un catalogue SharePoint Online


1.  Connectez-vous à SharePoint Online ou Office 365, puis accédez au **centre d’administration SharePoint** en choisissant **Admin** dans la barre d’outils Office 365 en haut de la page.

2. Dans le  **centre d’administration SharePoint**, accédez à la barre de liens de gauche et choisissez  **compléments**. Ceci vous dirigera vers la vue des compléments.

3. Dans le volet central de la page, choisissez  **Catalogue de compléments**. Ceci vous conduit à la page  **Catalogue**.

4. Dans la page  **Catalogue**, choisissez  **Distribuer les compléments Office**. Ceci vous conduit à une page d’annuaire appelée  **Compléments Office** qui répertorie toutes les Compléments Office installées.

5. En haut de la page  **Compléments Office**, choisissez  **Nouveau complément**. Ceci permet d’afficher la boîte de dialogue **Ajouter un document**.

6. Dans la boîte de dialogue  **Ajouter un document**, choisissez  **Parcourir**, puis allez à l’emplacement du fichier manifeste dans votre projet Visual Studio. Si vous avez copié l’adresse de votre fichier manifeste précédemment, vous pouvez la coller dans cette boîte de dialogue.

7. Choisissez le fichier manifeste dans votre projet, puis cliquez sur  **OK**. SharePoint ajoute désormais votre complément à la bibliothèque SharePoint locale.


 >**Remarque**  Cette procédure suppose que vous avez créé un site de test sur SharePoint. Dans le cas contraire, vous pouvez le faire à partir de l’onglet  **Sites** en haut de la fenêtre de SharePoint. Vous pouvez utiliser une applications web Access si vous en disposez d’une.


## Créer une application web Access pour héberger votre complément


1. Accédez à votre site de test. Dans la barre de liens, choisissez  **Contenu du site**. Vous accédez alors à la page  **Contenu du site** de votre site de test.

2. Sur la page  **Contenu du site**, choisissez  **Ajouter un complément**. Vous accédez alors à la page  **Contenu du site – Vos compléments**.

3. Sur la page  **Contenu du site – Vos compléments**, utilisez la barre de recherche en haut de la page pour rechercher  **Application**.

4. Vous devez maintenant voir une mosaïque pour **Application**.

     >**Remarque**  Gardez à l’esprit qu’il ne s’agit pas de votre complément Office, mais de nouvelles applications web Access. Ces applications web Access vont héberger votre complément Office.
5. Le choix de cette mosaïque affiche la boîte de dialogue  **L’ajout d’une application Access est en cours... Merci de patienter.**. Saisissez un nom unique pour votre applicationAccess et choisissez  **Créer**. La création de votre application peut prendre un certain temps dans SharePoint. Une fois l’opération terminée, votre applicationAccess est répertoriée sur la page  **Contenu du site** avec une étiquette **nouvelle**.

6. Vous devez ouvrir l’applicationAccess dans la version de bureau de Microsoft Access 2013 et y ajouter des données avant de l’ouvrir et de l’afficher dans SharePoint.


## Ajouter votre complément à une applications web Access


1. Ouvrez une applications web Access.

2. Dans la barre d’onglets de SharePoint, choisissez l’icône d’engrenage dans l’angle supérieur gauche. Un menu s’affiche. Choisissez l’élément de menu  **Compléments Office**. Cette action ouvre la boîte de dialogue  **Compléments Office**.

3. Choisissez la vue  **Mon organisation** et patientez pendant que SharePoint remplit la boîte de dialogue avec les Compléments Office dont vous disposez.

    L’un des compléments de la boîte de dialogue doit correspondre au complément Office que vous avez enregistré dans une procédure antérieure. Choisissez ce complément et insérez-le dans vos applications web Access. N’oubliez pas que l’application doit s’exécuter dans Visual Studio pour qu’il soit détecté et qu’il s’affiche sur votre page d’applications web Access.


## Débogage du complément pour Office

Pour déboguer votre complément, dans Internet Explorer, appuyez sur F12 ou cliquez sur l’icône d’engrenage dans la barre d’onglets des navigateurs (pas l’icône d’engrenage sur la page SharePoint). Ceci affiche les outils de débogage F12 fournis par Internet Explorer 11. Si vous utilisez un autre navigateur, consultez la documentation de votre navigateur pour savoir comment entrer en mode de débogage.

À ce stade, vous pouvez définir des points d’arrêt, parcourir votre code JavaScript, explorer les DOM et modifier le code pour vérifier que vos modifications apparaissent dans l’Complément Office ciblant applications web Access. Pour plus d’informations, voir [Utilisation des outils de développement F12](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29).


## Étapes suivantes

Téléchargez l’exemple sur la page [Office 365 : Liaison et manipulation des données dans une application web Access](https://code.msdn.microsoft.com/officeapps/Office-365-Bind-and-4876274e) pour savoir comment implémenter un Complément Office qui manipule des données dans une application web Access.


## Ressources supplémentaires



- [Présentation de l’API JavaScript pour compléments](../develop/understanding-the-javascript-api-for-office.md)

- [Interface API JavaScript pour Office](../../reference/javascript-api-for-office.md)

