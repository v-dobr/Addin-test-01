
# Déploiement et publication de votre complément Office


Vous pouvez utiliser plusieurs méthodes pour déployer votre complément Office à des fins de test ou de distribution auprès des utilisateurs :

- [Chargement de version test](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) : utilisez cette méthode dans le cadre de votre processus de développement pour tester l’exécution de votre complément sur Windows, Office Online, iPad ou Mac.
- [Catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) : utilisez cette méthode dans le cadre de votre processus de développement pour tester votre complément ou le distribuer auprès des utilisateurs de votre organisation.
- [Aperçu du Centre d’administration Office 365](https://support.office.com/en-ie/article/Deploy-Office-Add-Ins-in-Office-365-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE) : utilisez cette méthode pour distribuer votre complément auprès des utilisateurs de votre organisation.
- [Office Store] : utilisez cette méthode pour distribuer publiquement votre complément auprès des utilisateurs.

Les options disponibles dépendent de l’hôte Office que vous ciblez et du type de complément.

### Options de déploiement pour les compléments Word, Excel et PowerPoint

| Point d’extension            | Chargement de version test | Catalogue SharePoint | Aperçu du Centre d’administration Office 365 | Office Store |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| Contenu         | X           | X                  | X                               | X            |
| Volet Office       | X           | X                  | X                               | X            |
| Commande         | X           |                    | X                               | X            |

> **REMARQUE :** Les catalogues SharePoint ne sont pas pris en charge dans Office 2016 pour Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à l’[Office Store].    

### Options de déploiement pour les compléments Outlook

| Point d’extension     | Chargement de version test | Serveur Exchange | Office Store |
|:---------|:-----------:|:---------------:|:------------:|
| Application de messagerie | X           | X               | X            |
| Commande  | X           | X               | X            |

Pour élargir la portée de votre complément, assurez-vous qu’il fonctionne sur toutes les plateformes. Les compléments Office sont pris en charge sur Windows, Mac, le web, iOS et Android. Pour avoir une vue d’ensemble des fonctionnalités prises en charge par chaque plateforme, voir la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes].   

Pour plus d’informations sur la gestion des licences pour vos compléments Office Store, consultez la rubrique [Gérer les licences de compléments pour Office et SharePoint](https://msdn.microsoft.com/EN-US/library/office/jj163257.aspx).

Pour plus d’informations sur l’acquisition, l’insertion et l’exécution des compléments par les utilisateurs finals, voir [Commencer à utiliser votre complément Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

## Ressources supplémentaires

- [Disponibilité des compléments Office sur les plateformes et les hôtes]
- [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md) 
- [Soumission de compléments et d’applications web dans l’Office Store][Office Store]
- [Instructions de conception pour les compléments Office](../design/add-in-design)
- [Création de compléments efficaces pour l’Office Store](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)

[Office Store]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Disponibilité des compléments Office sur les plateformes et les hôtes]: http://dev.office.com/add-in-availability
