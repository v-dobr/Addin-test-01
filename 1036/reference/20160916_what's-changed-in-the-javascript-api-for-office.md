
# Nouveautés de l’API JavaScript pour Office
Afin d’étendre la fonctionnalité de vos Compléments Office, des objets, méthodes, propriétés, événements et énumérations sont régulièrement ajoutés et mis à jour dans l’API JavaScript pour Office. Utilisez les liens ci-dessous pour afficher les membres de l’API qui ont été ajoutés ou mis à jour.

Pour développer des compléments utilisant les nouveaux membres de l’API, vous devez [mettre à jour l’API JavaScript pour les fichiers de l’API JavaScript pour Office dans votre projet](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

Pour visualiser tous les membres de l’API, y compris ceux qui sont identiques par rapport aux versions précédentes, voir [API JavaScript pour Office](../reference/javascript-api-for-office.md).


## API ajoutées et mises à jour

 **Objets ajoutés et mis à jour**


|**Objet**|**Description**|**Version ajoutée ou mise à jour**|
|:-----|:-----|:-----|
|[Élément](../reference/outlook/Office.context.mailbox.item.md)|Mises à jour et ajouts pour les éléments suivants :<br><ul><li><p>Méthodes <a href="../reference/outlook/Office.context.mailbox.item.md#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> et <a href="../reference/outlook/Office.context.mailbox.item.md#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> pour prendre en charge l’obtention de la sélection de l’utilisateur et son remplacement dans l’objet et le corps d’un message ou d’un rendez-vous.</p></li><li><p>Méthodes <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyAllForm" target="_blank">displayReplyAllForm</a> et <a href="../reference/outlook/Office.context.mailbox.item.md#displayReplyForm" target="_blank">displayReplyForm</a> pour prendre en charge l’ajout d’une pièce jointe au formulaire de réponse d’un rendez-vous.</p></li></ul>|Boîte aux lettres 1.2|
|[Élément](../reference/outlook/Office.context.mailbox.item.md)|Mis à jour pour inclure des méthodes et des champs utiles à la création de compléments Outlook en mode composition. |1.1|
|[Liaison](../reference/shared/binding.md)|Mis à jour pour prendre en charge la liaison de tableau dans les compléments de contenu pour Access.|1.1|
|[Bindings](../reference/shared/bindings.bindings.md)|Mis à jour pour prendre en charge la liaison de tableau dans les compléments de contenu pour Access.|1.1|
|[Body](../reference/outlook/Body.md)|Ajouté pour permettre la création et la modification du corps d’un message ou d’un rendez-vous dans les compléments Outlook en mode composition.|1.1|
|[Document](../reference/shared/document.md)|Mis à jour et ajouté pour : <ul><li><p>Prendre en charge les propriétés <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>, <a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a> et <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> dans les compléments de contenu pour Access.</p></li><li><p>Obtenir le document au format PDF à l’aide de la méthode <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> dans les compléments pour PowerPoint et Word.</p></li><li><p>Obtenir les propriétés de fichier à l’aide de la méthode <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> dans les compléments pour Excel, PowerPoint et Word.</p></li><li><p>Accéder aux emplacements et aux objets au sein du document à l’aide de la méthode <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> dans les compléments pour Excel et PowerPoint.</p></li><li><p>Obtenir l’ID, le titre et l’index des diapositives sélectionnées à l’aide de la méthode <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> (lorsque vous spécifiez la nouvelle énumération <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a>) dans les compléments pour PowerPoint.</p></li></ul>|1.1|
|[Emplacement](../reference/outlook/Location.md)|Ajouté pour permettre la définition de l’emplacement d’un rendez-vous dans les compléments Outlook en mode composition.|1.1|
|[Bureau](../reference/shared/office.md)|Mise à jour de la méthode Select pour prendre en charge l’obtention des liaisons dans les compléments de contenu pour Access.|1.1|
|[Destinataires](../reference/outlook/Recipients.md)|Ajouté pour permettre l’obtention et la définition des destinataires d’un message ou d’un rendez-vous en mode composition.|1.1|
|[Paramètres](../reference/shared/document.settings.md)|Mis à jour pour prendre en charge la création de paramètres personnalisés dans les compléments de contenu pour Access.|1.1|
|[Objet](../reference/outlook/Subject.md)|Ajouté pour permettre l’obtention et la définition de l’objet d’un message ou d’un rendez-vous dans les compléments Outlook en mode composition.|1.1|
|[Heure](../reference/outlook/Time.md)|Ajouté pour permettre l’obtention et la définition de l’heure de début et de fin d’un rendez-vous dans les compléments Outlook en mode composition.|1.1|



**Énumérations ajoutées et énumérations mises à jour**


|**Objet**|**Description**|**Version**|
|:-----|:-----|:-----|
|[ActiveView](../reference/shared/activeview-enumeration.md)|Spécifie l’état de l’affichage dynamique du document, par exemple, si l’utilisateur peut modifier le document.Ajouté pour permettre aux compléments pour PowerPoint de déterminer si un utilisateur visualise une présentation ( **Diaporama**) ou modifie des diapositives. |1.1|
|[CoercionType](../reference/shared/coerciontype-enumeration.md)|Mis à jour avec  **Office.CoercionType.SlideRange** pour permettre la prise en charge de l’obtention des diapositives sélectionnées à l’aide de la méthode **getSelectedDataAsync** dans les compléments pour PowerPoint.|1.1|
|[EventType](../reference/shared/eventtype-enumeration.md)|Mis à jour pour inclure le nouvel événement ActiveViewChanged.|1.1|
|[FileType](../reference/shared/filetype-enumeration.md)|Mis à jour pour spécifier la sortie au format PDF.|1.1|
|[GoToType](../reference/shared/gototype-enumeration.md)|Ajouté pour spécifier l’emplacement ou l’objet à atteindre dans le document.|1.1|

## Ressources supplémentaires


- [API et schémas de référence pour les compléments Office](../reference/reference.md)
    
- [Compléments Office](../docs/overview/office-add-ins.md)
    
