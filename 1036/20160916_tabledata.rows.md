
# Propriété TableData.rows
Obtient ou définit les lignes dans le tableau.

|||
|:-----|:-----|
|**Hôtes :**|Excel, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Ajouté dans**|1.1|

```
var myRows = tableBindingObj.rows;
```


## Valeur renvoyée

Renvoie un tableau de tableaux qui contient les données du tableau. Renvoie un **tableau**`[]` vide en l’absence de lignes.


## Remarques

Pour spécifier des lignes, vous devez spécifier un tableau de tableaux qui correspond à la structure de la table. Par exemple, pour spécifier les deux lignes de valeurs **string** dans une table à deux colonnes, vous devez définir la propriété **rows** sur ` [['a', 'b'], ['c', 'd']]`.

Si vous spécifiez une valeur **null** pour la propriété **rows** (ou si vous laissez la propriété vide quand vous construisez un objet **TableData**), vous obtenez les résultats suivants quand votre code s’exécute :


- Si vous insérez un nouveau tableau, une ligne vide est insérée.
    
- Si vous remplacez ou mettez à jour un tableau existant, les lignes existantes ne sont pas modifiées.
    

## Exemple

L’exemple suivant crée un tableau d’une seule colonne avec un en-tête et trois lignes.


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v|v|v|


|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|TableBindings|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de Word Online.|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad|
|1.0|Introduit|
