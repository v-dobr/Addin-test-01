
# Conseils pour la gestion des valeurs de date dans les compléments Outlook

L’interface API JavaScript pour Office utilise l’objet JavaScript [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp) pour stocker et récupérer la plupart des dates et des heures. Cet objet **Date** fournit des méthodes telles que [getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp) et [toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp), qui renvoient la date ou l’heure UTC demandée.<br/><br/>
L’objet **Date** fournit également d’autres méthodes telles que [getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp) et [toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp), qui renvoient la date ou l’heure locale demandée.<br/><br/>Le concept d’« heure locale » est principalement déterminé par le navigateur et le système d’exploitation de l’ordinateur client. Par exemple, dans la plupart des navigateurs s’exécutant sur un ordinateur client Windows, un appel JavaScript à **getDate** renvoie une date en fonction du fuseau horaire défini dans Windows sur l’ordinateur client.<br/><br/>
L’exemple suivant crée un objet **Date**<code>myLocalDate</code> au format de l’heure locale, et appelle **toUTCString** pour convertir cette date en chaîne de date au format UTC.




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Alors que vous pouvez utiliser l’objet JavaScript  **Date** pour obtenir une valeur de date ou d’heure basée sur UTC ou le fuseau horaire de l’ordinateur client, l’objet **Date** présente une restriction en ce sens qu’il ne fournit pas de méthodes pour renvoyer une valeur de date ou d’heure pour tout autre fuseau horaire spécifique. Par exemple, si votre ordinateur est réglé sur le fuseau horaire Heure normale de l’Est (EST), aucune méthode **Date** ne vous permet d’obtenir la valeur horaire autre que EST ou UTC, par exemple Heure normale du Pacifique (PST).


## Fonctionnalités liées à la date pour les compléments Outlook


La restriction JavaScript mentionnée plus haut a une incidence pour vous lorsque vous utilisez l’interface API JavaScript pour Office pour gérer des valeurs de date et d’heure dans des compléments Outlook qui s’exécutent sur le client riche Outlook et dans Outlook Web App ou OWA pour périphériques.


### Fuseaux horaires pour les clients Outlook

Pour clarifier les choses, définissons les fuseaux horaires en question.



|**Fuseau horaire**|**Description**|
|:-----|:-----|
|Fuseau horaire de l’ordinateur client|Cette valeur est définie sur le système d’exploitation de l’ordinateur client. La plupart des navigateurs utilisent le fuseau horaire de l’ordinateur client pour afficher les valeurs de date ou d’heure de l’objet JavaScript **Date**.<br/><br/>Le client riche Outlook utilise ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur. <br/><br/>Par exemple, sur un ordinateur client exécutant Windows, Outlook utilise le fuseau horaire défini sur Windows comme fuseau horaire local. Sous Mac, si l’utilisateur change le fuseau horaire sur l’ordinateur client, Outlook pour Mac invite l’utilisateur à le mettre à jour dans Outlook également.|
|Fuseau horaire EAC (Exchange Admin Center)|L’utilisateur définit cette valeur de fuseau horaire (et la langue préférée) lorsqu’il se connecte à Outlook Web App ou OWA pour les appareils pour la première fois. <br/><br/>Outlook Web App et OWA pour les appareils utilisent ce fuseau horaire pour afficher les valeurs de date ou d’heure dans l’interface utilisateur.|
Comme le client riche Outlook utilise le fuseau horaire de l’ordinateur client, alors que l’interface utilisateur d’Outlook Web App et d’OWA pour périphériques utilise le fuseau horaire EAC, l’heure locale du même complément installé pour la même boîte aux lettres peut être différente lors d’une exécution sur le client riche Outlook et sur Outlook Web App ou OWA pour périphériques. En tant que développeur de complément Outlook, vous devez entrer et sortir de façon appropriée les valeurs de date afin qu’elles soient toujours en accord avec le fuseau horaire attendu par l’utilisateur sur le client correspondant.


### API liée à la date

Les propriétés et les méthodes suivantes de l’interface API JavaScript pour Office prennent en charge des fonctionnalités associées à la date.reference/outlook/Office.context.mailbox.item.md



**Membre de l'API**|**Représentation du fuseau horaire**|**Exemple dans un client riche Outlook**|**Exemple dans Outlook Web App ou OWA pour périphériques**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|Dans un client riche Outlook, cette propriété renvoie le fuseau horaire de l’ordinateur client. Dans Outlook Web App et OWA pour périphériques, cette propriété renvoie le fuseau horaire EAC. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) et [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Chacune de ces propriétés renvoie un objet JavaScript  **Date**. Cette valeur **Date** est conforme au format UTC, comme illustré dans l’exemple suivant - `myUTCDate` a la même valeur dans un client riche Outlook, Outlook Web App et OWA pour les appareils.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Toutefois, l’appel de `myDate.getDate` renvoie une valeur de date dans le fuseau horaire de l’ordinateur client, qui est cohérente avec le fuseau horaire utilisé pour afficher les valeurs de date et d’heure dans l’interface du client riche Outlook, mais qui peut être différente du fuseau horaire du CAE utilisé par Outlook Web App et OWA pour les appareils dans l’interface utilisateur.|Si l’élément est créé à 9 h 00 UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h 00 UTC :<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` renvoie 6 h 00 EST.|Si l’élément est créé à 9 h 00 UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` renvoie 4 h 00 EST.<br/><br/>Si l’élément est modifié à 11 h 00 UTC :<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` renvoie 6 h 00 EST.<br/><br/>Notez que si vous souhaitez afficher l’heure de création ou de modification dans l’interface utilisateur, vous pouvez d’abord convertir l’heure au format PST pour rester cohérent avec le reste de l’interface utilisateur.
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|Chacun des paramètres  _Start_ et _End_ nécessite un objet JavaScript **Date**. Les arguments doivent être conformes à UTC, quel que soit le fuseau horaire utilisé dans l’interface utilisateur d’un client riche Outlook, Outlook Web App ou OWA pour périphériques.|Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>|Si les heures de début et de fin du formulaire de rendez-vous sont 9h00 UTC et 11h00 UTC, vous devez vous assurer que les arguments `start` et `end` sont conformes au format UTC, autrement dit :<br/><br/><ul><li>`start.getUTCHours` renvoie 9 h 00 UTC</li><li>`end.getUTCHours` renvoie 11 h 00 UTC</li></ul>

## Méthodes d’assistance pour les scénarios liés à la date


Comme indiqué précédemment, l’« heure locale » d’un utilisateur dans Outlook Web App ou OWA pour les appareils peut être différente sur un client riche Outlook, mais le code JavaScript  **Date** prend uniquement en charge la conversion vers le fuseau horaire ou l’heure UTC de l’ordinateur client. Par conséquent, l’interface API JavaScript pour Office fournit deux méthodes d’assistance : [Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) et [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md). <br/><br/>
Ces méthodes d’assistance vous permettent de gérer différemment la date ou l’heure pour ces deux scénarios liés à la date dans un client riche Outlook, Outlook Web App et OWA pour les appareils, ce qui vous permet de renforcer l’« écriture unique » pour les autres clients de votre complément.


### Scénario A : affichage de l’heure de création ou de modification d’un élément

Si vous affichez l’heure de création (**Item.dateTimeCreated**) ou de modification (**Item.dateTimeModified**) d’un élément dans l’interface utilisateur, utilisez d’abord  **convertToLocalClientTime** pour convertir l’objet **Date** fourni par ces propriétés pour obtenir une représentation de dictionnaire dans l’heure locale appropriée. Affichez ensuite les parties de la date de dictionnaire. L’exemple suivant illustre ce scénario :


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook web app or OWA for Devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Notez que  **convertToLocalClientTime** gère la différence entre le client riche Outlook et Outlook Web App ou OWA pour périphériques :


- Si  **convertToLocalClientTime** détecte que l’hôte actuel est un client riche, la méthode convertit la représentation **Date** en une représentation de dictionnaire dans le fuseau horaire de l’ordinateur client, en accord avec le reste de l’interface utilisateur du client riche.
    
- Si  **convertToLocalClientTime** détecte que l’hôte actuel est Outlook Web App ou OWA pour périphériques, la méthode convertit la représentation **Date** conforme à UTC en un format de dictionnaire dans le fuseau horaire EAC, en accord avec le reste de l’interface utilisateur d’Outlook Web App ou d’OWA pour périphériques.
    

### Scénario B : affichage des dates de début et de fin dans un formulaire de nouveau rendez-vous

Si vous obtenez différentes parties d’une valeur d’entrée date-heure à l’heure locale et que vous souhaitez fournir la valeur d’entrée du dictionnaire sous la forme d’une heure de début ou de fin dans un formulaire de rendez-vous, utilisez d’abord la méthode d’assistance  **convertToUtcClientTime** pour convertir la valeur de dictionnaire en objet **Date** au format UTC.<br/><br/>Dans l’exemple suivant, supposons que  `myLocalDictionaryStartDate` et `myLocalDictionaryEndDate` sont des valeurs de date et d’heure au format de dictionnaire que vous avez obtenues auprès de l’utilisateur. Ces valeurs sont basées sur l’heure locale, qui dépend elle-même de l’application hôte.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

Les valeurs qui en résultent, `myUTCCorrectStartDate` et `myUTCCorrectEndDate`, sont au format UTC. Transférez ensuite ces objets  **Date** en tant qu’arguments pour les paramètres _Start_ et _End_ de la méthode **Mailbox.displayNewAppointmentForm** pour afficher le nouveau formulaire de rendez-vous.<br/><br/>
Notez que **convertToUtcClientTime** gère la différence entre le client riche Outlook et Outlook Web App ou OWA pour les appareils :


- Si  **convertToUtcClientTime** détecte que l’hôte actuel est un client riche Outlook, la méthode convertit simplement la représentation de dictionnaire en objet **Date**. Cet objet  **Date** est conforme à UTC, tel qu’attendu par **displayNewAppointmentForm**.
    
- Si  **convertToUtcClientTime** détecte que l’hôte actuel est Outlook Web App ou OWA pour périphériques, la méthode convertit le format de dictionnaire des valeurs de date et d’heure exprimées dans le fuseau horaire EAC en objet **Date**. Cet objet  **Date** est conforme à UTC, comme le prévoit **displayNewAppointmentForm**.
    

## Ressources supplémentaires



- [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md)
    


