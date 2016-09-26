
# Outlook アドインで日付値を扱うためのヒント

JavaScript API for Office では、日付と時刻の保存および取得のほとんどで、JavaScript の [Date](http://www.w3schools.com/jsref/jsref_obj_date.asp) オブジェクトを使用します。この **Date** オブジェクトには、[getUTCDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getUTCHour](http://www.w3schools.com/jsref/jsref_getutchours.asp)、[getUTCMinutes](http://www.w3schools.com/jsref/jsref_getutcminutes.asp)、[toUTCString](http://www.w3schools.com/jsref/jsref_toutcstring.asp) などのメソッドがあります。これらのメソッドは、要求された日付または時刻の値を協定世界時 (UTC) に従って返します。<br/><br/>
**Date** オブジェクトには、[getDate](http://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getHour](http://www.w3schools.com/jsref/jsref_getutchours.asp)、[getMinutes](http://www.w3schools.com/jsref/jsref_getminutes.asp)、[toString](http://www.w3schools.com/jsref/jsref_tostring_date.asp) などのメソッドもあります。これらのメソッドは、要求された日付または時刻を "現地時刻" に従って返します。<br/><br/>"現地時刻" の概念は、主にクライアント コンピューター上のブラウザーおよびオペレーティング システムによって判断されます。たとえば、Windows ベースのクライアント コンピューター上で動作している大部分のブラウザーでは、JavaScript で **getDate** を呼び出すと、クライアント コンピューター上の Windows で設定されているタイム ゾーンに基づく日付が返されます。<br/><br/>
次の例では、**myLocalDate** という名前の <code>Date</code> オブジェクトを現地時刻で作成し、**toUTCString** を呼び出して、その日付を UTC の日付文字列に変換します。




```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

JavaScript の **Date** オブジェクトを使用して、UTC またはクライアント コンピューターのタイム ゾーンに基づく日付や時刻の値を取得することはできますが、**Date** オブジェクトには次の点で制限があります。他のタイム ゾーンを指定して日付や時刻の値を返すメソッドはありません。たとえば、東部標準時 (EST) に設定されているクライアント コンピューターの場合、太平洋標準時 (PST) など、EST または UTC 以外の時刻値を取得できる **Date** メソッドはありません。


## Outlook アドインの日付関連機能


前述の JavaScript の制限は、JavaScript API for Office を使用して、Outlook リッチ クライアントおよび Outlook Web App またはデバイス用 OWA 上で動作する Outlook アドインの日付または時刻の値を処理する場合にも影響します。


### Outlook クライアントのタイム ゾーン

わかりやすくするため、問題のタイム ゾーンを定義します。



|**タイム ゾーン**|**説明**|
|:-----|:-----|
|クライアント コンピューターのタイム ゾーン|これは、クライアント コンピューターのオペレーティング システムで設定されています。ほとんどのブラウザーは、クライアント コンピューターのタイム ゾーンを使用して、JavaScript の **Date** オブジェクトの日付または時刻の値を表示します。<br/><br/>Outlook リッチ クライアントでは、このタイム ゾーンを使用して、ユーザー インターフェイスの日付または時刻の値を表示します。 <br/><br/>たとえば、Windows を実行しているクライアント コンピューター上の Outlook では、Windows 上で設定されているタイム ゾーンをローカル タイム ゾーンとして使用します。Mac の場合、ユーザーがクライアント コンピューター上のタイム ゾーンを変更すると、Outlook for Mac によってタイム ゾーンを更新するように求めるメッセージが、Outlook と同様に表示されます。|
|Exchange 管理センター (EAC) のタイム ゾーン|ユーザーは、最初に Outlook Web App またはデバイス用 OWA にログオンするとき、このタイム ゾーンの値 (および優先する言語) を設定します。 <br/><br/>Outlook Web App およびデバイス用 OWA では、このタイム ゾーンを使用して、ユーザー インターフェイスの日付または時刻の値を表示します。|
Outlook リッチ クライアントはクライアント コンピューターのタイム ゾーンを使用し、Outlook Web App とデバイス用 OWA のユーザー インターフェイスは EAC タイム ゾーンを使用するので、同じメールボックスに対してインストールされている同じアドインの現地時刻が、Outlook リッチ クライアントで実行しているときと、Outlook Web App またはデバイス用 OWA で実行しているときとで、異なる場合があります。Outlook アドイン開発者としては、日付値の入出力を適切に行い、その値が、対応するクライアント上で期待されるタイム ゾーンと一致するようにしておく必要があります。


### 日付関連の API

日付関連機能をサポートする JavaScript API for Office のプロパティおよびメソッドを次に示します。reference/outlook/Office.context.mailbox.item.md



**API メンバー**|**タイム ゾーン表現**|**Outlook リッチ クライアントの例**|**Outlook Web App またはデバイス用 OWA の例**
--------------|----------------------------|-------------------------------------|-------------------------------------------------
[Office.context.mailbox.userProfile.timeZone](../../reference/outlook/Office.context.mailbox.userProfile.md)|Outlook リッチ クライアントでは、このプロパティはクライアント コンピューターのタイム ゾーンを返します。Outlook Web App およびデバイス用 OWA では、このプロパティは EAC タイム ゾーンを返します。 |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) および [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|これらのプロパティはどちらも、JavaScript の **Date** オブジェクトを返します。次の例に示すように、この **Date** の値は正しい UTC 時刻です。`myUTCDate` には、Outlook リッチ クライアント、Outlook Web App、およびデバイス用 OWA で同じ値が入ります。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>ただし、`myDate.getDate` を呼び出すと、クライアント コンピューターのタイム ゾーンで日付値が返されます。この値は、Outlook リッチ クライアントのインターフェイスで日時値を表示する際に使用されるタイム ゾーンと一致しますが、Outlook Web App およびデバイス用 OWA のユーザー インターフェイスで使用される EAC タイム ゾーンとは異なる場合があります。|アイテムが午前 9 時 (UTC) に作成された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。|アイテムの作成時刻が午前 9 時 (UTC) の場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。<br/><br/>ユーザー インターフェイスで作成時刻や変更時刻を表示する場合は、まず時刻を PST に変換して、他のユーザー インターフェイスと一貫性を保つようにします。
[Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)|_Start_ パラメーターと _End_ パラメーターには、それぞれ JavaScript の **Date** オブジェクトが必要です。引数は正しい UTC である必要があります。これは Outlook リッチ クライアント、Outlook Web App、またはデバイス用 OWA のユーザー インターフェイスで使用されているタイム ゾーンのいずれでも同じです。|予定フォームの開始時刻と終了時刻が午前 9 時 (UTC) と午前 11 時 (UTC) の場合、`start` と `end` の引数は正しい UTC 時刻である必要があります。つまり、<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>|予定フォームの開始時刻と終了時刻が午前 9 時 (UTC) と午前 11 時 (UTC) の場合、`start` と `end` の引数は正しい UTC 時刻である必要があります。つまり、<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>

## 日付関連のシナリオ向けのヘルパー メソッド


前のセクションで説明したように、Outlook Web App またはデバイス用 OWA のユーザーの "現地時刻" は Outlook リッチ クライアント上で異なる場合があるのに対して、JavaScript の **Date** オブジェクトではクライアント コンピューターのタイム ゾーンまたは UTC への変換のみサポートされるため、JavaScript API for Office には 2 つのヘルパー メソッド[Office.context.mailbox.convertToLocalClientTime](../../reference/outlook/Office.context.mailbox.md) と [Office.context.mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md) があります。 <br/><br/>
このヘルパー メソッドを使用すると、次の 2 つの日付関連のシナリオ (Outlook リッチ クライアント、Outlook Web App、およびデバイス用 OWA) で、日付または時刻に異なる処理が必要な場合に対処できます。つまり、このヘルパー メソッドは、アドインの異なるクライアントごとの "ライト ワンス" を強化するものです。


### シナリオ A: アイテムの作成時刻または変更時刻を表示する

ユーザー インターフェイスにアイテムの作成時刻 (**Item.dateTimeCreated**) または変更時刻 (**Item.dateTimeModified**) を表示している場合、まず **convertToLocalClientTime** を使用して、これらのプロパティで提供される **Date** オブジェクトを変換し、適切な現地時刻の辞書表現を取得します。その後、辞書の日付部分を表示します。このシナリオの例を次に示します。


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

**convertToLocalClientTime** は Outlook リッチ クライアントと Outlook Web App またはデバイス用 OWA との違いに、次のように対処します。


- **convertToLocalClientTime** メソッドでは、現在のホストがリッチ クライアントであることを検出すると、**Date** 表現を同じクライアント コンピューター タイム ゾーンの辞書表現に変換して、他のリッチ クライアント ユーザー インターフェイスとの一貫性を保ちます。
    
- **convertToLocalClientTime** メソッドでは、現在のホストが Outlook Web App またはデバイス用 OWA であることを検出すると、正しい UTC の **Date** 表現を EAC タイム ゾーンの辞書形式に変換して、他の Outlook Web App またはデバイス用 OWA ユーザー インターフェイスとの一貫性を保ちます。
    

### シナリオ B: 新しい予定フォームの開始日付と終了日付を表示する

現地時刻で表された日付と時刻の値の異なる各部分を、入力として取得しているときに、この辞書の入力値を予定フォームの開始時刻または終了時刻として提供する場合は、まず **convertToUtcClientTime** ヘルパー メソッドを使用して、ディクショナリ値を正しい UTC の **Date** オブジェクトに変換します。<br/><br/>次の例では、`myLocalDictionaryStartDate` および `myLocalDictionaryEndDate` をユーザーから取得した辞書形式の日付と時刻の値と仮定しています。これらの値は、ホスト アプリケーションに依存する、現地時刻に基づいています。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

出力結果の値 `myUTCCorrectStartDate` と `myUTCCorrectEndDate` は、正しい UTC です。次に、これらの **Date** オブジェクトを _Mailbox.displayNewAppointmentForm_ メソッドの _Start_ パラメーターと **End** パラメーターの引数として渡し、新しい予定フォームを表示します。<br/><br/>
**convertToUtcClientTime** は Outlook リッチ クライアントと Outlook Web App またはデバイス用 OWA との違いに、次のように対処します。


- **convertToUtcClientTime** では、現在のホストが Outlook リッチ クライアントであることを検出すると、単純に辞書表現を **Date** オブジェクトに変換します。この **Date** オブジェクトは、**displayNewAppointmentForm** で想定される正しい UTC です。
    
- **convertToUtcClientTime** では、現在のホストが Outlook Web App またはデバイス用 OWA であることを検出すると、EAC タイム ゾーンで表される辞書形式の日付と時刻の値を、**Date** オブジェクトに変換します。この **Date** オブジェクトは、**displayNewAppointmentForm** で想定される正しい UTC です。
    

## その他のリソース



- [テスト用に Outlook アドインを展開してインストールする](../outlook/testing-and-tips.md)
    


