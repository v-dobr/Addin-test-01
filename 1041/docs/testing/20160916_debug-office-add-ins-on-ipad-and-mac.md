
# iPad と Mac で Office アドインをデバッグする

Windows でのアドインの開発とデバッグには Visual Studio を使用できますが、iPad と Mac で使用して アドインをデバッグすることはできません。アドインは HTML と Javascript を使用して開発されているため、さまざまなプラットフォームで機能するように設計されていますが、さまざまなブラウザーで HTML の表示方法に微妙な違いがあります。この記事では、iPad または Mac で動作するアドインをデバッグする方法を説明します。 

## Vorlon.js を使用したデバッグ 

Vorlon.js は、リモートで動作しさまざまなデバイスで Web ページをデバッグできる、F12 ツールに似た Web ページ用のデバッガーです。詳しくは、[Vorlon の Web サイト](http://www.vorlonjs.com)をご覧ください。  

Vorlon をインストールして設定するには 

1.  [Node.js](https://nodejs.org) と [Git](https://git-scm.com/) をインストールします (まだインストールしていない場合)。 

2.  git を使用して、次のコマンドで Vorlon をインストールします。`git clone https://github.com/MicrosoftDX/Vorlonjs.git`

3.  次のコマンドで依存関係をインストールします。`npm install`

4.  アドインは HTTPS を必要とするため、アドインで使用するすべてのスクリプトも同様に HTTPS になるように拡張する必要があります。これには、Vorlon スクリプトも含まれます。 そのため、アドインで Vorlon を使用するには、SSL を使用するように Vorlon を構成することが必要になります。 Vorlon のインストール フォルダーにある、/Server フォルダーに移動して config.json ファイルを編集します。 **useSSL** プロパティを **true** に変更します。 このとき、Office アドインのプラグインも有効にすることができます (プラグインの "enabled" プロパティを true に変更します)。 

5.  コマンド `sudo vorlon` を使用して Vorlon サーバーを実行します。 

6.  ブラウザー ウィンドウを開き、Vorlon インターフェイスの [http://localhost:1337](http://localhost:1337) に進みます。 セキュリティ証明書の信頼を求めるプロンプトが表示されるので、この証明書を信頼します。 セキュリティ証明書は、Vorlon フォルダーの /Server/cert 内にもあります。 

7.  次のスクリプト タグを、アドインの home.html ファイル (またはメイン HTML ファイル) の `<head>` セクションに追加します。
```    
<script src="https://localhost:1337/vorlon.js"></script>    
```  

これで、デバイスでアドインを表示したときに、アドインは常に Vorlon のクライアントのリスト (Vorlon インターフェイスの左側) に表示されます。リモートでの DOM 要素の強調表示、リモートでのコマンドの実行、その他多くの処理を実行できます。  

![Vorlon.js インターフェイスを示すスクリーン ショット](../../images/vorlon_interface.png)

Office プラグインにより Office.js に特別な機能 (オブジェクト モデルを調査する機能や Office.js の呼び出しを実行する機能など) が追加されます。 
