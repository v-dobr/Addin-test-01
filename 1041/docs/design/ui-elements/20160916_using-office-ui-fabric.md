
#Office アドインでの Office UI Fabric の使用

Office アドインを作成する場合は、[Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) を使用して、ユーザー エクスペリエンスを作成することをお勧めします。次の手順では、Fabric の基本的な使用方法について説明しています。  

##1.Fabric の設定
HTML の head セクション内に次の行を追加して、CDN の Fabric を参照します。

     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">


##2.Fabric のアイコンとフォントの使用
アイコンは簡単に使用できます。"i" 要素を使用して、適切なクラスを参照するだけです。アイコンのサイズは、フォント サイズを変更することで制御できます。

    <i class="ms-Icon ms-Icon--group" style="font-size:xx-large" aria-hidden="true"></i>


##3.単純なコンポーネントへのスタイルの使用
Fabric には、さまざまな UI 要素 (ボタンやチェック ボックスなど) に適したスタイルが用意されています。次の例に示すように、適切なクラスを参照するだけで、それに対応するスタイルを追加できます。

    <button class="ms-Button" id="get-data-from-selection">
    <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
    <span class="ms-Button-label">Get Data from selection</span>
    <span class="ms-Button-description">Get Data from the document selection</span>
    </button>

##4.サンプル動作を備えたコンポーネントの使用
Fabric には、動作 (クリック時の動作など) をサポートするコンポーネントもあります。 手始めに、Fabric 2.6.1 に含まれている jQuery UI プラグイン形式の**サンプル コード**を使用してみてください。 その他のフレームワークを使用して、必要なコードを記述することもできます。 サンプルを使用する場合は、サンプル コードが CDN の一部として配布されていない点に注意してください。サンプル コードは、[Fabric GitHub プロジェクト](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1)の 2.6.1 リリースからダウンロードし、そのコードを参照して、自分のコードで初期化する必要があります。 

たとえば、SearchBox コンポーネントを使用するには、次の手順を実行します。

1. [GitHub](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1/src/components/SearchBox) から、SearchBox コンポーネントをダウンロードします。
2. 次の参照をコードに追加します: `<script src="SearchBox/Jquery.SearchBox.js"></script>`
3. ページのロード時に、次の行が実行されるようにして、コンポーネントを初期化します: `$(".ms-SearchBox").SearchBox();`。作成するアドインの `Office.Initialize` ブロックに、これを含めるようにしてください。     

**注:**いくつかの Fabric コンポーネントのみを使用する場合は、コンポーネントごとに個別の CSS ファイルをホストすることで、ダウンロードするリソースのサイズを小さくできます。 CSS ファイルは、[Fabric 2.6.1 GitHub リポジトリ](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1)のコンポーネント フォルダーから入手できます。 


##次のステップ
Fabric の使用方法がすべてわかるようなサンプルをご用意しています。「[Office アドイン Fabric UI サンプル](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)」を参照してください。対話型の [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) Web サイトを参照することもできます。

