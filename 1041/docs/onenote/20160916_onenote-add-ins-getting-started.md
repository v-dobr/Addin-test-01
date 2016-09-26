# 最初の OneNote 用アドインをビルドする

この記事では、いくつかのテキストを OneNote ページに追加する簡単な作業ウィンドウ アドインのビルドについて説明します。

次の画像は、作成するアドインを示しています。

   ![このチュートリアルでビルドした OneNote アドイン](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## 手順 1:開発環境のセットアップ
1- [インストール手順](https://dev.office.com/docs/add-ins/get-started/create-an-office-add-in-using-any-editor)に従って、Yeoman Office ジェネレーターとその前提条件をインストールします。

   Yeoman Office ジェネレーターを使うと、Visual Studio がない場合や普通の HTML、CSS、JavaScript 以外のテクノロジを使う場合に、アドイン プロジェクトの作成が簡単になります。また、テスト用にローカルの Gulp Web サーバーにすばやくアクセスできます。 

   >オプションで [Visual Studio を使用](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio)して、プロジェクト ファイルを作成できますが、組み込み Gulp サーバーのサポートは利用できません。

<a name="create-project"></a>
## 手順 2:アドイン プロジェクトの作成 
1- *onenote add-in* という名前のローカル フォルダーを作成します。

2- **cmd** プロンプトを開いて、**[onenote アドイン]** フォルダーに移動します。以下に示す `yo office` コマンドを実行します。

```
C:\your-local-path\onenote add-in\> yo office
```
>これらの手順には、Windows コマンド プロンプトを使いますが、その他のシェル環境でも同じように適用されます。 

3- 次のオプションを使ってプロジェクトを作成します。

| オプション | 値 |
|:------|:------|
| プロジェクト名 | OneNote アドイン |
| プロジェクトのルート フォルダー | (既定値の適用) |
| Office Project の種類 | 作業ウィンドウ アドイン |
| サポートされている Office アプリケーション | (任意で選択してください--OneNote ホストを後で追加します) |
| 使うテクノロジ | HTML、CSS、JavaScript |

<a name="manifest"></a>
## 手順 3:アドイン マニフェストの構成 
1- プロジェクト ファイルにある **manifest-onenote-add-in.xml** を開きます。**ホスト** セクションに次の行を追加します。これは、アドインが OneNote ホスト アプリケーションをサポートすることを指定します。

```
<Host Name="Notebook" />
```

**SourceLocation** が既に Gulp Web サーバー用にセットアップされていることに注目してください。

```
<SourceLocation DefaultValue="https://localhost:8443/app/home/home.html"/>
```

<a name="develop"></a>
## 手順 4:アドインの開発
任意のテキスト エディターや IDE を使ってアドインを開発できます。まだ Visual Studio Code をお試しいただいていない場合は、Linux、Mac OSX、Windows で[無料でダウンロード](https://code.visualstudio.com/)できます。

1- **[アプリ] または [ホーム]** フォルダーにある *home.html* を開きます。 

2- Office JavaScript API と [Office UI Fabric](http://dev.office.com/fabric) のスタイルとコンポーネントへの参照を編集します。

   a.fabric.components.min.css へのリンクのコメントを解除します。

   b.Office.js へのスクリプト参照を、*ベータ*版への次の参照に置き換えます。

```
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

Office 参照は次のようになります。

```
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.min.css" rel="stylesheet">
<link href="//appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css" rel="stylesheet">
<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

3- `<body>` 要素を次のコードに置き換えます。 これは、[Office UI Fabric コンポーネント](http://dev.office.com/fabric/components)を使用してテキスト領域とボタンを追加します。 **応答性の高いグリッド** レイアウトは、[Office UI Fabric スタイル](http://dev.office.com/fabric/styles)のセットから作成されます。 

```
<body class="ms-font-m">
    <div class="home flex-container">
        <div class="ms-Grid">
            <div class="ms-Grid-row ms-bgColor-themeDarker">
                <div class="ms-Grid-col">
                    <span class="ms-font-xl ms-fontColor-themeLighter ms-fontWeight-semibold">OneNote Add-in</span>
                </div>
            </div>
        </div>
        <br />
        <div class="ms-Grid">
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <label class="ms-Label">Enter content here</label>
                    <div class="ms-TextField ms-TextField--placeholder">
                        <textarea id="textBox" rows="5"></textarea>
                    </div>
                </div>
            </div>
            <div class="ms-Grid-row">
                <div class="ms-Grid-col">
                    <div class="ms-font-m ms-fontColor-themeLight header--text">
                        <button class="ms-Button ms-Button--primary" id="addOutline">
                            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
                            <span class="ms-Button-label">Add outline</span>
                            <span class="ms-Button-description">Adds the content above to the current page.</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
```

4- **[アプリ] または [ホーム]** フォルダーにある *home.js* を開きます。次に示すように、**Office.initialize** 関数を編集し、**[アウトラインの追加]** ボタンにクリック イベントを追加します。 

```
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
    $(document).ready(function () {
        app.initialize();

        // Set up event handler for the UI.
        $('#addOutline').click(addOutlineToPage);
    });
};
```
 
5- **getDataFromSelection** メソッドを次の **addOutlineToPage** メソッドに置き換えます。これにより、テキスト領域からコンテンツを取得し、そのコンテンツがページに追加されます。

```
// Add the contents of the text area to the page.
function addOutlineToPage() {        
    OneNote.run(function (context) {
       var html = '<p>' + $('#textBox').val() + '</p>';

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to load the page with the title property.             
        page.load('title'); 

        // Add an outline with the specified HTML to the page.
        var outline = page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function() {
                console.log('Added outline to page ' + page.title);
            })
            .catch(function(error) {
                app.showNotification("Error: " + error); 
                console.log("Error: " + error); 
                if (error instanceof OfficeExtension.Error) { 
                    console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                } 
            }); 
        });
}
```

<a name="test"></a>
## 手順 5:OneNote Online でのアドインのテスト
1- Gulp Web サーバーを実行します。  

   a. **onenote add-in** フォルダーで **cmd** プロンプトを開きます。 

   b. 以下に示す `gulp serve-static` コマンドを実行します。

```
C:\your-local-path\onenote add-in\> gulp serve-static
```

2- Gulp Web サーバーの自己署名証明書を信頼された証明書としてインストールします。Yeoman Office ジェネレーターを使って作成されたアドイン プロジェクトに対しては、コンピューターに一度だけインストールする必要があります。

   a.ホストされたアドイン ページに移動します。これは、既定ではマニフェストにあるのと同じ URL です。

```
https://localhost:8443/app/home/home.html
```

   b. 証明書を信頼された証明書としてインストールします。 詳しくは、「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/docs/trust-self-signed-cert.md)」をご覧ください。

3- OneNote Online でノートブックを開きます。

4- **[挿入] > [Office アドイン]** を選択します。 これで、[Office アドイン] ダイアログが開きます。
  - コンシューマー アカウントでログインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。
  - 職場または学校アカウントでログインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。 
  
  次の図は、コンシューマー ノートブックの **[マイ アドイン]** タブを示しています。

  ![[マイ アドイン] タブを示す [Office アドイン] ダイアログ](../../images/onenote-office-add-ins-dialog.png)
  
  >**注**:OneNote ページ内でクリックすると、**[Office アドイン]** ボタンが有効になります。

5- [アドインのアップロード] ダイアログで、プロジェクト ファイル内の **manifest-onenote-add-in.xml** を参照し、**[アップロード]** を選択します。 テスト中、マニフェスト ファイルはローカルに保存できます。

6- アドインは、OneNote ページの横にある iFrame で開きます。 テキスト領域にテキストを入力し、**[アウトラインの追加]** をクリックします。 入力したテキストは、ページに追加されます。 

## トラブルシューティングとヒント
- ブラウザーの開発者ツールを使ってアドインをデバッグできます。Gulp Web サーバーを使っており、Internet Explorer や Chrome でデバッグしている場合は、ローカルで変更を保存して、アドインの iFrame を更新するだけです。

- OneNote オブジェクトを調べる場合、現在使用可能なプロパティに実際の値が表示されます。読み込む必要のあるプロパティには、*undefined* と表示されます。`_proto_` ノードを展開し、オブジェクトで定義されているものの、まだ読み込まれていないプロパティを確認します。

      ![Unloaded OneNote object in the debugger](../../images/onenote-debug.png)

- アドインで任意の HTTP リソースを使っている場合は、ブラウザーで混在したコンテンツを有効にする必要があります。運用アドインでは、セキュリティで保護された HTTPS リソースのみを使う必要があります。

## その他のリソース

- [OneNote の JavaScript API のプログラミングの概要](onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API リファレンス](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](https://dev.office.com/docs/add-ins/overview/office-add-ins)
