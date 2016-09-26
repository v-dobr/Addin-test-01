
# Office アドインを展開し、発行する


さまざまな方法を利用し、テスト目的またはユーザーに配布する目的で、Office アドインを展開できます。

- [サイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) - 開発プロセスの一環として利用し、Windows、Office Online、iPad、Mac で実行されているアドインをテストします。
- [SharePoint カタログ](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) - 開発プロセスの一環として利用し、アドインをテストしたり、アドインを組織のユーザーに配布したりします。
- [Office 365 管理センター プレビュー](https://support.office.com/en-ie/article/Deploy-Office-Add-Ins-in-Office-365-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE) - 組織のユーザーにアドインを配布するために使用します。
- [Office ストア] ユーザーに配布する目的でアドインを公開するために使用します。

利用できるオプションは、ターゲットとする Office タイプや作成するアドインの種類によって異なります。

### Word、Excel、PowerPoint アドインの開発オプション

| 拡張点            | サイドロード | SharePoint カタログ | Office 365 管理センター プレビュー | Office ストア |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| コンテンツ         | X           | X                  | X                               | X            |
| 作業ウィンドウ       | X           | X                  | X                               | X            |
| コマンド         | X           |                    | X                               | X            |

> **注**:SharePoint カタログは Office 2016 for Mac ではサポートされていません。 Office アドインを Mac クライアントに展開するには、それを [Office ストア]に提出する必要があります。    

### Outlook アドインの展開オプション

| 拡張点     | サイドロード | Exchange サーバー | Office ストア |
|:---------|:-----------:|:---------------:|:------------:|
| メール アプリ | X           | X               | X            |
| コマンド  | X           | X               | X            |

アドインの範囲を広げるには、アドインがプラットフォームを横断して動作するようにします。 Office アドインは、Windows、Mac、Web、iOS、Android でサポートされています。 各プラットフォームでサポートされている機能の概要については、「[Office アドインを使用できるホストおよびプラットフォーム]」を参照してください。   

Office ストアのアドインについては、「[アドインのライセンス](https://msdn.microsoft.com/EN-US/library/office/jj163257.aspx)」を参照してください。

エンド ユーザーがアドインを取得、挿入、実行する方法については、「[Office アドインの使用を開始する](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)」を参照してください。

## その他のリソース

- [Office アドインを使用できるホストおよびプラットフォーム]
- [テスト用に Outlook アドインを展開してインストールする](../outlook/testing-and-tips.md) 
- [Office ストアにアドインと Web アプリを提出する][Office ストア]
- [Office アドインの設計ガイドライン](../design/add-in-design)
- [効果的な Office ストア アドインを作成する](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)

[Office ストア]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office アドインを使用できるホストおよびプラットフォーム]: http://dev.office.com/add-in-availability
