# 青空 OSC: 青空文庫からデータを取得し VirtualCast へ送信するプログラム
このプログラムは、青空文庫からテキストデータを取得し、Open Sound Control (OSC) プロトコルを使用して VirtualCast へ送信するものです。別途公開中の「青空読書 OSC Aozora Reader」は、受信したデータを表示する VCI です。この連携により、VirtualCast 内で青空文庫のテキストが読めるようになります。  

## 動作環境
VirtualCast との連携を前提に作成されています。  

## インストール
インストール手順は以下の通りです：  
1. 「aozoraOSC.exe」 をダウンロードします（またはソースコードから自分でコンパイルしてください）。  
2. ウィルス対策ソフト等で「aozoraOSC.exe」をスキャンし、問題がないことを確認します。  
3. 「aozoraOSC.exe」を右クリックし、プロパティのセキュリティ項目で「許可する」にチェックを入れます。  
   - 右クリック→プロパティ→セキュリティ:このファイルは…☑許可する(K)  
4. 「aozoraOSC.exe」を実行します（設定ファイル等はありません）。  

## アンインストール
アンインストール手順は以下の通りです：  
1. 「aozoraOSC.exe」を削除します。   

## 使用方法
使用方法は以下の通りです：  
1. VirtualCast 内で VCI「青空読書 OSC Aozora reader」を出しておきます。  
2. 青空文庫の本文が表示されるページの URL を入力します。  
3. ［取得/Get］ボタンを押します。  
4. 送信したいページ範囲を選択します。  
5. ［VirtualCastへ送信］ボタンを押します。

## VirtualCast との連携
VCI と連携させるには、VirtualCast のタイトル画面で「VCI」メニュー内の「OSC受信機能」を「creator-only」または「enabled」に設定してください。「OSC送信機能」は使用していません。  

## VCIの取得
「青空読書 OSC Aozora reader」は [VirtualCastで公開中の商品ページ](https://virtualcast.jp/users/100215#products) から取得できます。

## プログラムの挙動
1. ［取得/Get］ボタンを押すと WebBrowser コントロールでデータを取得します。  

   ```C#:Form1.cs
   WebBrowser browser;  
   browser.Navigate(url);  
   ```

   対象としているタグは以下の通りです：  

   ```html:example.html
   <div class="metadata">  
     <h1 class="title">example</h1>  
     <h2 class="original_title">example</h2>  
     <h2 class="subtitle">example</h2>  
     <h2 class="author">example</h2>  
     <h2 class="translator">example</h2>  
   </div>  
   <div class="main_text">example</div>  
   <div class="bibliographical_information">example</div>  
   <div class="after_text">example</div>  
   ```
   - タグを合わせればローカルディスクから自作の HTML ファイルも読み込み可能ですが、それを前提に作成してはいません。  

2. ［VirtualCast へ送信］ボタンを押すと、以下のように OSC でデータを送信します：  
   2-1. OSC でメタ情報の送信  

   ```C#:Form1.cs
   UDPSender("127.0.0.1", 19100);  
   OscMessage("/Taki/aozoraReader/metaInfo", blob_title, blob_author, blob_translator, int_pages, blob_version);  
   ```
   - `blob_title`, `blob_author`, `blob_translator` には縦書き用のタグを追加しています。  

   2-2. OSC でページ及びテキスト情報の送信  

   ```C#:Form1.cs
   UDPSender("127.0.0.1", 19100);  
   OscMessage("/Taki/aozoraReader/partText", int_pageNumber, blob_partText);  
   ```
   - `blob_partText` は 20 文字 ✖ 10 行 に禁則処理を行った後、縦書き用のタグを追加しています。  

## ライセンス
このプログラムは MIT ライセンスのもとで公開されています。 

# OSC Aozora: A Program to Retrieve Data from Aozora Bunko and Send it to VirtualCast
This program retrieves text data from Aozora Bunko and sends it to VirtualCastusing the Open Sound Control (OSC) protocol. The separately released "OSC Aozora Reader" is a VCI that displays the received data. This integration allows you to read Aozora Bunko texts within VirtualCast.  

## System Requirements
This program is designed to be used in conjunction with VirtualCast.  

## Installation
Follow these steps to install:  
1. Download aozoraOSC.exe (or compile it from the source code yourself).  
2. Scan aozoraOSC.exe with antivirus software to ensure it is safe.  
3. Right-click aozoraOSC.exe, go to Properties, and check the security option to allow the program.  
   - Right-click → Properties → Security: This file is… ☑ Allow (K)  
4. Run aozoraOSC.exe (there are no configuration files required).  

## Uninstallation
Follow these steps to uninstall:  
1. Delete aozoraOSC.exe.   

## Usage
To use the program, follow these steps:  
1. Display the VCI "Aozora Reader OSC" within VirtualCast.  
2. Enter the URL of the page displaying the main text in Aozora Bunko.
3. Press the [Get] button.  
4. Select the page range you wish to send.  
5. Press the [Send to VirtualCast] button.  

## Integration with VirtualCast
To integrate with a VCI, set the "OSC Receive Function" in the VirtualCast title screen's "VCI" menu to "creator-only" or "enabled". This program and VCI, OSC Aozora reader, don't use the "OSC Send Function".  

## Obtaining the VCI
The "Aozora Reader OSC" can be obtained from the [products page on VirtualCast](https://virtualcast.jp/users/100215#products).  

## Program Behavior
1. Pressing the [Get] button retrieves data using the WebBrowser control:  

   ```C#:Form1.cs
   WebBrowser browser;  
   browser.Navigate(url);  
   ```

   The tags of interest are as follows:  

   ```html:example.html
   <div class="metadata">  
     <h1 class="title">example</h1>  
     <h2 class="original_title">example</h2>  
     <h2 class="subtitle">example</h2>  
     <h2 class="author">example</h2>  
     <h2 class="translator">example</h2>  
   </div>  
   <div class="main_text">example</div>  
   <div class="bibliographical_information">example</div>  
   <div class="after_text">example</div>  
   ```

   - Although it is possible to read custom HTML files from the local disk if the tags match, this is not the intended use.  

2. Pressing the [Send to VirtualCast] button sends the data via OSC as follows:  
   2-1. Sending metadata via OSC  

   ```C#:Form1.cs
   UDPSender("127.0.0.1", 19100);  
   OscMessage("/Taki/aozoraReader/metaInfo", `blob_title`, `blob_author`, `blob_translator`, `int_pages`, `blob_version`);  
   ```  

    - `blob_title`, `blob_author`, `blob_translator` include tags for vertical writing.  

   2-2. Sending page and text information via OSC  

   ```C#:Form1.cs
   UDPSender("127.0.0.1", 19100);  
   OscMessage("/Taki/aozoraReader/partText", `int_pageNumber`, `blob_partText`);  
   ```

    - `blob_partText` contains 20 characters ✖ 10 lines with typesetting rules applied and tags for vertical writing.  

## License
This program is licensed under the MIT License.  
