using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace aozoraOSC
{
    public partial class Form1 : Form
    {
        // aozora OSC version
        private readonly string aozoraVersion = "v2.0";
        private string homeUrl = "https://www.aozora.gr.jp";

        // WebBrowser コントロールの実体が IE (ver.11) であることや、
        // 読み込み中から復帰できないページがあるといった関係から、
        // 青空文庫とローカルファイルに限定。
        //
        // 許可する開始文字列
        // たくさん設定できるので、List を採用
        //
        // 許可する終了文字列は指定しない
        // ※拡張子指定やディレクトリ判定を行ってもMIME判定不能…
        // 　例えばこれは回避不能 → https://www.aozora.gr.jp/cards/000148/files/789
        // 
        private readonly List<string> _allowedStartWith = new List<string>
        {
            "about:blank",
            @"^[A-Z]:\\",
            @"^[A-Z]:/",
            "file:///",
            "http://www.aozora.gr.jp",
            "https://www.aozora.gr.jp"

            // "http://search.bungo.app",
            // "https://search.bungo.app",
            // "http://yozora.main.jp",
            // "https://yozora.main.jp",
            // "https://virtualcast.jp", IE v.11 対象外: サイトの作り上、 DocumentCompleted しない?
        };


        // 初期設定
        private readonly int firstPage = 1;
        private int lastPage = 2;
        private string[] pages = { "　", "<rotate=90><rotate=0>（<rotate=90>本文<rotate=0>：<rotate=90>情報無し<rotate=0>）<rotate=90>" };

        // 無限ループ回避用（C#では bool 型のデフォルト値は false）
        private bool firstPageReset = false;
        private bool lastPageReset = false;

        // WebBrowserコントロールをクラスレベルで宣言
        private WebBrowser textBrowser;

        // null対応
        private WebBrowser nullBrowser;

        // 連打防止用
        private bool isProcessing = false;

        // OSC 中止用
        private bool isBreak = false;

        //はじめに表示用
        private bool isIntroduce = true;

        //
        // 起動後最初に実行される
        //
        public Form1()
        {
            InitializeComponent();

            // フォームタイトルにバージョンを追加
            this.Text += " " + aozoraVersion;

            backButton.Enabled = false;
            forwardButton.Enabled = false;

            // 既定値 URL を選択状態にしておく（URL 貼り付け操作を楽にする）
            textBox1.SelectAll();

            // 送信ページの範囲指定（lastPage は自動で設定される）
            comboBox1.SelectedItem = firstPage.ToString();


            string introduction = "<!DOCTYPE html><html lang=\"ja\"><head><meta charset=\"utf-8\"></head><body>はじめに<br><br>便利なショートカットキー：[Alt] + [←] / [→]<br><br>このウィンドウは、インターネットエクスプローラー (ver.11) の技術を使っています。技術面、安全面の問題からアクセスは以下のサイトに限定しています。<br>・青空文庫：<a href=\"https://www.aozora.gr.jp/\">https://www.aozora.gr.jp</a><br>・ローカルドライブ：例）file:///c:/<br><br><br>Introduction<br><br>Useful Shortcut Keys: [Alt] + [←] / [→]<br><br>This window utilizes Internet Explorer (ver.11) technology. For technical and security reasons, access is restricted to the following sites:<br><br>・Aozora Bunko: <a href=\"https://www.aozora.gr.jp/\">https://www.aozora.gr.jp</a><br>・Local Drive: e.g. file:///c:/<br></body></html>";

            // 右側テキストブラウザ
            textBrowser = new WebBrowser();
            textBrowser.ScriptErrorsSuppressed = true;
            textBrowser.DocumentText = introduction;
            textBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(textBrowser_DocumentCompleted);

            nullBrowser = new WebBrowser();
            string dummyHtml = "";
            nullBrowser.DocumentText = dummyHtml;

            webBrowser1.DocumentText = introduction;

            // ref. https://learn.microsoft.com/ja-jp/dotnet/desktop/winforms/controls/how-to-add-web-browser-capabilities-to-a-windows-forms-application?view=netframeworkdesktop-4.8
            // The following events are not visible in the designer, so
            // you must associate them with their event-handlers in code.
            webBrowser1.CanGoBackChanged +=
                new EventHandler(webBrowser1_CanGoBackChanged);
            webBrowser1.CanGoForwardChanged +=
                new EventHandler(webBrowser1_CanGoForwardChanged);
        }

        //
        // URL 入力後の Enter キー処理
        //
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter キーが押されたときの処理
                button1_Click(sender, e); // button1_Click を呼び出す
                e.Handled = true; // イベントの処理を完了したことを示す
            }
        }

        private void isProcessingTrue()
        {             
            // 処理中
            isProcessing = true;

            // ボタンとコンボボックス（ページ範囲指定）を無効化し、テキストを変更
            textBox1.Enabled = false;
            button1.Enabled = false;
            button1.Text = "⌛";
            button2.Enabled = false;
            button2.Text = "⌛取得中";
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;

            // ブラウザ関係
            backButton.Enabled = false;
            forwardButton.Enabled = false;
            homeButton.Enabled = false;
        }


        private void isProcessingFalse()
        {
            // ボタンとコンボボックスを再度有効化し、テキストを戻す
            textBox1.Enabled = true;
            button1.Enabled = true;
            button1.Text = "➤";
            button2.Enabled = true;
            button2.Text = "➤VirtualCast";
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            label3.Text = "";
            label8.Text = "";

            // ブラウザ関係
            backButton.Enabled = webBrowser1.CanGoBack;
            forwardButton.Enabled = webBrowser1.CanGoForward;
            homeButton.Enabled = true;

            // 処理終了
            isProcessing = false;
        }

        //
        // URL からデータ取得
        //
        private async void button1_Click(object sender, EventArgs e)
        {
            // 処理中でないことを確認
            if (isProcessing)
                return;

            string url = textBox1.Text;

            // 許可された URL かチェック
            bool startsWithAllowed = _allowedStartWith.Any(pattern => Regex.IsMatch(url, pattern, RegexOptions.IgnoreCase));

            if (startsWithAllowed)
            {
                try
                {
                    // ページ範囲の無限ループ回避用
                    firstPageReset = true;
                    lastPageReset = true;

                    // ボタンを無効化し、「⌛」を表示
                    isProcessingTrue();
                    label3.Text = "✖";
                    
                    // 連打防止の強化
                    await Task.Delay(250);

                    // ページ作成処理を開始し、完了を待つ
                    await RetrieveHtml();
                    // RetrieveHtml()
                    //  -> webBrowser1.Navigate(url) => webBrowser1_Navigating() ここで処理中にする
                    //    -> textBrowser.Navigate(url) => textBrowser_DocumentCompleted() ここで処理終了

                }
                finally // 例外が発生しても実行される
                {
                    // 念のため復帰しておく just in case
                    // ボタンを再度有効化し、テキストを戻す
                    isProcessingFalse();
                }
            }
            else
            {
                // 許可されていないドメインへのナビゲーションをキャンセル
                textBox2.Text = "⚠️許可範囲外⚠️";
                textBox3.Text = "⚠️リンク先: " + url + " ⚠️";

                // ボタンを再度有効化し、テキストを戻す
                isProcessingFalse();
            }
        }


        //
        // OSC で VirtualCast へ送信
        //
        private async void button2_Click(object sender, EventArgs e)
        {
            // 処理中でないことを確認
            if (isProcessing)
                return;

            // ボタンを無効化し、「⌛」を表示
            isProcessingTrue();

            try
            {
                // 時間のかかる処理を行う
                await SelectDataForSending();
            }
            finally
            {
                // ボタンを再度有効化し、テキストを戻す
                isProcessingFalse();
            }
        }


        //
        // ［取得/Get］ボタンが押されたとき
        // 指定ページを読み込み、 EditData を実行
        //
        // 例:
        // 『注文の多い料理店』（既定値）
        // https://www.aozora.gr.jp/cards/000081/files/43754_17659.html
        // 『吾輩は猫である』約2000分割
        // https://www.aozora.gr.jp/cards/000148/files/789_14547.html
        // 『河童』 subtitle = どうか Kappa と発音してください。
        // https://www.aozora.gr.jp/cards/000879/files/69_14933.html
        // 『マッチ売りの少女』　original_title = THE LITTLE MATCH-SELLER
        // https://www.aozora.gr.jp/cards/000019/files/194_23024.html
        // 『盲腸』 指定 Attribute (属性) 無し
        //  https://www.aozora.gr.jp/cards/000168/files/3626.html
        //
        private async Task RetrieveHtml()
        {
            // URL 取得
            string url = textBox1.Text;

            // ページを読み込む
            webBrowser1.Navigate(url);

            // DocumentCompleted イベントを待つ
            await Task.Run(() =>
            {
                var tcs = new TaskCompletionSource<bool>();
                WebBrowserDocumentCompletedEventHandler handler = null;
                handler = (s, args) =>
                {
                    tcs.SetResult(true);
                    //browser.DocumentCompleted -= handler;
                    webBrowser1.DocumentCompleted -= handler;
                };
                //browser.DocumentCompleted += handler;
                webBrowser1.DocumentCompleted += handler;
                tcs.Task.Wait();
            });
        }

        //
        // 青空文庫のみ 
        //
        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            //Uri uri = e.Url;

            // 許可された URL かチェック
            string url = e.Url.ToString();

            bool startsWithAllowed = _allowedStartWith.Any(pattern => Regex.IsMatch(url, pattern, RegexOptions.IgnoreCase));

            if (startsWithAllowed)
            {
                if (isIntroduce)
                {
                    isIntroduce = false;
                }
                else
                {

                    // ページ範囲の無限ループ回避用
                    firstPageReset = true;
                    lastPageReset = true;

                    // ボタンを無効化し、「⌛」を表示
                    isProcessingTrue();
                    label3.Text = "✖";

                    textBrowser.Navigate(url);
                }
            }
            else
            {
                // 許可されていないドメインへのナビゲーションをキャンセル
                e.Cancel = true;
                textBox2.Text = "⚠️許可範囲外⚠️";
                textBox3.Text = "⚠️リンク先: " + url + " ⚠️";
            }
        }

        //
        // ポップアップウィンドウが開くのを防ぐ
        //
        private void webBrowser1_NewWindow(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // 新しいウィンドウが開くのを防ぐ
            // 新しいウィンドウのURLは取得不能
            // 新しいウィンドウが開いたら、aozoraOSC からは制御不能
            e.Cancel = true;
            textBox2.Text = "⚠️新しいウィンドウをキャンセル⚠️";
            textBox3.Text = "⚠️右クリックからリンクをコピーして⚠️";
            textBox4.Text = "⚠️URL 欄に貼り付けてください⚠️";
        }

        //
        // 読み込んだ情報から必要な情報を抽出し
        // 文字列処理や本文分割等を実施
        //
        private void textBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            // WebBrowser コントロールの Document プロパティを取得
            WebBrowser browser = new WebBrowser();
            browser = (WebBrowser)sender;

            textBox1.Text = browser.Url.ToString();

            HtmlDocument doc;

            if (browser.Document == null)
            {
                // null対応
                doc = nullBrowser.Document;
            }
            else
            {
                doc = browser.Document;
            }

            // 要素を初期化（情報がなかった時の既定値）
            string titleText = "（作品名：情報無し）";
            string subtitleText = "";
            string original_titleText = "";
            string authorText = "（著者名：情報無し）";
            string translatorText = "（翻訳者名：情報無し）";
            string main_textText = "（本文：情報無し）";
            string biblioText = "";
            string after_textText = "";

            //  h1: "title"
            //  h2: "subtitle",  "original_title", "author", "translator"
            // div: "main_text",  "bibliographical_information", "after_text"
            HtmlElementCollection h1Elements = doc.GetElementsByTagName("h1");
            HtmlElementCollection h2Elements = doc.GetElementsByTagName("h2");
            HtmlElementCollection divElements = doc.GetElementsByTagName("div");

            // 指定 Attribute (属性) 無し
            // title: <title>
            //  body: <body>
            //  meta: <meta name = "DC.Title, DC.Creator... 必ずあるわけではない
            HtmlElementCollection titleElements = doc.GetElementsByTagName("title");
            HtmlElementCollection bodyElements = doc.GetElementsByTagName("body");
            HtmlElementCollection metaElements = doc.GetElementsByTagName("meta");

            // "h1" のすべての要素を取得
            foreach (HtmlElement element in h1Elements)
            {
                // 属性 className が "title" である要素を探す
                if (element.GetAttribute("className") == "title")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        titleText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                    }
                    break; // ループを終了
                }
            }

            // <h1 class="title"> が無かったら <meta name="DC.Title" content="作品名" /> から取得を試みる
            if (titleText == "（作品名：情報無し）")
            {
                // "<meta>" を取得
                foreach (HtmlElement element in metaElements)
                {
                    string name = element.GetAttribute("name");
                    if (name.Equals("DC.Title", StringComparison.OrdinalIgnoreCase))
                    {
                        // "DC.Title" タグが見つかった場合、content 属性を取得
                        string innerText = element.GetAttribute("content");
                        if (innerText != null)
                        {
                            titleText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                        }
                        break;
                    }
                }
            }

            // <h1 class="title"> が無く、
            // <meta name="DC.Title" content="作品名" /> も無かったら
            // <title> から取得を試みる
            if (titleText == "（作品名：情報無し）")
            {
                // "<title>" を取得
                foreach (HtmlElement element in titleElements)
                {
                    // テキストを取得
                    string innerText = element.InnerText;
                    if (innerText != null)
                    {
                        titleText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                    }
                    break;
                }
            }


            // "h2" のすべての要素を取得
            foreach (HtmlElement element in h2Elements)
            {
                // 属性 className が "subtitle" である要素を探す
                if (element.GetAttribute("className") == "subtitle")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除後 "\n" を１つ追加
                        subtitleText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                        subtitleText = "\n" + subtitleText;
                    }
                    break; // ループを終了
                }
            }

            // "h2" のすべての要素を取得
            foreach (HtmlElement element in h2Elements)
            {
                // 属性 className が "original_title" である要素を探す
                if (element.GetAttribute("className") == "original_title")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除後 "\n" を１つ追加
                        original_titleText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                        original_titleText = "\n" + original_titleText;
                    }
                    break; // ループを終了
                }
            }


            // "h2" のすべての要素を取得
            foreach (HtmlElement element in h2Elements)
            {
                // 属性 className が "author" である要素を探す
                if (element.GetAttribute("className") == "author")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        authorText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                    }
                    break; // ループを終了
                }
            }

            // <h2 class="author"> が無かったら <meta name="DC.Creator" content="著者名" /> から取得を試みる
            if (authorText == "（著者名：情報無し）")
            {
                // "<meta>" を取得
                foreach (HtmlElement element in metaElements)
                {
                    string name = element.GetAttribute("name");
                    if (name.Equals("DC.Creator", StringComparison.OrdinalIgnoreCase))
                    {
                        // "DC.Creator" タグが見つかった場合、content 属性を取得
                        string innerText = element.GetAttribute("content");
                        if (innerText != null)
                        {
                            authorText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                        }
                        break;
                    }
                }
            }

            // <h2 class="author"> が無く、
            // <meta name="DC.Creator" content="著者名" /> も無かったら
            // <h2> から取得を試みる
            if (authorText == "（著者名：情報無し）")
            {
                // "<h2>" を取得
                foreach (HtmlElement element in h2Elements)
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        authorText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                    }
                    break;
                }
            }


            // class 名が "translator" のすべての要素を取得
            foreach (HtmlElement element in h2Elements)
            {
                // class 属性が "translator" である要素を探す
                if (element.GetAttribute("className") == "translator")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        translatorText = innerText.TrimStart('\r', '\n'); // 念のため just in case
                    }
                    break; // ループを終了
                }
            }


            // "div" のすべての要素を取得
            foreach (HtmlElement element in divElements)
            {
                // 属性 className が "main_text" である要素を探す
                if (element.GetAttribute("className") == "main_text")
                {
                    // Create a new HTML string to replace <img> tags
                    string newHtml = element.InnerHtml;

                    // Find all <img> tags and replace them
                    HtmlElementCollection imgElements = element.GetElementsByTagName("img");
                    foreach (HtmlElement img in imgElements)
                    {
                        if (img.GetAttribute("className") == "gaiji")
                        {
                            string altText = img.GetAttribute("alt");
                            newHtml = newHtml.Replace(img.OuterHtml, $"▒［外字{altText}］");
                        }
                        else
                        {
                            string altText = img.GetAttribute("alt");
                            newHtml = newHtml.Replace(img.OuterHtml, $"🖼［{altText}］");
                        }
                    }

                    // Update the element's inner HTML
                    element.InnerHtml = newHtml;

                    // Get the updated text
                    string updatedText = element.InnerText;

                    if (updatedText != null)
                    {
                        // 先頭の改行を削除
                        main_textText = updatedText.TrimStart('\r', '\n');

                        // 本文終了：底本との仕切り
                        main_textText += "\r\n\r\n\r\n――――――――――――――――――――\r\n\r\n";
                    }
                    break; // ループを終了
                }
            }

            // <div class="main_text"> が無かったら <body> から取得を試みる
            if (main_textText == "（本文：情報無し）")
            {
                // "<body>" を取得
                foreach (HtmlElement element in bodyElements)
                {
                    // Create a new HTML string to replace <img> tags
                    string newHtml = element.InnerHtml;

                    // Find all <img> tags and replace them
                    HtmlElementCollection imgElements = element.GetElementsByTagName("img");
                    foreach (HtmlElement img in imgElements)
                    {
                        if (img.GetAttribute("className") == "gaiji")
                        {
                            string altText = img.GetAttribute("alt");
                            newHtml = newHtml.Replace(img.OuterHtml, $"▒［外字{altText}］");
                        }
                        else
                        {
                            string altText = img.GetAttribute("alt");
                            newHtml = newHtml.Replace(img.OuterHtml, $"🖼［{altText}］");
                        }
                    }

                    // Update the bodyElement's inner HTML
                    element.InnerHtml = newHtml;

                    // Get the updated text
                    string updatedText = element.InnerText;

                    if (updatedText != null)
                    {
                        main_textText = updatedText.TrimStart('\r', '\n');
                    }
                    break;
                }
            }

            // class 名が "bibliographical_information" のすべての要素を取得
            foreach (HtmlElement element in divElements)
            {
                // class 属性が "bibliographical_information" である要素を探す
                if (element.GetAttribute("className") == "bibliographical_information")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        biblioText = innerText.TrimStart('\r', '\n');
                    }
                    break; // ループを終了
                }
            }

            // class 名が "after_text" のすべての要素を取得
            foreach (HtmlElement element in divElements)
            {
                // class 属性が "after_text" である要素を探す
                if (element.GetAttribute("className") == "after_text")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        after_textText = innerText.TrimStart('\r', '\n');
                    }
                    break; // ループを終了
                }
            }


            // 作品名をまとめて表示
            string totalTitleText = titleText + subtitleText + original_titleText;

            // 表示幅が 40 を超える時の対応
            string resultTitleText = totalTitleText;
            string resultAuthorText = authorText;
            string resultTranslatorText = translatorText;

            // 本文をまとめて表示
            string mainText = main_textText + biblioText + after_textText;


            // 表示幅が 40 を超える時は「…」を追加し以降の文字列削除
            // 作品名
            if (GetWidth(totalTitleText) > 40)
            {
                resultTitleText = TruncateWithEllipsis(totalTitleText, 41);
            }
            textBox2.Text = resultTitleText;

            // 著者名
            if (GetWidth(authorText) > 40)
            {
                resultAuthorText = TruncateWithEllipsis(authorText, 41);
            }
            textBox3.Text = resultAuthorText;

            // 翻訳者名
            if (GetWidth(translatorText) > 40)
            {
                resultTranslatorText = TruncateWithEllipsis(translatorText, 41);
            }
            textBox4.Text = resultTranslatorText;

            // 本文
            textBox5.Text = mainText;

            // 文字列操作の関係から "\r" CR: Carriage Return を削除（改行は "\n" のみとする）
            // 全角基準で 20, 21, 22 文字目を禁則処理の対象とするため、引数は 19 * 2 = 38 を渡す
            string mainTextWithoutCR = mainText.Replace("\r", "");
            string formattedText = AddLineBreaks(mainTextWithoutCR, 38);

            // タグ付け等
            bool isMainText = true; // 連続タグの削除（ metaInfo では連続タグを削除しない）
            string result = AddTagsToSpecialCharacters(formattedText, isMainText);

            // ページ作成: 10 行ごとに配列に追加
            string[] tempChunks = SplitIntoLines(result);

            // 最初のページ: 全角空白追加
            string[] chunks = new string[tempChunks.Length + 1];
            chunks[0] = "　";
            Array.Copy(tempChunks, 0, chunks, 1, tempChunks.Length); // Array.Copy(元, 元開始位置, 先, 先開始位置, 個数)

            // 本文が分割格納されている配列を上書き
            pages = chunks;


            // 送信ページ範囲の再設定
            // ドロップダウンリストの消去
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();

            // ドロップダウンリストの作成（全ページ数）
            for (int i = 0; i < chunks.Length; i++)
            {
                int pageNumber = i + 1;
                comboBox1.Items.Add(pageNumber.ToString());
                comboBox2.Items.Add(pageNumber.ToString());
            }
            // 全ページ数で再設定し最後のページを選択
            // firstPage = 1; // 最初のページは初期設定から変更なし
            lastPage = chunks.Length;
            comboBox1.SelectedItem = firstPage.ToString();
            comboBox2.SelectedItem = lastPage.ToString();

            // ボタンを再度有効化し、テキストを戻す
            isProcessingFalse();
        }

        //
        // 文字列の幅を計算
        //
        public static int GetWidth(string text)
        {
            int width = 0;
            for (int i = 0; i < text.Length; i++)
            {
                string character = text[i].ToString();
                width += GetCharacterWidth(character);
            }
            return width;
        }

        //
        // 文字列を指定の幅で切り取り「…」を追加
        //
        public static string TruncateWithEllipsis(string text, int maxWidth)
        {
            int width = 0;
            StringBuilder result = new StringBuilder();
            foreach (char c in text)
            {
                string character = c.ToString();
                int charWidth = GetCharacterWidth(character);
                if (width + charWidth > maxWidth - 1)
                {
                    result.Append("…");
                    break;
                }
                result.Append(c);
                width += charWidth;
            }
            return result.ToString();
        }

        //
        // 表示幅を取得
        //
        static int GetCharacterWidth(string character)
        {
            // 文字の表示幅を取得する関数
            if (character == "\n")
            {
                return 1; // 改行文字の幅は1とする
            }
            else if (character.Length > 1)  // サロゲートペアの場合
            {
                return 2; // サロゲートペアの幅は2とする
            }
            else if (IsFullWidth(character)) // 全角文字の場合
            {
                return 2; // 全角文字の幅は2とする
            }
            else // それ以外（半角文字など）
            {
                return 1; // 半角文字の幅は1とする
            }
        }

        //
        // 全角文字の判定
        //
        static bool IsFullWidth(string input)
        {
            // 文字列のバイト数を調べ、 2 バイト以上であれば全角文字と判定する
            return Encoding.UTF8.GetByteCount(input) >= 2;
        }


        //
        // 20文字基準で禁則処理を行い改行を追加
        //
        private string AddLineBreaks(string text, int maxLineLength)
        {
            var result = new StringBuilder();
            int index = 0;
            int length = text.Length;
            int lineLength = 0;

            while (index < length)
            {
                // 改行コードが途中見つかった場合は、その位置で改行を挿入
                if (IsLF(text[index]))
                {
                    result.Append('\n');
                    lineLength = 0; // 行の文字数をリセット
                    index++;
                    continue; // 以下のコードをスキップ
                }

                // サロゲートペアを考慮して次の文字、20文字目を取得
                int charLength = char.IsSurrogatePair(text, index) ? 2 : 1;
                string character = text.Substring(index, charLength);
                // 表示幅を取得
                int characterWidth = GetCharacterWidth(character);

                // 20文字目で判断
                if (lineLength + characterWidth > maxLineLength)
                {
                    // 21文字目を取得
                    int nextCharLength = 0;
                    string nextCharacter = "";
                    if (index + charLength < length)
                    {
                        nextCharLength = char.IsSurrogatePair(text, index + charLength) ? 2 : 1;
                        nextCharacter = text.Substring(index + charLength, nextCharLength);
                    }
                    else // 21文字目が無かったら（ちょうど 20の倍数で mainText が終わったとき）
                    {
                        result.Append(character);
                        break;
                    }

                    // 20文字目が始め括弧なら次の行にする（その行は19文字）
                    if (IsOpeningBracket(character[0]))
                    {
                        result.Append('\n');
                        result.Append(character);
                        lineLength = characterWidth; // 行の文字数をリセット
                    }
                    else if (IsClosingBracket(nextCharacter[0])) // 21文字目が句読点や終わり括弧の時
                    {
                        if (index + charLength + nextCharLength < length) // 22文字目がある
                        {
                            // 『The Affair of Two Watches』　谷崎潤一郎
                            // https://www.aozora.gr.jp/cards/001383/files/58153_71313.html
                            // p.2 何でも十二月の末の、とある夕暮の事だった。   [20字 + "。" + \n]
                            // p.8 「………飲みたいなあ。お互に血の出るよう
                            //     な冗談を云うたって仕様がない。え、杉さん。」 [20字 + "。」" + \n]
                            //
                            // 22文字目を取得
                            int afterNextCharLength = char.IsSurrogatePair(text, index + charLength + nextCharLength) ? 2 : 1;
                            string afterNextCharacter = text.Substring(index + charLength + nextCharLength, afterNextCharLength);

                            if (IsClosingBracket(afterNextCharacter[0])) // 22文字目も句読点や終わり括弧の場合
                            {
                                result.Append(character);
                                result.Append(nextCharacter);
                                result.Append(afterNextCharacter);
                                if (index + charLength + nextCharLength + afterNextCharLength < length) // 23文字目がある
                                {
                                    // 23文字目を取得
                                    int theThirdOneAheadCharLength = char.IsSurrogatePair(text, index + charLength + nextCharLength + afterNextCharLength) ? 2 : 1;
                                    string theThirdOneAheadCharacter = text.Substring(index + charLength + nextCharLength + afterNextCharLength, theThirdOneAheadCharLength);
                                    if (!IsLF(theThirdOneAheadCharacter[0])) // 23文字目が改行コード
                                    {
                                        result.Append('\n');
                                    }
                                }
                                lineLength = 0; // 行の文字数をリセット
                                index += charLength + nextCharLength + afterNextCharLength;
                                continue; // 以下のコードをスキップ
                            }
                            else if (IsLF(afterNextCharacter[0])) // 22文字目が改行コード
                            {
                                result.Append(character);
                                result.Append(nextCharacter);
                                result.Append(afterNextCharacter); // 改行コード
                                // result.Append('\n');
                                lineLength = 0; // 行の文字数をリセット
                                index += charLength + nextCharLength + afterNextCharLength;
                                continue; // 以下のコードをスキップ
                            }
                        }
                        result.Append(character);
                        result.Append(nextCharacter);
                        result.Append('\n');
                        lineLength = 0; // 行の文字数をリセット
                        index += charLength + nextCharLength;
                        continue; // 以下のコードをスキップ
                    }
                    else if (IsLF(nextCharacter[0])) // 21文字目が改行コード
                    {
                        result.Append(character);
                        lineLength = 0; // 行の文字数をリセット
                    }
                    else // 一般的な処理
                    {
                        result.Append(character);
                        result.Append('\n');
                        lineLength = 0; // 行の文字数をリセット
                    }
                }
                else
                {
                    // 文字を追加
                    result.Append(character);
                    lineLength += characterWidth;
                }

                // インデックスを進める
                index += charLength;
            }

            return result.ToString();
        }

        //
        // 始め括弧の文字かどうかを判定する
        // 参考 https://ja.wikipedia.org/wiki/%E7%A6%81%E5%89%87%E5%87%A6%E7%90%86
        //
        static bool IsOpeningBracket(char c)
        {
            // 始め括弧類
            // c == '‘' || c == '“' || 始めと終わりの区別に難あり
            return c == '「' || c == '『'
                || c == '【' || c == '〖' || c == '〔' || c == '〘'
                || c == '〈' || c == '《' || c == '｟' || c == '«'
                || c == '（' || c == '［' || c == '｛' // 全角
                || c == '(' || c == '[' || c == '{'; // 半角
        }

        //
        // 句読点や終わり括弧の文字かどうかを判定する
        // 参考 https://ja.wikipedia.org/wiki/%E7%A6%81%E5%89%87%E5%87%A6%E7%90%86
        //
        static bool IsClosingBracket(char c)
        {
            // 終わり括弧類
            // c == '’' || c == '”' || 始めと終わりの区別に難あり
            return c == '」' || c == '』'
                || c == '】' || c == '〗' || c == '〕' || c == '〙'
                || c == '〉' || c == '》' || c == '｠' || c == '»'
                || c == '）' || c == '］' || c == '｝' // 全角
                || c == ')' || c == ']' || c == '}'  // 半角
                                                     // 行頭禁則和字
                || c == 'ゝ' || c == 'ゞ'
                || c == 'ー' // カタカナの長音
                || c == 'ァ' || c == 'ィ' || c == 'ゥ' || c == 'ェ' || c == 'ォ'
                || c == 'ッ' || c == 'ャ' || c == 'ュ' || c == 'ョ' || c == 'ヮ' || c == 'ヵ' || c == 'ヶ'
                || c == 'ぁ' || c == 'ぃ' || c == 'ぅ' || c == 'ぇ' || c == 'ぉ'
                || c == 'っ' || c == 'ゃ' || c == 'ゅ' || c == 'ょ' || c == 'ゎ' || c == 'ゕ' || c == 'ゖ'
                || c == 'ㇰ' || c == 'ㇱ' || c == 'ㇲ' || c == 'ㇳ' || c == 'ㇴ'
                || c == 'ㇵ' || c == 'ㇶ' || c == 'ㇷ' || c == 'ㇸ' || c == 'ㇹ'
                || c == 'ㇺ' || c == 'ㇻ' || c == 'ㇼ' || c == 'ㇽ' || c == 'ㇾ' || c == 'ㇿ'
                || c == '々' || c == '〻'
                // ハイフン類：行頭に来る可能性あり? 例: "―" （ダッシュ）はある
                // || c == '‐' || c == '゠' || c == '–' || c == '〜' || c == '～'
                // 区切り約物
                || c == '？' || c == '！' || c == '‼' || c == '⁇' || c == '⁈' || c == '⁉'
                || c == '?' || c == '!'  //半角
                                         //中点類
                || c == '：' || c == '；' || c == '／' || c == '・'
                || c == ':' || c == ';' || c == '/'
                // 句点類
                || c == '，' || c == '．' // 全角
                || c == ',' || c == '.'  // 半角
                || c == '、' || c == '。' || c == '〟';
        }

        //
        // 改行コードかどうかを判定する
        //
        static bool IsLF(char c) // Line Feed
        {
            return c == '\n';
        }

        //
        // サロゲートペア、横文字、拗音等にタグ付け
        // 句読点変換
        // 連続タグ削除（本文）
        //
        private string AddTagsToSpecialCharacters(string input, bool tagsDelete)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < input.Length; i++)
            {
                char currentChar = input[i];

                // Check if current character is a high surrogate
                if (char.IsHighSurrogate(currentChar) && i + 1 < input.Length)
                {
                    char nextChar = input[i + 1];
                    // Check if the next character is a low surrogate
                    if (char.IsLowSurrogate(nextChar))
                    {
                        // Append tags around the surrogate pair
                        sb.Append("<size=67%>");
                        sb.Append(currentChar);
                        sb.Append(nextChar);
                        sb.Append("</size>");
                        i++; // Skip the low surrogate since we've already processed it
                    }
                    else
                    {
                        // Append just the high surrogate if not followed by a valid low surrogate
                        sb.Append(currentChar);
                    }
                }
                //
                // 参考 unicode 一覧
                // https://ja.wikipedia.org/wiki/Unicode%E4%B8%80%E8%A6%A7%E8%A1%A8
                // 横文字判定
                else if ((currentChar >= 0x0020 && currentChar <= 0x007E)
                      || (currentChar >= 0x00A1 && currentChar <= 0x00AC)
                      || (currentChar >= 0x00AE && currentChar <= 0x034E)
                      || (currentChar >= 0x0350 && currentChar <= 0x1FFF)
                      || (currentChar >= 0x2010 && currentChar <= 0x2027)
                      || (currentChar >= 0x2030 && currentChar <= 0x2065)
                      || (currentChar >= 0x2070 && currentChar <= 0x214F)
                      || (currentChar >= 0x2180 && currentChar <= 0x2319)
                      || (currentChar >= 0x231C && currentChar <= 0x23EF)
                      || (currentChar >= 0x23F4 && currentChar <= 0x245F)
                      || (currentChar >= 0x2500 && currentChar <= 0x25FF)
                      || (currentChar >= 0x261A && currentChar <= 0x261F)
                      || (currentChar >= 0x2768 && currentChar <= 0x2775)
                      || (currentChar >= 0x2794 && currentChar <= 0x27FF)
                      || (currentChar >= 0x2900 && currentChar <= 0x2E7F)
                      || (currentChar >= 0x3008 && currentChar <= 0x301C)
                      || (currentChar == 0x3030)
                      || (currentChar == 0x30FC)
                      || (currentChar >= 0x3371 && currentChar <= 0x33DF)
                      || (currentChar == 0x33FF)
                      || (currentChar >= 0xA000 && currentChar <= 0xABFF)
                      || (currentChar >= 0xFB00 && currentChar <= 0xFF60)
                      || (currentChar >= 0xFF62 && currentChar <= 0xFF63)
                      || (currentChar == 0xFF66)
                      || (currentChar >= 0xFF70 && currentChar <= 0xFF9F)
                      || (currentChar >= 0xFFE0 && currentChar <= 0xFFFF)
                      || (currentChar >= 0x10000 && currentChar <= 0x1EFFF)
                      || (currentChar >= 0x1F100 && currentChar <= 0x1F1FF)
                      || (currentChar >= 0x1F519 && currentChar <= 0x1F51D)
                      || (currentChar >= 0x1F597 && currentChar <= 0x1F5A3)
                      || (currentChar >= 0x1F5DA && currentChar <= 0x1F5DB)
                      || (currentChar >= 0x1F66C && currentChar <= 0x1F66F)
                      || (currentChar >= 0x1F800 && currentChar <= 0x1F8FF))
                {
                    // Append tags for characters within specified Unicode ranges
                    sb.Append("<rotate=0>");
                    sb.Append(currentChar);
                    sb.Append("<rotate=90>");
                }
                // 拗音・促音等、小さい文字判定
                else if (currentChar == 0x3041 || currentChar == 0x3043 || currentChar == 0x3045 || currentChar == 0x3047 || currentChar == 0x3049
                      || currentChar == 0x3063
                      || currentChar == 0x3083 || currentChar == 0x3085 || currentChar == 0x3087
                      || currentChar == 0x308E || currentChar == 0x3095 || currentChar == 0x3096
                      || currentChar == 0x30A1 || currentChar == 0x30A3 || currentChar == 0x30A5 || currentChar == 0x30A7 || currentChar == 0x30A9
                      || currentChar == 0x30C3
                      || currentChar == 0x30E3 || currentChar == 0x30E5 || currentChar == 0x30E7
                      || currentChar == 0x30EE || currentChar == 0x30F5 || currentChar == 0x30F6
                      || (currentChar >= 0xFF67 && currentChar <= 0xFF6F))
                {
                    // Append tags for characters within specified Unicode ranges
                    sb.Append("<voffset=0.2em>");
                    sb.Append(currentChar);
                    sb.Append("</voffset>");
                }
                // 読点
                else if (currentChar == 0x3001 || currentChar == 0xFF64)
                {
                    sb.Append("` ");
                }
                // 句点
                else if (currentChar == 0x3002 || currentChar == 0xFF61)
                {
                    sb.Append("゜");
                }
                else
                {
                    // Append the character as it is if it does not meet the above conditions
                    sb.Append(currentChar);
                }
            }

            string ret = sb.ToString();

            // 本文の連続タグ削除（ metaInfo は削除しない）
            if (tagsDelete)
            {
                ret = ret.Replace("<rotate=90><rotate=0>", "");
            }

            // 句点＋閉じ
            ret = ret.Replace("゜<rotate=0>」<rotate=90>", "゜<space=-0.5em><rotate=0>」<rotate=90>");
            ret = ret.Replace("゜<rotate=0>』<rotate=90>", "゜<space=-0.5em><rotate=0>』<rotate=90>");
            ret = ret.Replace("゜<rotate=0>）<rotate=90>", "゜<space=-0.5em><rotate=0>）<rotate=90>");

            return ret;
        }

        //
        // ページ作成: 10 行ごとに配列に追加
        //
        private string[] SplitIntoLines(string text)
        {
            // 改行で分割して配列に格納
            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            // 10 行ごとに配列に格納するためのリスト
            var result = new List<string>();

            // 10 行ごとにループして配列に追加
            for (int i = 0; i < lines.Length; i += 10)
            {
                int count = Math.Min(10, lines.Length - i); // 残りの行数が10行未満の場合も考慮
                string[] tempArray = new string[count];
                Array.Copy(lines, i, tempArray, 0, count);

                // ページの頭にタグ「 <rotate=90> 」を追加
                string rotatedChunk = string.Join(Environment.NewLine, tempArray);
                result.Add($"<rotate=90>{rotatedChunk}");
            }

            return result.ToArray();
        }


        //
        // ［VirtualCastへ送信］ボタンを押したとき
        //
        private async Task SelectDataForSending()
        {
            // OSC で送る metaInfo の情報まとめ
            bool isMainText = false; // 連続タグを削除しない
            string sendTitle = "<rotate=90>" + AddTagsToSpecialCharacters(textBox2.Text, isMainText);
            string sendAuthor = "<rotate=90>" + AddTagsToSpecialCharacters(textBox3.Text, isMainText);
            string sendTranslator = "<rotate=90>" + AddTagsToSpecialCharacters(textBox4.Text, isMainText);

            int sendSelectedFirstPage = Convert.ToInt32(comboBox1.SelectedItem);
            int sendSelectedLastPage = Convert.ToInt32(comboBox2.SelectedItem);

            int sendPages = sendSelectedLastPage - sendSelectedFirstPage + 1; // 全ページ数

            string sendVersion = aozoraVersion;

            // metaInfo を OSC で送る
            SendOscMetaInfo(sendTitle, sendAuthor, sendTranslator, sendPages, sendVersion);



            // for debugging purposes
            // string path = @"C:\example\";
            // File.WriteAllText(path + "SendOscMetaInfo.txt", sendTitle + "\n" + sendAuthor + "\n" + sendTranslator + "\n" + sendPages.ToString() + "\n" + sendVersion);



            // 過密にならないよう、40ミリ秒待機(FPS25?) just in case
            await Task.Delay(40);


            // 分割テキストを OSC で送信
            string[] chunks = pages; // 本文全文を代入

            // OSC 中止用
            label8.Text = "✖";

            // 選択されたページ範囲
            for (int i = sendSelectedFirstPage - 1; i < sendSelectedLastPage; i++)
            {

                // 中止
                if (isBreak)
                    break;


                // ページ番号と分割テキスト
                int sendPageNumber = i - sendSelectedFirstPage + 2;
                string sendPartText = chunks[i];

                // ページ番号と分割テキストを OSC で送信
                SendOscPartText(sendPageNumber, sendPartText);



                // for debugging purposes
                // File.WriteAllText(path + sendPageNumber.ToString() + ".txt", sendPartText);



                // 送信中の進捗表示 "🍣 1/42" -> "🍣 42/42"
                button2.Text = "🍣" + sendPageNumber.ToString() + "/" + sendPages.ToString();

                // 過密にならないよう、40ミリ秒待機(FPS25?) just in case
                await Task.Delay(40);
            }
            // 連打防止用
            await Task.Delay(250);
        }


        // OSC で送信
        // SharpOSC (MIT license) を部分利用
        // https://github.com/ValdemarOrn/SharpOSC
        //
        // 作品名、著者名、翻訳者名、本文分割数を送る
        //
        static void SendOscMetaInfo(string title, string author, string translator, int pages, string version)
        {
            try
            {
                Encoding utf8 = Encoding.UTF8;
                byte[] blob_title = utf8.GetBytes(title);
                byte[] blob_author = utf8.GetBytes(author);
                byte[] blob_translator = utf8.GetBytes(translator);
                int int_pages = pages;
                byte[] blob_version = utf8.GetBytes(version);

                var message = new OscMessage("/Taki/aozoraReader/metaInfo", blob_title, blob_author, blob_translator, int_pages, blob_version);
                var sender = new UDPSender("127.0.0.1", 19100);
                sender.Send(message);
            }
            // for debugging purposes
            //catch (Exception error)
            catch (Exception)
            {
                // for debugging purposes
                // MessageBox.Show(error.ToString());
            }
        }

        // OSC で送信
        // SharpOSC (MIT license) を部分利用
        // https://github.com/ValdemarOrn/SharpOSC
        //
        // ページ番号と分割テキストを OSC で送信
        //
        static void SendOscPartText(int pageNumber, string partText)
        {
            try
            {
                Encoding utf8 = Encoding.UTF8;
                int int_pageNumber = pageNumber;
                byte[] blob_partText = utf8.GetBytes(partText);

                var message = new OscMessage("/Taki/aozoraReader/partText", int_pageNumber, blob_partText);
                var sender = new UDPSender("127.0.0.1", 19100);
                sender.Send(message);
            }
            // for debugging purposes
            //catch (Exception error)
            catch (Exception)
            {
                // for debugging purposes
                // MessageBox.Show(error.ToString());
            }
        }


        //
        // 送信ページの範囲指定（最初が変更されたら）
        //
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (firstPageReset == true) // 無限ループ回避用
            {
                firstPageReset = false;
            }
            else
            {
                lastPageReset = true; // 無限ループ回避用

                int currentFirstPage = Convert.ToInt32(comboBox1.SelectedItem); // string を int に変換
                int currentLastPage = Convert.ToInt32(comboBox2.SelectedItem);

                comboBox2.SuspendLayout(); // comboBox2 の描画更新を一時停止
                try
                {
                    // 起動時は 0 が入る
                    if (currentLastPage == 0)
                    {
                        currentLastPage = 2;
                    }

                    // ドロップダウンリストの再設定
                    // 指定ページから最後のページまで
                    comboBox2.Items.Clear();
                    for (int i = currentFirstPage; i <= lastPage; i++)
                    {
                        comboBox2.Items.Add(i.ToString());
                    }

                    // 念のため just in case
                    if (currentLastPage < currentFirstPage)
                    {
                        currentLastPage = currentFirstPage;
                    }

                    comboBox2.SelectedItem = currentLastPage.ToString();
                }
                finally
                {
                    comboBox2.ResumeLayout(true); // 描画更新を再開し、変更を適用
                }
            }
        }

        //
        // 送信ページの範囲指定（最後が変更されたら）
        //
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lastPageReset == true) // 無限ループ回避用
            {
                lastPageReset = false;
            }
            else
            {
                firstPageReset = true; // 無限ループ回避用

                int currentFirstPage = Convert.ToInt32(comboBox1.SelectedItem); // string を int に変換
                int currentLastPage = Convert.ToInt32(comboBox2.SelectedItem);

                comboBox1.SuspendLayout(); // comboBox1 の描画更新を一時停止
                try
                {
                    // ドロップダウンリストの再設定
                    // 最初のページから指定ページまで
                    comboBox1.Items.Clear();
                    for (int i = firstPage; i <= currentLastPage; i++)
                    {
                        comboBox1.Items.Add(i.ToString());
                    }

                    // 念のため just in case
                    if (currentFirstPage > currentLastPage)
                    {
                        currentFirstPage = currentLastPage;
                    }

                    comboBox1.SelectedItem = currentFirstPage.ToString();
                }
                finally
                {
                    comboBox1.ResumeLayout(true); // 描画更新を再開し、変更を適用
                }
            }
        }

        // フォームを閉じたとき
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (webBrowser1 != null)
            {
                webBrowser1.Dispose();
                webBrowser1 = null;
            }
        }

        //
        // ファイルとフォルダの D&D
        // リンクの D&D はできなかった (JavaScript 組み込みでもだめだった)
        //
        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグされているデータがファイルまたはフォルダであるか確認
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            // ドロップされたデータがファイルまたはフォルダであるか確認
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (paths.Length > 0)
                {
                    // 最初のパスがファイルまたはフォルダであるかを確認
                    string path = paths[0];

                    if (File.Exists(path) || Directory.Exists(path))
                    {
                        // ファイルまたはフォルダのパスを textBox1 に表示
                        textBox1.Text = path;
                    }
                    else
                    {
                        textBox1.Text = "ドラッグ＆ドロップされたアイテムは存在しません。";
                    }
                }
            }
        }

        private async void label8_Click(object sender, EventArgs e)
        {
            isBreak = true;

            await Task.Delay(100);
            label8.Text = "";
            isBreak = false;
        }

        // Navigates webBrowser1 to the previous page in the history.
        private void backButton_Click(object sender, EventArgs e)
        {
            // 処理中でないことを確認
            if (isProcessing)
                return;

            webBrowser1.GoBack();
        }

        // Disables the Back button at the beginning of the navigation history.
        private void webBrowser1_CanGoBackChanged(object sender, EventArgs e)
        {
            backButton.Enabled = webBrowser1.CanGoBack;
        }

        // Navigates webBrowser1 to the next page in history.
        private void forwardButton_Click(object sender, EventArgs e)
        {
            // 処理中でないことを確認
            if (isProcessing)
                return;

            webBrowser1.GoForward();
        }

        // Disables the Forward button at the end of navigation history.
        private void webBrowser1_CanGoForwardChanged(object sender, EventArgs e)
        {
            forwardButton.Enabled = webBrowser1.CanGoForward;
        }

        private void homeButton_Click(object sender, EventArgs e)
        {
            // 処理中でないことを確認
            if (isProcessing)
                return;

            // ページを読み込む
            webBrowser1.Navigate(homeUrl);
        }

        private void buttonGrayout_Paint(object sender, PaintEventArgs e)
        {
            Button btn = (Button)sender;
            Color textColor;

            // ボタンが有効かどうかでテキストの色を決定
            if (btn.Enabled)
            {
                textColor = btn.ForeColor; // 通常時は設定されているForeCrolorを使用
            }
            else
            {
                textColor = Color.DimGray; // 無効時は灰色
            }

            // テキスト描画の準備
            TextFormatFlags flags = TextFormatFlags.HorizontalCenter |
                                    TextFormatFlags.VerticalCenter |
                                    TextFormatFlags.WordBreak; // 必要に応じて改行などを考慮

            // ボタンの背景を描画 (オプション: デフォルトの背景が不要な場合)
            // e.Graphics.FillRectangle(new SolidBrush(btn.BackColor), btn.ClientRectangle);

            // テキストを描画
            TextRenderer.DrawText(e.Graphics, btn.Text, btn.Font, btn.ClientRectangle, textColor, flags);

            // フォーカスがある場合の描画 (オプション)
            if (btn.Focused && btn.Enabled)
            {
                ControlPaint.DrawFocusRectangle(e.Graphics, btn.ClientRectangle);
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {
            webBrowser1.Stop();
            textBrowser.Stop();
            label3.Text = "";
            isProcessingFalse();
        }

        private void label3_MouseEnter(object sender, EventArgs e)
        {
            label3.BackColor = Color.Gainsboro;
        }

        private void label3_MouseLeave(object sender, EventArgs e)
        {
            label3.BackColor = Color.White;
        }
    }
}
