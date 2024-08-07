using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO; // for debugging purposes

namespace aozoraOSC
{
    public partial class Form1 : Form
    {
        // aozora OSC version
        private readonly string aozoraVersion = "v1.0";

        // 初期設定
        private readonly int firstPage = 1;
        private int lastPage = 2;
        private string[] pages = { "　", "<rotate=90><rotate=0>（<rotate=90>本文<rotate=0>：<rotate=90>情報無し<rotate=0>）<rotate=90>" };

        // 無限ループ回避用（C#では bool 型のデフォルト値は false）
        private bool firstPageReset = false;
        private bool lastPageReset = false;

        // WebBrowserコントロールをクラスレベルで宣言
        private WebBrowser browser;

        // 連打防止用
        private bool isProcessing = false;


        //
        // 起動後最初に実行される
        //
        public Form1()
        {
            InitializeComponent();

            // 既定値 URL を選択状態にしておく（URL 貼り付け操作を楽にする）
            textBox1.SelectAll();

            // 送信ページの範囲指定（lastPage は自動で設定される）
            comboBox1.SelectedItem = firstPage.ToString();
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


        //
        // 青空文庫からデータ取得
        //
        private async void button1_Click(object sender, EventArgs e)
        {
            // 処理中でないことを確認
            if (isProcessing)
                return;

            string url = textBox1.Text;
            // URL の拡張子をチェックして、HTML ファイルのみを処理
            if (url.EndsWith(".html", StringComparison.OrdinalIgnoreCase) ||
                url.EndsWith(".htm", StringComparison.OrdinalIgnoreCase))
            {
                // 処理中
                isProcessing = true;

                // ページ範囲の無限ループ回避用
                firstPageReset = true;
                lastPageReset = true;

                // ボタンを無効化し、テキストを変更
                button1.Enabled = false;
                button1.Text = "⌛取得中";
                button2.Enabled = false;
                button2.Text = "⌛取得中";

                // 連打防止の強化
                await Task.Delay(300);

                try
                {
                    // ページ作成処理を開始し、完了を待つ
                    await RetrieveHtml();
                }
                finally
                {
                    // ボタンを再度有効化し、テキストを戻す
                    button1.Enabled = true;
                    button1.Text = "取得/Get";
                    button2.Enabled = true;
                    button2.Text = "VirtualCastへ送信";

                    // 処理終了
                    isProcessing = false;
                }
            }
            else
            {
                MessageBox.Show("青空文庫 XHTML ファイルの URL '***.html' を入力してください。\n\n例)  https://www.aozora.gr.jp/cards/000081/files/43754_17659.html");
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

            // 処理中
            isProcessing = true;

            // ボタンとコンボボックス（ページ範囲指定）を無効化し、テキストを変更
            // button2 のテキストは処理中に動的変更
            button1.Enabled = false;
            button1.Text = "🍣送信中";
            button2.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;

            try
            {
                // 時間のかかる処理を行う
                await SelectDataForSending();
            }
            finally
            {
                // ボタンとコンボボックスを再度有効化し、テキストを戻す
                button1.Enabled = true;
                button1.Text = "取得/Get";
                button2.Enabled = true;
                button2.Text = "VirtualCastに送信";
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;

                // 処理終了
                isProcessing = false;
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
        //
        private async Task RetrieveHtml()
        {
            // URL 取得
            string url = textBox1.Text;

            // WebBrowserコントロールを作成（既に作成されていれば再利用）
            if (browser == null)
            {
                browser = new WebBrowser();
                browser.ScriptErrorsSuppressed = true;
                browser.DocumentCompleted += EditData;
            }

            // ページを読み込む
            browser.Navigate(url);

            // DocumentCompleted イベントを待つ
            await Task.Run(() =>
            {
                var tcs = new TaskCompletionSource<bool>();
                WebBrowserDocumentCompletedEventHandler handler = null;
                handler = (s, args) =>
                {
                    tcs.SetResult(true);
                    browser.DocumentCompleted -= handler;
                };
                browser.DocumentCompleted += handler;
                tcs.Task.Wait();
            });
        }

        //
        // 読み込んだ情報から必要な情報を抽出し
        // 文字列処理や本文分割等を実施
        //
        private void EditData(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            // WebBrowser コントロールの Document プロパティを取得
            WebBrowser browser = (WebBrowser)sender;
            HtmlDocument doc = browser.Document;

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

            // class 名が "title" のすべての要素を取得
            foreach (HtmlElement element in h1Elements)
            {
                // class 属性が "title" である要素を探す
                if (element.GetAttribute("className") == "title")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        titleText = innerText.TrimStart('\n', '\r'); // 念のため just in case
                    }
                    break; // ループを終了
                }
            }

            // class 名が "subtitle" のすべての要素を取得
            foreach (HtmlElement element in h2Elements)
            {
                // class 属性が "subtitle" である要素を探す
                if (element.GetAttribute("className") == "subtitle")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除後 "\n" を１つ追加
                        subtitleText = innerText.TrimStart('\n', '\r'); // 念のため just in case
                        subtitleText = "\n" + subtitleText;
                    }
                    break; // ループを終了
                }
            }

            // class 名が "original_title" のすべての要素を取得
            foreach (HtmlElement element in h2Elements)
            {
                // class 属性が "original_title" である要素を探す
                if (element.GetAttribute("className") == "original_title")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除後 "\n" を１つ追加
                        original_titleText = innerText.TrimStart('\n', '\r'); // 念のため just in case
                        original_titleText = "\n" + original_titleText;
                    }
                    break; // ループを終了
                }
            }
            // 作品名をまとめて表示
            string totalTitleText = titleText + subtitleText + original_titleText;
            string resultTitle = totalTitleText;
            // 表示幅が 40 を超える時は「…」を追加し以降の文字列削除
            if (GetWidth(totalTitleText) > 40)
            {
                resultTitle = TruncateWithEllipsis(totalTitleText, 41);
            }
            textBox2.Text = resultTitle;


            // class 名が "author" のすべての要素を取得
            foreach (HtmlElement element in h2Elements)
            {
                // class 属性が "author" である要素を探す
                if (element.GetAttribute("className") == "author")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        authorText = innerText.TrimStart('\n', '\r'); // 念のため just in case
                    }
                    break; // ループを終了
                }
            }
            // 著者名を表示
            string resultAuthorText = authorText;
            // 表示幅が 40 を超える時は「…」を追加し以降の文字列削除
            if (GetWidth(authorText) > 40)
            {
                resultAuthorText = TruncateWithEllipsis(authorText, 41);
            }
            textBox3.Text = resultAuthorText;


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
                        translatorText = innerText.TrimStart('\n', '\r'); // 念のため just in case
                    }
                    break; // ループを終了
                }
            }
            // 翻訳者名を表示
            string resultTranslatorText = translatorText;
            // 表示幅が 40 を超える時は「…」を追加し以降の文字列削除
            if (GetWidth(translatorText) > 40)
            {
                resultTranslatorText = TruncateWithEllipsis(translatorText, 41);
            }
            textBox4.Text = resultTranslatorText;


            // class 名が "main_text" のすべての要素を取得
            foreach (HtmlElement element in divElements)
            {
                // class属性が "main_text" である要素を探す
                if (element.GetAttribute("className") == "main_text")
                {
                    // テキストを取得
                    string innerText = element.InnerText;

                    if (innerText != null)
                    {
                        // 先頭の改行を削除
                        main_textText = innerText.TrimStart('\n', '\r');

                        // 本文終了：底本との仕切り
                        main_textText += "\r\n\r\n\r\n――――――――――――――――――――\r\n\r\n";
                    }
                    break; // ループを終了
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
                        biblioText = innerText.TrimStart('\n', '\r');
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
                        after_textText = innerText.TrimStart('\n', '\r');
                    }
                    break; // ループを終了
                }
            }
            // 本文をまとめて表示
            string mainText = main_textText + biblioText + after_textText;
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
        static string AddLineBreaks(string text, int maxLineLength)
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
                || c == '('  || c == '['  || c == '{'; // 半角
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
                || c == ')'  || c == ']'  || c == '}'  // 半角
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
                || c == '?'  || c == '!'  //半角
                //中点類
                || c == '：' || c == '；' || c == '／' || c == '・'
                || c == ':'  || c == ';'  || c == '/'
                // 句点類
                || c == '，' || c == '．' // 全角
                || c == ','  || c == '.'  // 半角
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
        static string AddTagsToSpecialCharacters(string input, bool tagsDelete)
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
        public static string[] SplitIntoLines(string text)
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

            // 選択されたページ範囲
            for (int i = sendSelectedFirstPage - 1; i < sendSelectedLastPage; i++)
            {
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
            await Task.Delay(300);
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
        }
    }
}
