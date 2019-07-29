English version is under construction

# 概要:
Excel形式の翻訳ファイルをチェックするためのPowerShell関数です。

## 用例1
ファイル内に「value」という語が含まれているかどうかを確認したい場合は、検索用のウィンドウの「Search string (1)」に「value」と入力して「OK」をクリックします。

<img src="https://github.com/17-minute/self-script/blob/img/new_search_window_value.PNG" width="60%">

実行すると、関数実行時に指定したディレクトリ内のExcelファイルから「value」を含むセルを検索します。
検索後、次の情報を出力します。
- 検索対象となるディレクトリ名
- 検索文字列
- 検索文字列と一致した行数
- 一致した行に関する情報
  - その行が含まれるファイル名
  - 行位置
  - 翻訳データ（原文・訳文、検索文字列の部分は赤字）
  
<img src="https://github.com/17-minute/self-script/blob/img/new_result_value.PNG" width="60%">

## 用例2
任意で「Search string (2)」に文字列を指定すると、「Search string (1)」指定した検索文字列で得られたセルのテキストに対応する訳文（原文）に対してその文字列を検索します。

<img src="https://github.com/17-minute/self-script/blob/img/search_window_enable.PNG" width="60%">

文字列がテキスト内にある場合はそのまま出力されますが、ない場合はまとめて最後に出力されます。それらの情報は黄色の文字になります。

<img src="https://github.com/17-minute/self-script/blob/img/new_result_Enable.PNG" width="60%">

## オプション
ウィンドウのチェックボックスで次の2つを制御できます。
- 「Search string (1)」を原文から検索するか訳文から検索するか
- ラテン文字を検索する際、大文字と小文字を区別するかどうか

<img src="https://github.com/17-minute/self-script/blob/img/new_search_window_variable_JP_case.PNG" width="60%">

<img src="https://github.com/17-minute/self-script/blob/img/new_result_variable_JP_case.PNG" width="60%">

# 実行環境:
- OS: Windows 10 64-bit operating system
- PowerShell: 
<img src="https://github.com/17-minute/self-script/blob/img/version.png" width="60%">
※ PowerShell 6には対応していません。おそらく、PowerShell 7には対応すると思われます。 

# Excelファイルの整形
検索対象となるExcelファイルは次のように整えます。
- B列に原文、C列に訳文を入れる
- 同じ行のB列とC列同士の内容は対応させる
- 翻訳データはシート1にのみに入れる（シート1以外は検索対象外となる）

<img src="https://github.com/17-minute/self-script/blob/master/excel.PNG" width="60%">

# 使いかた
1. 関数を実行

   実行演算子などを使用
   <img src="https://github.com/17-minute/self-script/blob/master/excel.PNG" width="60%">
2. 関数の呼び出し

   検索対象となるExcelファイルが入ったディレクトリを-directoryパラメータの引数にして実行（パラメータ名は省略可能）
   <img src="https://github.com/17-minute/self-script/blob/master/excel.PNG" width="60%">
3. 検索ウィンドウの入力
   - Search string (1): ファイル内で検索したい文字列を入力
   - Search string (2): 「Search string (1)」を含む文に対応する訳文（原文）内で検索したい文字列を入力（任意）
4. チェックボックスのチェック
   - Search string (1) is the target language: 「Search string (1)」が訳文の言語である場合はチェック
   - Case-sensitive: ラテン文字（いわゆるアルファベット）の大文字と小文字を区別する場合はチェック

# リファレンス
- [【PowerShell】Windowsフォームにテキストボックスを表示して入力できるようにする](https://hosopro.blogspot.com/2017/11/powershell-windows-form-textbox.html)
- [PowerShellで複数のExcelファイルを一括検索する](https://qiita.com/nejiko96/items/b423e2dda90181ef524e)
- [Microsoft.Office.Tools.Excel Namespace](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2013/7fzyhc74%28v%3dvs.120%29)
- Lee Homes, Windows PowerShell Cookbook 3rd Edition (O'reilly) 
- 吉澤生, PowerShell実践ガイドブック (マイナビ出版)
- 山田祥寛, 独習C# 新版 (翔泳社) 

# 作成者
17-minute
Qiita: @yasaram (https://qiita.com/yasaram)

