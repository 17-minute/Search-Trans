English version is under construction

# 概要 / Overview
Excel形式の翻訳ファイルをチェックするためのPowerShell関数です。
B列に原文、C列に対応する訳文が入っている翻訳ファイルの訳文が原文通りになっているか確かめるために使う関数になります。

This is a PowerShell function for checking Excel files that have translation data. 
The function can help us to check if translated texts are translated correctly. The specified translation files are supposed to  contain original texts in B column and translated texts in C column.

<img src="https://github.com/17-minute/Search-Trans/blob/img/excel.PNG" width="60%">

## 用例1 / Example 1
ファイル内に「value」という語が含まれているかどうかを確認したい場合は、検索用のウィンドウの「Search string (1)」に「value」と入力して「OK」をクリックします。

If you want to check if the specified files contain the word "value", enter "value" at "Search string (1)" in the window for search and click "OK".

<img src="https://github.com/17-minute/Search-Trans/blob/img/search_window_value.PNG" width="60%">

実行すると、関数実行時に指定したディレクトリ内のExcelファイルから「value」を含むセルを検索します。
検索後、次の情報を出力します。
- 検索対象となるディレクトリ名
- 検索文字列
- 検索文字列と一致した行数
- 一致した行に関する情報
  - その行が含まれるファイル名
  - 行位置
  - 翻訳データ（原文・訳文、検索文字列の部分は赤字）
  
After executing the command, it searches cells containing the string "value" from Excel files whose parent folder is specified as the augment when executing the function.
It outputs the following information after finishing the search.

- Name of the directory for search
- Search string
- Number of rows that contain the search string
- Information on the rows that match the search string
  - Name of the file that contains the row
  - Location of the row
  - Translation data (original/translation, the search string's color is red)

<img src="https://github.com/17-minute/Search-Trans/blob/img/result_value.PNG" width="60%">

## 用例2 / Example 2
任意で「Search string (2)」に文字列を指定すると、「Search string (1)」指定した検索文字列で得られたセルのテキストに対応する訳文（原文）に対してその文字列を検索します。

If you specify a string as "Search string (2)"(optional), it searchs the translated(original) sentences corresponding to the cell text that contains the string "Search string (1)" to see if the optional string is included in these sentences.  

<img src="https://github.com/17-minute/Search-Trans/blob/img/search_window_Enable.PNG" width="60%">

文字列がテキスト内にある場合はそのまま出力されますが、ない場合はまとめて最後に出力されます。それらの情報は黄色の文字になります。

If the string is included in the text, the text is output as it is. But if not, such sentences are displayed at the end of the output. The letter colour is yellow.

<img src="https://github.com/17-minute/Search-Trans/blob/img/result_Enable.PNG" width="60%">

## オプション / Option
ウィンドウのチェックボックスで次の2つを制御できます。
- 「Search string (1)」の言語設定（デフォルト（未チェック）は原文）
- 大文字と小文字を区別するかどうかの切り替え（デフォルト（未チェック）は大文字と小文字の区別なし）

You can control the following using the checkbox in the window.
- Language setting for "Search string (1)" (default(uncehcked) is original language)
- Switching case-sensitive or not (deafult(unchecked) is not case-sensitive)

<img src="https://github.com/17-minute/Search-Trans/blob/img/search_window_variable.PNG" width="60%">

<img src="https://github.com/17-minute/Search-Trans/blob/img/result_variable.PNG" width="60%">

# 実行環境 / Environment
- OS: Windows 10 64-bit operating system
- PowerShell: 
<img src="https://github.com/17-minute/Search-Trans/blob/img/version.png" width="60%">
※ PowerShell 6には対応していません。おそらく、PowerShell 7には対応すると思われます。 

※ This function does not support PowerShell 6, but maybe does PowerShell 7.

# 事前準備 - Excelファイルの整形 / Before execution - Excel files arrangement
検索対象となるExcelファイルは次のように整えます。
- B列に原文、C列に訳文を入れる
- 同じ行のB列とC列同士の内容は対応させる
- 翻訳データはシート1にのみに入れる（シート1以外は検索対象外となる）

Excelファイルの拡張子は.xls、.xlsx、.xlsmのいずれかであれば対応できます。

Excel files for search need to be arranged in the following manner.
- Original texts are in B column, and translated in C column
- Texts in column B and C are corresponded with each other in light of what they mean
- Only sheet 1 can contain translation data (sheets other than sheet 1 are ignored) 


# 用法 / Usage
1. 関数を実行 / Execute the function

   実行演算子などを使用
   Use dot sourcing operator
   
   <img src="https://github.com/17-minute/Search-Trans/blob/img/execution_func.PNG" width="60%">
2. 関数の呼び出し / Call the function

   検索対象となるExcelファイルが入ったディレクトリを-directoryパラメータの引数にして実行（パラメータ名は省略可能）
   
   Specify directory that contains Excel files for search as the augment of -directory parameter (you can omit the parameter name)
   
   <img src="https://github.com/17-minute/Search-Trans/blob/img/execution_command.PNG" width="60%">
   
3. 検索ウィンドウの入力 / Enter the form in the search window
   - Search string (1): ファイル内全体で検索したい文字列を入力
   - Search string (2): 「Search string (1)」を含む文に対応する訳文（原文）内で検索したい文字列を入力（任意）
   
   - Search string (1): Enter a string you want to search for in the whole files
   - Search string (2): Enter a string you want to search for in the sentences corresponding to the sentences that contain the specified string at "Search string (1)"
   
4. チェックボックスのチェック / Check in the checkbox
   - Search string (1) is the target language: 「Search string (1)」が訳文の言語である場合はチェック
   - Case-sensitive: 大文字と小文字を区別する場合はチェック
   
   - Search string (1) is the target language: Check if the string at "Search string (1)" is the target language 
   - Case-sensitive: Check if you want to search the string(s) as case-sensitive


# リファレンス / Reference
- [【PowerShell】Windowsフォームにテキストボックスを表示して入力できるようにする](https://hosopro.blogspot.com/2017/11/powershell-windows-form-textbox.html)
- [PowerShellで複数のExcelファイルを一括検索する](https://qiita.com/nejiko96/items/b423e2dda90181ef524e)
- [Microsoft.Office.Tools.Excel Namespace](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2013/7fzyhc74%28v%3dvs.120%29)
- Lee Homes, Windows PowerShell Cookbook 3rd Edition (O'reilly) 
- 吉澤生, PowerShell実践ガイドブック (マイナビ出版)
- 山田祥寛, 独習C# 新版 (翔泳社) 

# 作成者 / Author
17-minute

Qiita: @yasaram (https://qiita.com/yasaram)

