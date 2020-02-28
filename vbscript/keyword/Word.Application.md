# Word.Application
## 概要
Word.Application は Word アプリケーション オブジェクトを表す ProgID です。

Word アプリケーション オブジェクトは、通常、Word をオートメーション操作する場合に使用します。

### サンプル
```vbscript
' File : Make9x9TableByWord.vbs
' Usage : CScript //Nologo Make9x9TableByWord.vbs
' Description : VBScript から Word を使用して「九九表」を作成するサンプル
Option Explicit

Call Main()

Sub Main()
    Dim word
    Set word = CreateObject("Word.Application")
    word.Visible = True
    Dim doc
    Set doc = word.Documents.Add()
    Dim rng
    Set rng = doc.range()
    doc.Tables.Add rng, 9, 9
    Dim tbl
    Set tbl = doc.Tables(1)
    Call Make9x9Table( tbl )
End Sub

Sub Make9x9Table( tbl )
    Dim x
    Dim y
    For y = 1 To 9
        For x = 1 To 9
            tbl.Cell( y, x ).Range.Text = x * y
        Next
    Next
End Sub
```

### 実行結果

|1|1|2|3|4|5|6|7|8|9|
|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
|2|2|4|6|8|10|12|14|16|18|
|3|3|6|9|12|15|18|21|24|27|
|4|4|8|12|16|20|24|28|32|36|
|5|5|10|15|20|25|30|35|40|45|
|6|6|12|18|24|30|36|42|48|54|
|7|7|14|21|28|35|42|49|56|63|
|8|8|16|24|32|40|48|56|64|72|
|9|9|18|27|36|45|54|63|72|81|

### 参考情報
- Word オブジェクト モデル - MSDN Library
- Hey, Scripting Guy! フォルダ内のすべての Word 文書の合計ページ数を取得する方法はありますか - Microsoft TechNet
- Hey, Scripting Guy! Microsoft Word の既定のファイル保存形式を変更する方法はありますか - Microsoft TechNet
- 複数のバージョンの Office を Office 2003 と併用することに関する情報 - Microsoft KB
- 複数のバージョンの Office がインストールされている場合の Office オートメーションについて - Microsoft KB
- Office アプリケーションのパスを調べる方法 - Microsoft KB
