# PowerPoint.Application
## 概要
PowerPoint.Application は PowerPoint アプリケーション オブジェクトを表す ProgID です。

PowerPoint アプリケーション オブジェクトは、通常、PowerPoint をオートメーション操作する場合に使用します。

### サンプル
```vbscript
' File : Make9x9TableByPowerPoint.vbs
' Usage : CScript //Nologo Make9x9TableByPowerPoint.vbs
' Description : VBScript から PowerPoint を使用して「九九表」を作成するサンプル
Option Explicit

Call Main()

Sub Main()
    Dim ppt
    Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    
    Dim pre
    Set pre = ppt.Presentations.Add
    Dim slide
    Set slide = pre.Slides.Add(1, 12)   ' Const ppLayoutBlank = 12
    Call slide.Shapes.AddTable(9, 9)

    Dim tbl
    Set tbl = slide.Shapes(1).Table
    Call Make9x9Table(tbl)
End Sub

Sub Make9x9Table(tbl)
    Dim x
    Dim y
    For y = 1 To 9
        For x = 1 To 9
            tbl.Cell(y, x).Shape.TextFrame.TextRange.Text = x * y
        Next
    Next
End Sub
```

### 実行結果
|1  |2  |3  |4  |5  |6  |7  |8  |9  |
|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
|2|4|6|8|10|12|14|16|18|
|3|6|9|12|15|18|21|24|27|
|4|8|12|16|20|24|28|32|36|
|5|10|15|20|25|30|35|40|45|
|6|12|18|24|30|36|42|48|54|
|7|14|21|28|35|42|49|56|63|
|8|16|24|32|40|48|56|64|72|
|9|18|27|36|45|54|63|72|81|

### 参考情報
- PowerPoint オブジェクト モデル - MSDN Library
- Hey, Scripting Guy! スクリプトから PowerPoint スライド ショーを実行する方法はありますか - Microsoft TechNet
- 複数のバージョンの Office を Office 2003 と併用することに関する情報 - Microsoft KB
- 複数のバージョンの Office がインストールされている場合の Office オートメーションについて - Microsoft KB
- Office アプリケーションのパスを調べる方法 - Microsoft KB
