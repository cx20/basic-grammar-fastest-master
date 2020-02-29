# InternetExplorer.Application
## 概要
InternetExplorer.Application は、Internet Explorer オブジェクトを表す ProgID です。

このオブジェクトを使用することで Internet Explorer の機能をオートメーション操作することが可能です。

### サンプル
```vbscript
// File : ShowBrowser.uws
// Description : UWSC から InternetExplorer.Application を使用して IE を操作するサンプル
Option Explicit

Main()

Procedure Main()
    Dim strURL
    strURL = "http://www.microsoft.com/japan/"
    Print "[" + strURL + "] を Internet Explorer で表示"
    ShowBrowser( strURL )
Fend

Procedure ShowBrowser( strURL )
    Dim ie
    ie = CreateOleObj("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate( strURL )
Fend
```

### 実行結果
```
[http://www.microsoft.com/japan/] を Internet Explorer で表示
```

### 参考情報
- InternetExplorer Object - MSDN Library (英語)
- Internet Explorer のインスタンスの作成 - Microsoft TechNet
- Hey, Scripting Guy! 現在の画面の内容を問わず、画面への出力を上書きする方法はありますか - Microsoft TechNet
- Hey, Scripting Guy! スクリプトの実行中に進行状況バー (またはそれに似たもの) を表示する方法はありますか - Microsoft TechNet
- Hey, Scripting Guy! Web ページを定期的に更新する方法はありますか - Microsoft TechNet
