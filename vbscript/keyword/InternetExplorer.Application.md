# InternetExplorer.Application
## 概要
InternetExplorer.Application は、Internet Explorer オブジェクトを表す ProgID です。

このオブジェクトを使用することで Internet Explorer の機能をオートメーション操作することが可能です。

### サンプル
```vbscript
' File : ShowBrowser.vbs
' Usage : CScript //Nologo ShowBrowser.vbs
' Description : VBScript から InternetExplorer.Application を使用して IE を操作するサンプル
Option Explicit

Call Main()

Sub Main()
    Dim strURL
    strURL = "http://www.microsoft.com/japan/"
    DebugPrint "[" & strURL & "] を Internet Explorer で表示"
    Call ShowBrowser( strURL )
End Sub

Sub ShowBrowser( strURL )
    Dim ie
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate strURL
End Function

Sub DebugPrint( strMessage )
    ' WSH で実行する場合
    WScript.Echo strMessage
    ' VBA で実行する場合
    ' Debug.Print strMessage
End Sub
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
