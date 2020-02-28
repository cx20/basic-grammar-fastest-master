# Shell.Application
## 概要
Shell.Application は、Windows シェル のオートメーション オブジェクトを表す ProgID です。

このオブジェクトを使用することでエクスプローラの機能をオートメーション操作することが可能です。

また、FTP やファイルの更新日時の変更といったファイル操作も可能です。

### サンプル
```vbscript
' File : ShowExplore.vbs
' Usage : CScript //Nologo ShowExplore.vbs
' Description : VBScript から Shell.Application を使用してエクスプローラを起動するサンプル
Option Explicit

Call Main()

Sub Main()
    Dim strFilePath
    strFilePath = "C:\"
    DebugPrint "[" & strFilePath & "] をエクスプローラで表示"
    Call ShowExplore( strFilePath )
End Sub

Sub ShowExplore( strFilePath )
    Dim shell
    Set shell = CreateObject("Shell.Application")
    shell.Explore strFilePath
End Sub

Sub DebugPrint( strMessage )
    ' WSH で実行する場合
    WScript.Echo strMessage
    ' VBA で実行する場合
    ' Debug.Print strMessage
End Sub
```

### 実行結果
```
[C:\] をエクスプローラで表示
```

### 参考情報
- Scriptable Shell Objects - MSDN Library (英語)
- Hey, Scripting Guy! ファイルのコピー中に視覚的なインジケータを表示する方法はありますか - Microsoft TechNet
- Hey, Scripting Guy! Internet Explorer のお気に入りに関連付けられている URL を返す方法はありますか - Microsoft TechNet
- Hey, Scripting Guy! デスクトップ上にウィンドウを並べて表示する方法はありますか - Microsoft TechNet
- Hey, Scripting Guy!: お決まりのごまかし - Microsoft TechNet
