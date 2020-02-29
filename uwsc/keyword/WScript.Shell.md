# WScript.Shell
## 概要
WScript.Shell は、Windows シェル オブジェクトを表す ProgID です。

Shell.Application と似たような機能を持っていますが、WScript.Shell はシステム環境的なものにアクセスする場合に使用されることが多いです。

### サンプル
```vbscript
// File : ShowSpecialFolder.uws
// Description : UWSC から WScript.Shell を使用して特殊フォルダを表示するサンプル
Option Explicit

Main()

Procedure Main()
    Dim strFolderType
    strFolderType = "Desktop" 
    ShowSpecialFolder( strFolderType )
Fend

Procedure ShowSpecialFolder( strFolderType )
    Dim strFolderName
    strFolderName = GetSpecialFolder( strFolderType )
    Print "Folder Type : [" + strFolderType + "]"
    Print "Folder Name : [" + strFolderName + "]"
Fend

Function GetSpecialFolder( strFolderType )
    Dim strResult
    Dim shell
    shell = CreateOleObj("WScript.Shell")
    strResult = shell.SpecialFolders( strFolderType )
    Result = strResult
Fend
```

### 実行結果
```
Folder Type : [Desktop]
Folder Name : [C:\Documents and Settings\Administrator\デスクトップ]
```

### 参考情報
- Windows Script Host リファレンス - MSDN Library
- WshShell オブジェクト - MSDN Library
- Hey, Scripting Guy! UNC で指定した場所をウィンドウで開く方法はありますか - Microsoft TechNet
- Hey, Scripting Guy! コマンド ライン ツールの実行後にコマンド ウィンドウを開いたままにする方法はありますか - Microsoft TechNet
- Hey, Scripting Guy! Ping コマンドの出力を変更する方法はありますか - Microsoft TechNet
