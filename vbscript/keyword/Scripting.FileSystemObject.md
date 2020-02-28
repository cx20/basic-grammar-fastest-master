# Scripting.FileSystemObject
## 概要
Scripting.FileSystemObject は、ファイル システムにアクセスするためのオブジェクトを表す ProgID です。

### サンプル
```vbscript
' File : ShowFileInfo.vbs
' Usage : CScript //Nologo ShowFileInfo.vbs
' Description : VBScript から FileSystemObject を使用してファイル情報を表示するサンプル
Option Explicit

Call Main()

Sub Main()
   Dim strFileName
   strFileName = "ShowFileInfo.vbs"
   Call ShowFileInfo( strFileName )
End Sub

Sub ShowFileInfo( strFileName )
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   Dim file
   Set file = fso.GetFile(strFileName)   
   
   DebugPrint "ファイル名   : [" & file.Name & "]"
   DebugPrint "サイズ       : [" & file.Size & "]"
   DebugPrint "作成日時 　  : [" & file.DateCreated & "]"
   DebugPrint "最終更新日時 : [" & file.DateLastModified & "]"
   DebugPrint "アクセス日時 : [" & file.DateLastAccessed & "]"
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
ファイル名   : [ShowFileInfo.vbs]
サイズ       : [935]
作成日時 　  : [2008/07/29 1:32:35]
最終更新日時 : [2008/07/29 1:33:58]
アクセス日時 : [2008/07/29 1:33:58]
```

### 参考情報
- FileSystemObject の概要 - MSDN Library
- スクリプト ラインタイム リファレンス - MSDN Library
- FileSystemObjectオブジェクトを利用する（1） － ＠IT - @IT
