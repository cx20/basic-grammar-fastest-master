# Scripting.FileSystemObject
## 概要
Scripting.FileSystemObject は、ファイル システムにアクセスするためのオブジェクトを表す ProgID です。

### サンプル
```vbscript
// File : ShowFileInfo.uws
// Description : UWSC から FileSystemObject を使用してファイル情報を表示するサンプル
Option Explicit

Main()

Procedure Main()
   Dim strFileName
   strFileName = "ShowFileInfo.uws"
   ShowFileInfo( strFileName )
Fend

Procedure ShowFileInfo( strFileName )
   Dim fso
   fso = CreateOleObj("Scripting.FileSystemObject")
   
   Dim file
   file = fso.GetFile(strFileName)   
   
   Print "ファイル名   : [" + file.Name + "]"
   Print "サイズ       : [" + file.Size + "]"
   Print "作成日時 　  : [" + file.DateCreated + "]"
   Print "最終更新日時 : [" + file.DateLastModified + "]"
   Print "アクセス日時 : [" + file.DateLastAccessed + "]"
Fend
```

### 実行結果
```
ファイル名   : [ShowFileInfo.uws]
サイズ       : [696]
作成日時 　  : [2010/02/09 0:12:48]
最終更新日時 : [2010/02/09 0:15:03]
アクセス日時 : [2010/02/09 0:12:48]
```

### 参考情報
- FileSystemObject の概要 - MSDN Library
- スクリプト ラインタイム リファレンス - MSDN Library
- FileSystemObjectオブジェクトを利用する（1） － ＠IT - @IT
