# Scripting.Dictionary
## 概要
Scripting.Dictionary は、キーと値を関連付ける辞書オブジェクトを表す ProgID です。

### サンプル
```vbscript
' File : ShowDicData.vbs
' Usage : CScript //Nologo ShowDicData.vbs
' Description : VBScript から Scripting.Dictionary を使用して辞書データを表示するサンプル
Option Explicit

Call Main()

Sub Main()
    Dim dic
    Set dic = CreateObject("Scripting.Dictionary")
    dic.Add "a", "123"
    dic.Add "b", "456"
    dic.Add "c", "789"
    
    Call ShowDicData( dic )
End Sub

Sub ShowDicData( dic )
    Dim key
    For Each key In dic.Keys
        WScript.Echo "key : [" & key & "], value = [" & dic(key) & "]"
    Next
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
key : [a], value = [123]
key : [b], value = [456]
key : [c], value = [789]
```

### 参考情報
- Dictionary オブジェクト - MSDN Library
- Dictionary オブジェクト - Microsoft TechNet
- Hey, Scripting Guy! 配列から重複した項目を削除する方法はありますか - Microsoft TechNet
