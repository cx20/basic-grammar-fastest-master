# Scripting.Dictionary
## 概要
Scripting.Dictionary は、キーと値を関連付ける辞書オブジェクトを表す ProgID です。

### サンプル
```vbscript
// File : ShowDicData.uws
// Description : UWSC から Scripting.Dictionary を使用して辞書データを表示するサンプル
Option Explicit

Main()

Procedure Main()
    Dim dic
    dic = CreateOleObj("Scripting.Dictionary")
    dic.Add( "a", "123" )
    dic.Add( "b", "456" )
    dic.Add( "c", "789" )
    
    ShowDicData( dic )
Fend

Procedure ShowDicData( dic )
    Dim key
    Dim i
    For i = 0 To GetOleItem( dic )-1
        key = ALL_OLE_ITEM[i]
        Print "key : [" + key + "], value = [" + dic.Item(key) + "]"
    Next
Fend
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
