# System.Collections
## 概要
System.Collections は .NET Framework の名前空間（Namespace）です。

.NET Framework のクラスの一部が COM コンポーネントとして登録されている為、VBScript 等から .NET Framework の機能を利用することが可能です。

|ProgID|リファレンス|
|:-----|:-----|
|System.Collections.ArrayList|ArrayList クラス|
|System.Collections.Hashtable|Hashtable クラス|
|System.Collections.Queue|Queue クラス|
|System.Collections.SortedList|SortedList クラス|
|System.Collections.Stack|Stack クラス|

### サンプル
```vbscript
' File : ShowArrayList.vbs
' Usage : CScript //Nologo ShowArrayList.vbs
' Description : VBScript から .NET Framework の ArrayList を使用するサンプル
Option Explicit

Call Main()

Sub Main()
    Dim colItems
    Set colItems = CreateObject("System.Collections.ArrayList")

    colItems.Add "AAA"
    colItems.Add "BBB"
    colItems.Add "CCC"
    
    Call ShowCollections( colItems )
End Sub

Sub ShowCollections( colItems )
    Dim colItem
    For Each colItem In colItems
        DebugPrint colItem
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
AAA
BBB
CCC
```

### 参考情報
- Hey, Scripting Guy!: 発言には注意しましょう - Microsoft TechNet
- System.Collections 名前空間 - MSDN Library
