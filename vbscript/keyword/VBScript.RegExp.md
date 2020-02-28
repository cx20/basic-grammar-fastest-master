# VBScript.RegExp
## 概要
VBScript.RegExp は、正規表現を取り扱うオブジェクトを表す ProgID です。

### サンプル
```vbscript
' File : TestRegExp.vbs
' Usage : CScript //Nologo TestRegExp.vbs
' Description : VBScript から VBScript.RegExp を使用して正規表現をテストするサンプル
Option Explicit

Call Main()

Sub Main()
    Dim strSrc
    Dim strDst
    strSrc = "abcあいうえおDEF"
    strDst = RemoveJapaneseString( strSrc )
    DebugPrint "strSrc = [" & strSrc & "]"
    DebugPrint "strDst = [" & strDst & "]"
End Sub

Function RemoveJapaneseString( strData )
    Dim strResult
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True   
    re.Pattern = "[^A-Za-z]"
    strResult = re.Replace(strData, "")
    RemoveJapaneseString = strResult
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
strSrc = [abcあいうえおDEF]
strDst = [abcDEF]
```

### 参考情報
- Regular Expression (RegExp) オブジェクト - MSDN Library
- Clinick's Clinic on Scripting: 正規表現による Visual Basic Scripting Edition (VBScript) の機能強化 - MSDN Library
- Hey, Scripting Guy! テキスト ファイル内の電話番号を検索する方法はありますか - TechNet
- Hey, Scripting Guy! あるフォルダ内のすべてのテキスト ファイルについて電話番号を検索する方法はありますか - TechNet
- Hey, Scripting Guy! 文字列に含まれているアルファベット以外のすべての文字を削除する方法はありますか - TechNet
- Hey, Scripting Guy! 文字列のすべての文字が A ～ Z (大文字) または数字 0 ～ 9 のどちらかであるかどうかを確認する方法はありますか - TechNet
- Microsoft Visual Basic 6.0 で正規表現を使用する方法 - Microsoft KB
