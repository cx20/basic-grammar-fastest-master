# VBScript.RegExp
## 概要
VBScript.RegExp は、正規表現を取り扱うオブジェクトを表す ProgID です。

### サンプル
```vbscript
// File : TestRegExp.vbs
// Description : UWSC から VBScript.RegExp を使用して正規表現をテストするサンプル
Option Explicit

Main()

Procedure Main()
    Dim strSrc
    Dim strDst
    strSrc = "abcあいうえおDEF"
    strDst = RemoveJapaneseString( strSrc )
    Print "strSrc = [" + strSrc + "]"
    Print "strDst = [" + strDst + "]"
Fend

Function RemoveJapaneseString( strData )
    Dim strResult
    Dim re
    re = CreateOleObj("VBScript.RegExp")
    re.Global = True   
    re.Pattern = "[^A-Za-z]"
    strResult = re.Replace(strData, "")
    Result = strResult
Fend
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
