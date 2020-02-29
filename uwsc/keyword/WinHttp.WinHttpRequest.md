# WinHttp.WinHttpRequest
## 概要
WinHttp.WinHttpRequest は、WinHTTP オブジェクト を表す ProgID です。

WinHttp.WinHttpRequest を使用することで、Web サーバーに情報を送信したり受信することが可能です。

ライブラリの実装には WinHTTP（HTTP通信専用コンポーネント）が用いられています。

WinHTTP は Windows Update の更新用サービス（Background Intelligent Transfer Service）等でも使用されています。

同様のライブラリに WinINet ベースの MSXML2.XMLHTTP と WinHTTP ベースの MSXML2.ServerXMLHTTP があります。

### サンプル
```vbscript
// File : SaveUrlToFile.uws
// Description : UWSC から WinHTTP を使用して Web の内容を保存するサンプル。
Option Explicit

Main()

Procedure Main()
    Dim strURL
    strURL = "http://www.hatena.ne.jp/"
    Dim strFileName
    strFileName = "hatena.txt"
    SaveUrlToFile( strURL, strFileName )
Fend

Procedure SaveUrlToFile( strURL, strFileName )
    Dim http
    http = CreateOleObj("WinHttp.WinHttpRequest.5.1") 
    http.Open( "GET", strURL, False )
    http.Send

    Dim strMessage
    Dim bOverwrite
    Dim bUnicode
    strMessage = http.ResponseText
    bOverwrite = True
    bUnicode = True // UNICODE のページを SJIS のファイルに保存するとエラーが発生することがある為、UNICODE 形式にしています。
    WriteToToFile( strMessage, strFileName, bOverwrite, bUnicode )
Fend

Procedure WriteToToFile( strMessage, strFileName, bOverwrite, bUnicode )
    Dim fso
    fso = CreateOleObj("Scripting.FileSystemObject")
    Dim file
    file = fso.CreateTextFile(strFileName, bOverwrite, bUnicode)
    file.Write( strMessage )
    file.Close
Fend
```

### 実行結果
```
ようこそゲストさん … IE の Cookie が使用されない為、ゲストユーザーとしてアクセスします。
ユーザー登録
ログイン
ヘルプ
```

### 利用可能な WinHTTP のバージョン
|OS|5.0|5.1|
|:-|:--|:--|
|Microsoft Windows NT 4.0|○*1|　|
|Microsoft Windows 2000|○*2|○*3|
|Microsoft Windows XP|　|○*4|
|Microsoft Windows Server 2003|　|○*5|
|Microsoft Windows Vista|　|○|
|Microsoft Windows 7|　|○|
|Microsoft Windows 8(Customer Preview)|　|○|

### 参考情報
- WinHttpRequest Object Reference (Windows) - MSDN Library (英語)
- About WinHTTP (Windows) - MSDN Library (英語)
- Using the WinHttpRequest COM Object (Windows) - MSDN Library (英語)
- ServerXMLHTTP に関してよく寄せられる質問 (FAQ) - Microsoft KB
- ServerXMLHTTP が動作するには Proxycfg ツールを実行する必要がある - Microsoft KB
- [INFO] サービスでは WinInet の使用はサポートされない - Microsoft KB
- Porting WinINet Applications to WinHTTP (Windows) - MSDN Library (英語)
- Ask the Performance Team : Under the Hood: WinHTTP - TechNet Blogs (英語)
- Windows と C++: 非同期 WinHTTP - MSDN Library
- HTTP Reference - MSDN Library (英語)

*1：Windows NT 4.0 + IE5 以上。ただし WinHTTP 5.0 は 2004年10月にサポート切れの為、現在はダウンロードできません。

*2：Windows 2000 以上。ただし WinHTTP 5.0 は 2004年10月にサポート切れの為、現在はダウンロードできません。

*3：Windows 2000 SP3 以上

*4：Windows XP SP1 以上

*5：Windows Server 2003 SP1 以上
