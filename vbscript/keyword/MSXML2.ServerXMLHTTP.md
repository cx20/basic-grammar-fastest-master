# MSXML2.ServerXMLHTTP
## 概要
MSXML2.ServerXMLHTTP は Server XML HTTP オブジェクト を表す ProgID です。

Server XML HTTP を使用することで、Web サーバーに情報を送信したり受信することが可能です。

ライブラリの実装には WinHTTP（HTTP通信専用コンポーネント）が用いられています。

同様のライブラリに WinINet ベースの MSXML2.XMLHTTP や WinHTTP ベースの WinHttp.WinHttpRequest があります。

現在は WinHttp.WinHttpRequest を使うことが推奨されているようです。

### サンプル
```vbscript
' File : SaveUrlToFile.vbs
' Usage : CScript //Nologo SaveUrlToFile.vbs
' Description : VBScript から Server XML HTTP を使用して Web の内容を保存するサンプル。
option Explicit

Call Main()

Sub Main()
    Dim strURL
    strURL = "http://www.hatena.ne.jp/"
    Dim strFileName
    strFileName = "hatena.txt"
    Call SaveUrlToFile( strURL, strFileName )
End Sub

Sub SaveUrlToFile( strURL, strFileName )
    Dim http
    Set http = CreateObject("MSXML2.ServerXMLHTTP") 
    http.Open "GET", strURL, False
    http.Send

    Dim strMessage
    Dim bOverwrite
    Dim bUnicode
    strMessage = http.ResponseText
    bOverwrite = True
    bUnicode = True ' UNICODE のページを SJIS のファイルに保存するとエラーが発生することがある為、UNICODE 形式にしています。
    Call WriteToToFile( strMessage, strFileName, bOverwrite, bUnicode )
End Sub

Sub WriteToToFile( strMessage, strFileName, bOverwrite, bUnicode )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
    Set file = fso.CreateTextFile(strFileName, bOverwrite, bUnicode)
    file.Write strMessage
    file.Close
End Sub
```

### 実行結果
```
ようこそゲストさん … IE の Cookie が使用されない為、ゲストユーザーとしてアクセスします。
ユーザー登録
ログイン
ヘルプ
```

### 参考情報
- IServerXMLHTTPRequest/ServerXMLHTTP - MSDN Library (英語)
- About WinHTTP (Windows) - MSDN Library (英語)
- [MSXML] ServerXMLHTTP はリダイレクトされた要求を受信できない - Microsoft KB
- ServerXMLHTTP に関してよく寄せられる質問 (FAQ) - Microsoft KB
- ServerXMLHTTP が動作するには Proxycfg ツールを実行する必要がある - Microsoft KB
- [INFO] サービスでは WinInet の使用はサポートされない - Microsoft KB
- Porting WinINet Applications to WinHTTP (Windows) - MSDN Library (英語)
- Ask the Performance Team : Under the Hood: WinHTTP - TechNet Blogs (英語)
- Windows と C++: 非同期 WinHTTP - MSDN Library
- HTTP Reference - MSDN Library (英語)
