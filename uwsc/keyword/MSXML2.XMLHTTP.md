# MSXML2.XMLHTTP
## 概要
MSXML2.XMLHTTP は、XMLHttpRequest オブジェクト を表す ProgID です。

XMLHttpRequest を使用することで、Web サーバーに情報を送信したり受信することが可能です。

XMLHttpRequest オブジェクトは、IE7 からは IE の組み込みオブジェクトとして使用することが可能です。

このライブラリはクライアント用途に設計されたライブラリ WinINet（Microsoft Windows Internet）に依存しています。

サーバー用途（GUI を使用しないプログラム）では、WinHTTP（Windows HTTP Services）ベースの MSXML2.ServerXMLHTTP

もしくは WinHttp.WinHttpRequest を使うことが推奨されているようです。

### サンプル
```vbscript
// File : SaveUrlToFile.uws
// Description : UWSC から XMLHTTP を使用して Web の内容を保存するサンプル。
option Explicit

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
    http = CreateOleObj("MSXML2.XMLHTTP") 
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
ようこそ<UserID>さん … IE の Cookie が使用される為、ブラウザ経由と同じユーザーとしてアクセスします。
Myはてな
ログアウト
ヘルプ
```

### 参考情報
- XMLHttpRequest Object - MSDN Library (英語)
- About Native XMLHTTP - MSDN Library (英語)
- [HOWTO] XMLHTTP を使用してバイナリ ストリームを送信する方法 (2003/06/09 Rev:2.1) - Microsoft KB
- WebDAV を使用して Visual Basic からメッセージを送信する方法 (2005/10/25 Rev:3.0) - Microsoft KB
- Microsoft XML パーサー (MSXML) のバージョン一覧 - Microsoft KB
- INFO: MSXML 4.0 Specific GUIDs and ProgIDs (2001/11/21 Rev:1.0) - Microsoft KB (英語)
- MSXML 4.0 SP2 における標準仕様への準拠とセキュリティの変更 - Microsoft KB
- ServerXMLHTTP に関してよく寄せられる質問 (FAQ) - Microsoft KB
- Hey, Scripting Guy! Web ページがアクセス可能かどうかを確認する方法はありますか - Microsoft TechNet
- PRB:XMLHTTP ではサーバーからの HTTP 要求はサポートされない - Microsoft KB
- [INFO] サービスでは WinInet の使用はサポートされない - Microsoft KB
- Win32 インターネット拡張機能 (WinInet) - MSDN Library
- Porting WinINet Applications to WinHTTP (Windows) - MSDN Library (英語)
- Ask the Performance Team : Under the Hood: WinINet - TechNet Blogs (英語)
- XMLHttpRequest - Wikipedia
- HTTP Reference - MSDN Library (英語)
