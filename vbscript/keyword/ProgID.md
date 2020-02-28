# ProgID
## 概要
ProgID とは、COM クライアント*1から、COM サーバー*2を呼び出すときに使用する ID のことです。

内部的には、CLSID（クラスID）と呼ばれる 16進数の ID（GUID）で管理されています。この CLSID に名前を付けたものが ProgID になります。

ProgID を利用できる COM クライアントとしては VB、VBA の他、VBScript や JScript 等があります。

この情報は、実際にはレジストリに登録されています。
```
HKEY_CLASSES_ROOT\<ProgID>
```
ProgID の形式としては <プロジェクト名>.<クラス名> が多いですが、場合によっては、

<プロジェクト名>.<クラス名>.<バージョン番号> のようにバージョン番号が付与される場合もあります。

### サンプル
```vbscript
' File : SpeakHello.vbs
Dim voice                                ' コンポーネント参照用の変数を宣言します。
Set voice = CreateObject("SAPI.SpVoice") ' CreateObject(<ProgID>) でコンポーネントを生成します。
voice.Speak "Hello"                      ' コンポーネントのメソッドを呼び出します。
```

### 実行方法
1. メモ帳（notepad.exe）を起動し上記サンプルを貼り付けます。
2. 「SpeakHello.vbs」と名前を付けてデスクトップに保存します。
3. デスクトップ上にある「SpeakHello.vbs」をダブルクリックします。「Hello」と読み上げます。
※ 実行時エラーが表示される場合は、指定した ProgID が Windows に登録されていない可能性があります。

### Windows 標準 コンポーネント
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|ADO (ActiveX Data Objects)	        |ADODB.Connection	|ADO API リファレンス|
|DAO (Data Access Objects)	        |DAO.DBEngine	|DAO から ADO への移植|
|CDO（Collaboration Data Objects）	|CDO.Message	|CDO Library / リファレンス|
|FileSystem Object	                |Scripting.FileSystemObject	|スクリプト ラインタイム リファレンス|
|Script Runtime Dictionary	        |Scripting.Dictionary	|Dictionary オブジェクト|
|VBScript Regular Expression	      |VBScript.RegExp	|RegExp オブジェクト|
|Windows Script Host Shell Object	 |WScript.Shell	|Windows Script Host リファレンス|
|Shell Automation Service	         |Shell.Application	|Scriptable Shell Objects (英語)|
|Internet Explorer	                |InternetExplorer.Application	|InternetExplorer Object (英語)|
|SAPI (Microsoft Speech API)	      |SAPI.SpVoice	|Microsoft Speech API 5.3 (英語)|
|CAPICOM (CryptoAPI COM)	          |CAPICOM.Utilities	|CAPICOM Reference (英語)|
|WBEM Scripting Locator	            |WbemScripting.SWbemLocator	|WMI Reference (英語)|
|XML DOM	                          |MSXML2.DOMDocument	|XML DOM Objects (英語)|
|XMLHTTP (WinINet)	                |MSXML2.XMLHTTP	|XMLHttpRequest Object (英語)|
|ServerXMLHTTP (WinHTTP)	          |MSXML2.ServerXMLHTTP	|IServerXMLHTTPRequest/ServerXMLHTTP (英語)|
|WinHTTP	                          |WinHttp.WinHttpRequest	|WinHttpRequest Object (英語)|
|WUAPI (Windows Update API)	        |Microsoft.Update.Session	|Windows Update Agent (WUA) API Reference (英語)|

### Windows 7 標準 コンポーネント
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|Location API|LocationDisp.LatLongReportFactory|LocationDisp.LatLongReportFactory Object (英語)|

### .NET Framework クラス
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|Microsoft .NET Runtime	|System.Collections	|System.Collections 名前空間|
|Microsoft .NET Runtime	|System.IO	|System.IO 名前空間|
|Microsoft .NET Runtime	|System.Security	|System.Security 名前空間|
|Microsoft .NET Runtime	|System.Text	|System.Text 名前空間|

## 各製品付属コンポーネント
### Microsoft Office 付属コンポーネント
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|Microsoft Excel	|Excel.Application	|Microsoft Office Excel オブジェクト モデル|
|Microsoft Word	|Word.Application	|Word オブジェクト モデル|
|Microsoft Access	|Access.Application	|Microsoft Access オブジェクト モデル|
|Microsoft Outlook	|Outlook.Application	|Outlook オブジェクト モデル|
|Microsoft Power Point	|Powerpoint.Application	|PowerPoint オブジェクト モデル|

### Microsoft Office 2003/2007 付属コンポーネント（追加インストール）
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|Microsoft Office Document Imaging	|MODI.Document	|Microsoft Office Document Imaging オブジェクト モデル（英語）|

### Visual Studio 付属コンポーネント
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|TypeLib Information	|TLI.TLIApplication*3	|TypeLib Information オブジェクト（英語）|

### SQL Server 2000 付属コンポーネント
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|SQL-DMO	|SQLDMO.SQLServer	|SQL-DMO Reference|

### Google Earth 付属コンポーネント
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|Google Earth	|GoogleEarth.ApplicationGE	|Google Earth COM API Documentation (英語)|

### iTunes 付属コンポーネント

|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|iTunes	|iTunes.Application	|iTunes COM Windows SDK (要登録)|

### Evernote 付属コンポーネント
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|Evernote	|enapi.Evernote	|Evernote API(英語)|

### Sleipnir 付属コンポーネント
|製品名|サンプル|リファレンス|
|:----|:-------|:----------|
|Sleipnir	|Sleipnir.API	|Sleipnir.API リファレンス|

*1：オートメーション クライアントと呼ぶ場合もあります。

*2：オートメーション サーバーと呼ぶ場合もあります。
