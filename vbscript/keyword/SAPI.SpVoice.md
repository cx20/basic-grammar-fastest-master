# SAPI.SpVoice
## 概要
SAPI.SpVoice は、Speech API (SAPI) の 音声合成エンジン オブジェクトを表す ProgID です。

このオブジェクトを使用することで、任意のテキストを読み上げることができます。

（ただし、設定されている音声合成エンジンにより、日本語の読み上げが行えない場合があります。）


利用可能な 音声合成エンジンは [コントロール パネル] - [音声認識] にて確認することが可能です。

ただし、Windows XP より前のバージョンの OS には、SAPI が組み込まれておりません。

別途、下記サイトよりダウンロードするか SAPI が搭載されている製品（Office XP, Office 2003 等）を導入する必要があります。

SAPI 5.1 のコア コンポーネントの再配布方法

### バージョン
各 OS に搭載されている SAPI のバージョンとダウンロード可能なバージョン

|OS|5.1|5.3|5.4|10.1|
|:-|:--|:--|:--|:---|
|Microsoft Windows 98|DL|　|　|　|
|Microsoft Windows Me|DL|　|　|　|
|Microsoft Windows NT|DL|　|　|　|
|Microsoft Windows 2000|DL|　|　|　|
|Microsoft Windows XP|○|　|　|　|
|Microsoft Windows Vista|　|○|　|DL|
|Microsoft Windows 7|　|　|○|DL|

### 音声合成エンジン（TTS）
|音声|言語|製品|
|:---|:---|:---|
|Microsoft Sam|英語|Microsoft Windows XP|
|Microsoft Anna|英語|Microsoft Windows Vista, Windows 7|
|Microsoft Mike|英語|Microsoft Speech SDK 5.1|
|Microsoft Mary|英語|Microsoft Speech SDK 5.1|
|LH Kenji|日本語|Microsoft Office 2003|
|LH Naoko|日本語|Microsoft Office 2003|
|Microsoft Haruka|日本語|Microsoft Speech Platform - Server Runtime Languages (Version 10.1)|

### サンプル
```vbscript
' File : SpeakText.vbs
' Usage : CScript //Nologo SpeakText.vbs
' Description : VBScript から SAPI を使用してテキストを読み上げるサンプル。
Option Explicit

Call Main()

Sub Main()
    Dim strText
    strText = "Hello, VB Script World."
    DebugPrint "[" & strText & "] を SAPI で読み上げ"
    Call SpeakText( strText )
End Sub

Function SpeakText( strText )
    Dim voice
    Set voice = CreateObject("SAPI.SpVoice")
    ' SAPI 10.1 を使用する場合は、ProgID を「Speech.SpVoice」とする必要があります。
    'Set voice = CreateObject("Speech.SpVoice")
    voice.Speak strText
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
[Hello, VB Script World.] を SAPI で読み上げ
```

### 参考情報
- Microsoft Speech API 5.3 - MSDN Library (英語)
- SAPI 5.1 のコア コンポーネントの再配布方法 (2008/01/23 Rev:4.2) - Microsoft KB
- Download details: Speech Software Development Kit 5.1 (2001/08/08 Ver:5.1) - Microsoft Download Center(英語)
- Windows Vista において [ナレータを有効にします] を起動しても 日本語のナレータを利用できない - Microsoft KB
- Don't Worry, Get SAPI: Using the Speech API to Add Voice and Sound Effects to a Script - Microsoft TechNet (英語)
- Speech Application Programming Interface - Wikipedia
- CodeProject: Exploring the SpVoice Class of MS SAPI 5.1 to use different available features for TTS. Free source code and programming help (2006/12/11) - Code Project (英語)
