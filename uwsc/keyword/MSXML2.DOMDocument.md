# MSXML2.DOMDocument
## 概要
MSXML2.DOMDocument は、XML DOM オブジェクト を表す ProgID です。

XML DOM を使用することで XML データをツリー構造のオブジェクトとして使用することが可能です。

### サンプル
```vbscript
// File : ShowXmlNodes.uws
// Description : UWSC から XML DOM を使用して 人力検索の質問一覧（RSS）のタイトル情報を表示するサンプル。
Option Explicit

Main()

Procedure Main()
    Dim dom
    dom = CreateOleObj("MSXML2.DOMDocument") 
    
    Dim strURL
    strURL = "http://q.hatena.ne.jp/list?mode=rss"
    Dim strXML
    strXML = GetXmlData( strURL )
    dom.loadXML( strXML )
    ShowXmlNodes( dom.childNodes )
Fend

Procedure ShowXmlNodes(Nodes)
    Dim node
    Dim i
    For i = 0 To GetOleItem(Nodes)-1
        node = ALL_OLE_ITEM[i]
        Ifb node.parentNode.nodeName = "title" Then
            Print node.nodeValue
        EndIf
        Ifb node.hasChildNodes Then
            ShowXmlNodes( node.childNodes )
        EndIf
    Next
Fend

Function GetXmlData( strURL )
    Dim strResult
    Dim http
    http = CreateOleObj("MSXML2.XMLHTTP") 
    http.Open( "GET", strURL, False )
    http.Send

    strResult = http.ResponseText
    Result = strResult
Fend
```

### 実行結果
```
人力検索はてな - 質問一覧
麻生太郎氏が７月25日に行った講演の中で「元気な高齢..
軽箱バンに乗られたことがある方にお伺いします。 ４..
ザ･ドリフターズをはじめとする歴代のエンターティナ..
ホンダのアクティーバンのような、軽自動車を探してい..
警察官は全国に約25万人いるそうですが、指定職は何個..
過去一年の間に、何かしらの寄付をしたことがあります..
【器が小さい？】小泉ジュニアvsあいのり横粂【パフォ..
youtubeの検索タグを、様々な言語でヒットする様にし..
３０代４０代の方に質問します。 「８時だよ全員集合..
世界には助けを必要としている人が大勢います。たとえ..
http://image.searchina.ne.jp/view.cgi?d=0200658&p=..
鉱物の採取について質問です。 京都亀岡で採取できる..
iPhoneアプリの開発はマックでしかないんですかね？..
iphone3G（OS:3.0）を使っており、キーボード設定につ..
自分の父（母）の兄（姉）は、自分の伯父（伯母）です..
２，３日前からですが、耳で聞こえる電子的な音(CDを..
ホットペッパーの各お店のトップにあるflashのような..
「まさか…」「何でそこで、それなの…？」と思うよう..
NCネットワークは、SBR（旧　 テレウェイブリンクス ..
自分の脳内で内部対話をするとき、それは神経生理学的..
```

### 参考情報
- XML DOM Objects/Interfaces - MSDN Library (英語)
- 無料のVBScriptでXMLプログラミング - @IT
