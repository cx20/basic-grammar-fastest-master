# CDO.Message
## 概要
CDO.Message は、CDO（Collaboration Data Objects）のメッセージ オブジェクトを表す ProgID です。

CDO を使用することで、メールの送信が可能です。

### サンプル
```vbscript
//  File : SendTestMail.uws
//  Description : UWSC から CDO を使用してメールを送信するサンプル
Option Explicit

Const CDO_SEND_USING        = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Const CDO_SMTP_SERVER       = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const CDO_SMTP_SERVERPORT   = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Const CDO_SMTP_AUTHENTICATE = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
Const CDO_SEND_USERNAME     = "http://schemas.microsoft.com/cdo/configuration/sendusername"
Const CDO_SEND_PASSWORD     = "http://schemas.microsoft.com/cdo/configuration/sendpassword"

Main()

Procedure Main()
    Dim msg
    msg = CreateOleObj("CDO.Message")
    Dim cfg
    cfg = msg.Configuration
    
    Dim nSendUsing
    Dim strServer
    Dim nPort
    Dim nAuthenticate
    Dim strUserName
    Dim strPassword
    
    nSendUsing    = 2                   2: SMTP を使用
    strServer     = "smtp.hoge.co.jp"   SMTP サーバー
    nPort         = 25                  SMTP サーバーのポート
    nAuthenticate = 1                   1: Basic 認証
    strUserName   = "username"          ユーザー名
    strPassword   = "password"          パスワード

    SetSmtpServer( cfg, nSendUsing, strServer, nPort, nAuthenticate, strUserName, strPassword )

    msg.From     = "foo@hoge.co.jp"     送信元メールアドレス
    msg.To       = "bar@hoge.co.jp"     送信先メールアドレス
    msg.Subject  = "test mail"          件名
    msg.TextBody = "test message"       メール本文
    msg.Send
Fend

Procedure SetSmtpServer( Var cfg, nSendUsing, strServer, nPort, nAuthenticate, strUserName, strPassword )
    cfg.Fields.Item(CDO_SEND_USING)        = nSendUsing
    cfg.Fields.Item(CDO_SMTP_SERVER)       = strServer
    cfg.Fields.Item(CDO_SMTP_SERVERPORT)   = nPort
    cfg.Fields.Item(CDO_SMTP_AUTHENTICATE) = nAuthenticate
    cfg.Fields.Item(CDO_SEND_USERNAME)     = strUserName
    cfg.Fields.Item(CDO_SEND_PASSWORD)     = strPassword
    cfg.Fields.Update
Fend
```

### 実行結果
```
Return-Path: <foo@hoge.co.jp>
Delivered-To: bar@hoge.co.jp
Received: (qmail 24624 invoked from network); 22 Jun 2008 16:49:12 -0000
Received: from unknown (HELO xxxxxxxx) (xxx.xxx.xxx.xxx)
  by 0 with SMTP; 22 Jun 2008 16:49:12 -0000
thread-index: AcjUh+kfhxf6BMsMSieB6IX+cFWXyQ==
Thread-Topic: test mail
From: <foo@hoge.co.jp>
To: <bar@hoge.co.jp>
Subject: test mail
Date: Mon, 23 Jun 2008 01:49:17 +0900
Message-ID: <8DC719AC384D4DEE9D14F32D403B96C6@xxxxxxxx>
MIME-Version: 1.0
Content-Type: text/plain
Content-Transfer-Encoding: 7bit
X-Mailer: Microsoft CDO for Windows 2000
Content-Class: urn:content-classes:message
Importance: normal
Priority: normal
X-MimeOLE: Produced By Microsoft MimeOLE V6.00.2900.5512
```

### 利用可能な CDO のバージョン
|OS|6.1|6.2|6.5|6.6|
|:-|:--|:--|:--|:--|
|Microsoft Windows 2000|○|　|　|　|
|Microsoft Windows XP|　|○|　|　|
|Microsoft Windows Server 2003|　|　|○|　|
|Microsoft Windows Vista|　|　|　|○|
|Microsoft Windows Server 2008|　|　|　|○|
|Microsoft Windows 7|　|　|　|○|
|Microsoft Windows 8|　|　|　|○|

### 参考情報
- CDO Library > リファレンス - MSDN Library
- ＠IT：Windows TIPS -- Tips：Windows標準機能とWSHを使ってメールを送信する (2004/05/22) - @IT
- CDO を使用して送信メールにファイルを添付するにはどうすればよいでしょうか。 - Microsoft TechNet
- Where to acquire the CDO Libraries (all versions) - Microsoft KB (英語)
