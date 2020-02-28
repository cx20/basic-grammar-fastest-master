# WbemScripting.SWbemLocator
## 概要
WbemScripting.SWbemLocator は WBEM Scripting Locator を表す ProgID です。

この ProgID は、VBScript から WMI を使用する場合に使用します。

WMI (Windows Management Instrumentation) は、WBEM と CIM の Microsoft による実装です。

WMI を用いることでネットワークを通じてコンピュータ等の機器を管理することができます。

### サンプル
```vbscript
' File : ExportOcxListByWMI.vbs
' Usage : CScript //Nologo ExportOcxListByWMI.vbs
' Description : VBScript から WMI を使用して ActiveX Control の一覧を出力するサンプル
Option Explicit

Call Main()

Sub Main()
    Dim wmi
    Set wmi = CreateObject("WbemScripting.SWbemLocator")
    
    Dim strComputer
    Dim strNamespace
    Dim strUser
    Dim strPassword
    
    '＜ローカル接続する場合＞
    strComputer = "."
    strNamespace = "root\cimv2"
    
    '＜リモート接続する場合＞
    'strComputer = "ComputerName"
    'strNamespace = "root\cimv2"
    'strUser = "UserName"
    'strPassword = "Password"

    Dim svc
    Set svc = wmi.ConnectServer( strComputer, strNamespace, strUser, strPassword )
    'ユーザー資格情報を指定する必要がない場合は以下の指定でも可能です。
    '"winmgmts:～" は WMI のオブジェクトを取得する為の「モニカ文字列」（リソースに対するアクセス識別子）です。
    'Set svc = GetObject("winmgmts:\\" & strComputer & "\" & strNamespace)

    Dim strSQL
    strSQL = "SELECT * FROM Win32_ClassicCOMClassSetting WHERE Control = True"
    Dim items
    Set items = svc.ExecQuery(strSQL)
    
    Call ShowObjectset( items )
End Sub

Sub ShowObjectset( items )
    DebugPrint _
        "CLSID" & vbTab & _
        "ProgId" & vbTab & _
        "InprocServer32" & vbTab & _
        "LocalServer32" & vbTab & _
        "Description"
    
    Dim item
    For Each item In items
        DebugPrint _
            item.ComponentId & vbTab & _
            item.ProgId & vbTab & _
            item.InprocServer32 & vbTab & _
            item.LocalServer32 & vbTab & _
            item.Description
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
|CLSID|ProgId|InprocServer32|LocalServer32|Description|
|:----|:-----|:-------------|:------------|:----------|
|{00024522-0000-0000-C000-000000000046}|RefEdit.Ctrl|C:\PROGRA\~1\MICROS\~3\OFFICE11\REFEDIT.DLL|　|RefEdit.Ctrl|
|{00028C00-0000-0000-0000-000000000046}|MSDBGrid.DBGrid|C:\WINDOWS\SYSTEM32\DBGRID32.OCX|　|DBGrid Control|
|{0002E541-0000-0000-C000-000000000046}|OWC10.Spreadsheet.10|C:\PROGRA\~1\COMMON\~1\MICROS\~1\WEBCOM\~1\10\OWC10.DLL|　|Microsoft Office Spreadsheet 10.0|
|{0002E542-0000-0000-C000-000000000046}|OWC10.PivotTable.10|C:\PROGRA\~1\COMMON\~1\MICROS\~1\WEBCOM\~1\10\OWC10.DLL|　|Microsoft Office PivotTable 10.0|
|{0002E543-0000-0000-C000-000000000046}|OWC10.DataSourceControl.10|C:\PROGRA\~1\COMMON\~1\MICROS\~1\WEBCOM\~1\10\OWC10.DLL|　|Microsoft Office Data Source Control 10.0|
|{0002E546-0000-0000-C000-000000000046}|OWC10.ChartSpace.10|C:\PROGRA\~1\COMMON\~1\MICROS\~1\WEBCOM\~1\10\OWC10.DLL|　|Microsoft Office Chart 10.0|

　　：

以下省略

### 参考情報
- WMI Reference - MSDN Library (英語)
- WMI スクリプト入門 : 第 1 部 (2002/10/02) - MSDN Library
- Windows Management Instrumentation (WMI) : よく寄せられる質問 (2005/02/22) - Microsoft TechNet
- スクリプト センター > スクリプト一覧 > スクリプト テクニック > WMI - Microsoft TechNet
- Scripting Eye for the GUI Guy: WBEMTEST - Microsoft TechNet
- スクリプト センター : Scriptomatic ツール - Microsoft TechNet
- Windows Management Instrumentation - Wikipedia
- WMIとwmicコマンドを使ってシステムを管理する（基本編）(2008/04/18) - @IT
