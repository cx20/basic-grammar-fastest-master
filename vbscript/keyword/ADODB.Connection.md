# ADODB.Connection
## 概要
ADODB.Connection は ADO Connection オブジェクトを表す ProgID です。

Connection オブジェクトは、データソースへのセッションを表します。

### サンプル
```vbscript
' File : ShowTableData.vbs
' Usage : CScript //Nologo ShowTableData.vbs
' Description : VBScript から ADO を使用して SQL Server のテーブルを表示するサンプル
Option Explicit

Call Main()

Sub Main()
    Dim cn
    Set cn = CreateObject("ADODB.Connection")
    
    Dim strConnection
    Dim strSERVER
    Dim strDBNAME
    Dim strUID
    Dim strPWD
    
    strSERVER = "(local)"           ' 接続先のサーバー名
    strDBNAME = "AdventureWorks"    ' 接続先の DB 名（例は SQL Server 2005 のサンプルデータベース）
    strUID    = "sa"                ' DB アカウント（SQL Server 認証の場合）
    strPWD    = "passwordpassword"  ' DB パスワード（SQL Server 認証の場合）
    
    strConnection = "PROVIDER=SQLOLEDB;SERVER=" & strSERVER & ";DATABASE=" & strDBNAME & ";"
    cn.Open strConnection, strUID, strPWD

    Dim strTableName
    strTableName = "HumanResources.Employee" ' テーブル名（例は SQL Server 2005 のサンプルデータベースのテーブル）
    Call ShowTableData( cn, strTableName )
End Sub

Sub ShowTableData( cn, strTableName )
    Dim strSQL
    strSQL = "SELECT * FROM " & strTableName
    
    Dim rs
    Set rs = cn.Execute( strSQL )
    
    ' データ表示
    Dim strDelimiter
    strDelimiter = "," ' 区切り文字を指定
    Call ShowRecordset( rs, strDelimiter )
End Sub

' データ表示
Sub ShowRecordset( rs, strDelimiter )
    ' タイトル行
    Dim strLine
    strLine = GetFieldNameList( rs, strDelimiter )
    DebugPrint strLine ' タイトル行を出力
    
    ' データ行
    While Not rs.BOF And Not rs.EOF
        strLine = GetFieldValueList( rs, strDelimiter )
        DebugPrint strLine ' データ行を出力
        rs.MoveNext
    Wend
End Sub

' カラム名の一覧（１行分）を取得
Function GetFieldNameList( rs, strDelimiter )
    Dim strResult
    Dim bFirst
    bFirst = True
    Dim fld
    For Each fld In rs.Fields
        If bFirst Then
            strResult = Chr(34) & fld.Name & Chr(34)
            bFirst = False
        Else
            strResult = strResult & strDelimiter & Chr(34) & fld.Name & Chr(34)
        End If
    Next
    GetFieldNameList = strResult
End Function

' カラムデータの一覧（１行分）を取得
Function GetFieldValueList( rs, strDelimiter )
    Dim strResult
    Dim bFirst
    bFirst = True
    Dim fld
    For Each fld In rs.Fields
        If bFirst Then
            strResult = Chr(34) & fld.Value & Chr(34)
            bFirst = False
        Else
            strResult = strResult & strDelimiter & Chr(34) & fld.Value & Chr(34)
        End If
    Next
    GetFieldValueList = strResult
End Function

Sub DebugPrint( strMessage )
    ' WSH で実行する場合
    WScript.Echo strMessage
    ' VBA で実行する場合
    ' Debug.Print strMessage
End Sub
```

### 実行結果
```bat
"EmployeeID","NationalIDNumber","ContactID","LoginID","ManagerID","Title","BirthDate","MaritalStatus","Gender","HireDate","SalariedFlag","VacationHours","SickLeaveHours","CurrentFlag","rowguid","ModifiedDate"
"1","14417807","1209","adventure-works\guy1","16","Production Technician - WC60","1972/05/15","M","M","1996/07/31","False","21","30","True","{AAE1D04A-C237-4974-B4D5-935247737718}","2004/07/31"
"2","253022876","1030","adventure-works\kevin0","6","Marketing Assistant","1977/06/03","S","M","1997/02/26","False","42","41","True","{1B480240-95C0-410F-A717-EB29943C8886}","2004/07/31"
"3","509647174","1002","adventure-works\roberto0","12","Engineering Manager","1964/12/13","M","M","1997/12/12","True","2","21","True","{9BBBFB2C-EFBB-4217-9AB7-F97689328841}","2004/07/31"
     ：
 （以下省略）
 ```
 
### 参考情報
- ADO API リファレンス - MSDN Library
- ActiveX Data Objects (ADO) Frequently Asked Questions (2007/03/27 Rev:4.3) - Microsoft KB (英語)
- ASP で ADO データ接続コードを作成する方法 - Microsoft KB
