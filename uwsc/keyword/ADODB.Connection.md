# ADODB.Connection
## 概要
ADODB.Connection は ADO Connection オブジェクトを表す ProgID です。

Connection オブジェクトは、データソースへのセッションを表します。

### サンプル
```vbscript
//  File : ShowTableData.uws
//  Description : UWSC から ADO を使用して SQL Server のテーブルを表示するサンプル
Option Explicit

Main()

Procedure Main()
    Dim cn
    cn = CreateOleObj("ADODB.Connection")
    
    Dim strConnection
    Dim strSERVER
    Dim strDBNAME
    Dim strUID
    Dim strPWD
    
    strSERVER = "(local)"           //  接続先のサーバー名
    strDBNAME = "AdventureWorks"    //  接続先の DB 名（例は SQL Server 2005 のサンプルデータベース）
    strUID    = "sa"                //  DB アカウント（SQL Server 認証の場合）
    strPWD    = "passwordpassword"  //  DB パスワード（SQL Server 認証の場合）
    
    strConnection = "PROVIDER=SQLOLEDB;SERVER=" + strSERVER + ";DATABASE=" + strDBNAME + ";"
    cn.Open( strConnection, strUID, strPWD )

    Dim strTableName
    strTableName = "HumanResources.Employee" //  テーブル名（例は SQL Server 2005 のサンプルデータベースのテーブル）
    ShowTableData( cn, strTableName )
Fend

Procedure ShowTableData( cn, strTableName )
    Dim strSQL
    strSQL = "SELECT * FROM " + strTableName
    
    Dim rs
    rs = cn.Execute( strSQL )
    
    //  データ表示
    Dim strDelimiter
    strDelimiter = "," //  区切り文字を指定
    ShowRecordset( rs, strDelimiter )
Fend

//  データ表示
Procedure ShowRecordset( rs, strDelimiter )
    //  タイトル行
    Dim strLine
    strLine = GetFieldNameList( rs, strDelimiter )
    DebugPrint( strLine ) //  タイトル行を出力
    
    //  データ行
    While !rs.BOF And !rs.EOF
        strLine = GetFieldValueList( rs, strDelimiter )
        DebugPrint( strLine ) //  データ行を出力
        rs.MoveNext
    Wend
Fend

//  カラム名の一覧（１行分）を取得
Function GetFieldNameList( rs, strDelimiter )
    Dim strResult
    Dim bFirst
    bFirst = True
    Dim fld
    Dim i
    For i = 0 To rs.Fields.Count-1
        fld = rs.Fields(i)
        Ifb bFirst Then
            strResult = Chr(34) + fld.Name + Chr(34)
            bFirst = False
        Else
            strResult = strResult + strDelimiter + Chr(34) + fld.Name + Chr(34)
        EndIf
    Next
    Result = strResult
Fend

//  カラムデータの一覧（１行分）を取得
Function GetFieldValueList( rs, strDelimiter )
    Dim strResult
    Dim bFirst
    bFirst = True
    Dim fld
    Dim i
    For i = 0 To rs.Fields.Count-1
        fld = rs.Fields(i)
        Ifb bFirst Then
            strResult = Chr(34) + fld.Value + Chr(34)
            bFirst = False
        Else
            strResult = strResult + strDelimiter + Chr(34) + fld.Value + Chr(34)
        EndIf
    Next
    Result = strResult
Fend

Procedure DebugPrint( strMessage )
    Print strMessage
Fend
```
### 実行結果
```
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
