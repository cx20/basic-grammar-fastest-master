# VBScript 基礎文法最速マスター

VBScript の文法一覧です。他の言語をある程度知っている人はこれを読めば VBScript の基礎をマスターして VBScript を書くことができるようになっています。簡易リファレンスとしても利用できると思いますので、これは足りないと思うものがあれば教えてください。

## 0. はじめに

　VBScript は Visual Basic のサブセットの為、すべてのステートメント、組み込み関数をサポートしているわけではありません。

　例えば、VBScript には型がなかったり（すべて Variant 型）、Format 関数がなかったりします。

　しかし、Windows 標準で使える（インストール無しで使用できる）スクリプト言語としては十分に強力なものとなっています。

　メモ帳さえあれば開発はできますので、試してみてはいかがでしょうか？

## 1. 基礎
### 変数宣言の強制

　ソースコードの先頭に「Option Explicit」を入れるようにしましょう。変数の宣言が強制されコードの品質も上がります。
```vbscript
Option Explicit
```
### メッセージの表示
```vbscript
WScript.Echo "Hello, World." ' コンソールにメッセージを出力します
MsgBox "Hello, World." ' メッセージボックスにメッセージを表示します
```
### コメント

　コメントは「'」を用います。「Rem」も使用可能です。
```vbscript
' コメントその１
Rem コメントその２
```
### 変数の宣言

　変数の宣言は「Dim」を用います。固定サイズ配列の場合は「(10)」のようにサイズを指定します。サイズを指定しない場合は動的配列となります。

　VBScript の変数自体はバリアント型（Variant）で代入する値に応じて内部形式が変化します。
```vbscript
Dim num ' 変数
Dim students(10) ' 配列変数（配列は0～10となり要素数は11になります）
Dim students() ' 動的配列
```
### スクリプトの実行

　スクリプトを実行するにはコマンドラインで次のようにします。
```bat
CScript script.vbs
```
　出力結果をファイルに書き出すにはリダイレクトを使います。
```bat
CScript script.vbs > file.txt
```
### デバッガの起動

　デバッガを起動するにはコマンドラインで次のようにします。

　ただし、事前に Visual Studio 等のデバッガがインストールされている必要があります。
```bat
CScript script.vbs //X
```
## 2. 数値
### 数値の表現

　変数には整数でも小数でも代入できます。代入する値に応じて、変数の内部形式が変化します。
```vbscript
num = 1 ' 整数型（Integer）
num = 1.234 ' 倍精度浮動小数点型（Double）
num = 1000000000 ' 長整数型（Long）
```
### 四則演算

　四則演算です。
```vbscript
num = 1 + 1 ' 2
num = 1 - 1 ' 0
num = 1 * 2 ' 2
num = 1 / 2 ' 0.5
```
　商と余りの求め方です。
```vbscript
num = 3 \ 2 ' 1（商）
num = 3 Mod 2 ' 1（余り）
```
### インクリメントとデクリメント

　インクリメントとデクリメントです。
```vbscript
' インクリメント
i = i + 1

' デクリメント
i = i - 1
```

## 3. 文字列
### 文字列の表現

　文字列は「"」ダブルクォーテーションで囲みます。変数にタブや改行コードをセットするには「vbTab」や「vbCrLf」を用います。
```vbscript
str1 = "abc"
str2 = "a" & vbTab & "bc" & vbCrLf
```
### 文字列操作

　各種文字列操作です。
```vbscript
' 結合
str1 = "aaa" & "bbb" ' 文字列の連結
str2 = Join( Array("aaa", "bbb", "ccc"), "," ) ' 区切り文字に「,」（カンマ）を指定する場合

' 分割
record = Split("aaa,bbb,ccc", ",")

' 長さ
length = Len("abcdef") ' 6
length = Len("あいうえお") ' 5（文字数を数えるには Len 関数を使用します）
length = LenB("あいうえお") ' 10（文字のバイト数を数えるには LenB 関数を使用します）

' 切り出し
str = Mid("abcd", 1, 2) ' ab（1桁目から2文字）

' 検索
result = InStr("abcd", "cd") ' 見つかった場合はその位置、見つからなった場合は 0 が返ります
result = InStr("あいうえお", "うえ") ' 3（文字数で扱う場合は InStr 関数を使用します）
result = InStrB("あいうえお", "うえ") ' 5（バイト数で扱う場合は InStrB 関数を使用します）
```
## 4. 配列
### 配列変数の宣言と代入

　固定サイズ配列として宣言する場合
```vbscript
Dim ary(2)
ary(0) = 100
ary(1) = 200
ary(2) = 300
```
　動的配列として宣言する場合
```vbscript
Dim ary()
ReDim ary(2)
ary(0) = 100
ary(1) = 200
ary(2) = 300
```
　変数に Array 関数を使用して配列をセットする場合
```vbscript
Dim ary
ary = Array( 100, 200, 300 )
```
### 配列の要素の参照と代入
```vbscript
a = ary(0) ' 100
b = ary(1) ' 200

ary(0) = 1
ary(1) = 2
```
### 要素の個数
```vbscript
n = UBound(ary) - LBound(ary) + 1 ' 配列の上限 - 下限
```
### 配列の操作
```vbscript
Dim ary
ary = Array( 1, 2, 3 )

' 先頭を取り出す
a = ary(0) ' a は 1
' 末尾を取り出す
b = ary(UBound(ary)) ' b は 3
' 末尾に追加
ReDim Preserve ary(UBound(ary) + 1) ' 固定サイズ配列の場合は追加できません
ary(UBound(ary)) = 9 ' ary は [1,2,3,9] に
```
## 5. Dictionary オブジェクト

VBScript にはハッシュ変数はありませんが、Scripting.Dictionary オブジェクトを用いることで代替えが可能です。
### Dictionary オブジェクトの宣言と代入
```vbscript
Dim hash
Set hash = CreateObject("Scripting.Dictionary")
hash.Add "a", 1
hash.Add "b", 2
```
### Dictionary の要素の参照と代入
```vbscript
' 要素の参照
WScript.Echo hash("a") ' 1
WScript.Echo hash("b") ' 2

' 要素の代入
hash("a") = 5
hash("b") = 7
```
### Dictionary のプロパティとメソッド
```vbscript
' キーの取得
keys = hash.Keys

' 値の取得
values = hash.Items

' キーの存在確認
hash.Exists("a")

' キーの削除
hash.Remove "a"
```
## 6. 制御文
### If 文

　If 文です。１行に書く形式とブロック形式の構文が利用できます。
```vbscript
If 条件 Then 式 [Else 式]
```
　ブロック形式の If 文には End If が必要です。
```vbscript
If 条件 Then
式
[Else]
式
End If
```
### If ～ ElseIf 文

　If ～ ElseIf 文です。「Else If」ではなく「ElseIf」である（Else と If の間に空白は入らない）ことに注意しましょう。
```vbscript
If 条件 Then
式
[ElseIf 条件 Then]
式
[Else]
式
End If
```
### Do ～ Loop 文

　Do ～ Loop 文です。Do While ～（真の間ループ）や Do Until ～（真になるまでループ）が利用できます。
```vbscript
i = 0
Do While i < 5

  ' 処理

i = i + 1
Loop
```
### For 文

　For文です。
```vbscript
For i = 0 To 4

Next
```
### For Each 文

　For Each 文です。配列やコレクションオブジェクトの各要素を参照するときに使用します。
```vbscript
For Each field In fields

Next
```
### 比較演算子

　比較演算子の一覧です。「==」でなく「=」であることに注意してください。
```vbscript
num1 = num2 ' num1 は num2 と等しい
num1 <> num2 ' num1 は num2 と等しくない
num1 < num2 ' num1 は num2 より小さい
num1 > num2 ' num1 は num2 より大きい
num1 <= num2 ' num1 は num2 以下
num1 >= num2 ' num1 は num2 以上
```
## 7. サブルーチン
### Sub プロシージャ

戻り値を返さない処理は Sub プロシージャで定義します。
```vbscript
Sub Show_Sum( num1, num2 )
    Dim total
    total = num1 + num2
    WScript.Echo total
End Sub
```
### Function プロシージャ

　戻り値を返す処理は Function プロシージャで定義します。
```vbscript
Function Sum( num1, num2 )
    Dim total
    total = num1 + num2
    Sum = total ' 戻り値を指定
End Function
```
## 8. ファイル入出力

　VBScript でファイルの入出力を行うには Scripting.FileSystemObject オブジェクトを使用します。
```vbscript
Const ForReading = 1

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("C:\temp\hoge.txt", ForReading)
Do Until file.AtEndOfStream
    strLine = file.ReadLine ' ファイルの１行分を読み込みます
Loop

file.Close
```
## 9. 知っておいたほうがよい文法

VBScript でよく出てくる知っておいたほうがよい文法の一覧です。
### オブジェクトの代入と破棄

　オブジェクト変数に代入する場合は Set 文を使用します。
```vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
```
　オブジェクト変数を破棄する場合は Set 文で Nothing をセットします。
```vbscript
Set fso = Nothing
```
### Null と Empty と Nothing の違い

　Null … 変数に有効な値が変数に格納されていないことを示します。DB の Null 値に相当します。

　Empty … 変数が初期化されていない状態を表す値です。変数を宣言した直後はこの値です。

　Nothing … オブジェクト変数がオブジェクトを参照していない状態を表す値です。
VBScript の真偽値

　VBScript で偽と判断される値は、「False」「0」「Null」「Empty」「"False"」「"0"」です。これ以外は真になります。
ByVal と ByRef

　VBScript の引数には、値渡し（ByVal）と参照渡し（ByRef）があります。省略した場合は参照渡しになります。
```vbscript
Dim num1, num2, num3
num1 = 0
num2 = 0
num3 = 0

RefTest num1, num2, num3

WScript.Echo num1 ' 1
WScript.Echo num2 ' 0
WScript.Echo num3 ' 3

Sub RefTest( num1, ByVal num2, ByRef num3 )
    num1 = 1 ' 参照渡しの為、値は更新されます
    num2 = 2 ' 値渡しの為、値は更新されません
    num3 = 3 ' 参照渡しの為、値は更新されます
End Sub
```
### While ～ Wend 文

　While ～ Wend 文です。Do ～ Loop 文の方が柔軟性がありますが、覚えておいて損はありません。
```vbscript
i = 0
While i < 5

    ' 処理

    i = i + 1
Wend
```
### コマンドライン引数

　WScript.Arguments でコマンドラインの情報を取得できます。

　また、Arguments オブジェクトの Named プロパティを使用することで名前付き引数が利用可能です。
```vbscript
For i = 0 To WScript.Arguments.Count - 1
    WScript.Echo WScript.Arguments(i)
Next
```
### エラー処理

　VBScript では、実行時にエラーがあると処理が停止します。

　処理を停止させず続行させる場合には On Error Resume Next を用います。エラー内容は Err オブジェクトで判断します。
```vbscript
On Error Resume Next
num = 1 / 0 ' 0 除算
If Err.Number <> 0 Then
    WScript.Echo "Err.Source = [" & Err.Source & "]" ' Microsoft VBScript 実行時エラー
    WScript.Echo "Err.Description = [" & Err.Description & "]" ' 0 で除算しました。
    WScript.Echo "Err.Number = [" & Err.Number & "]" ' 11
End If
```
### 外部スクリプトのインクルード

　ExecuteGlobal ステートメントを使用することにより外部スクリプトをロードすることが可能です。
```vbscript
Option Explicit

Call Main()

Sub Main()
    Include "inc.vbs" ' 外部スクリプトをロード

    WScript.Echo Sum( 1, 2 )
End Sub

Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
End Function

' File : inc.vbs
Function Sum( num1, num2 )
    Sum = num1 + num2
End Function
```
### インタラクティブVBScript

　WSH は残念ながら対話型シェルの機能はありませんが、以下の 10 行ほどのプログラムを用いることで VBScript を対話的に実行することができるようになります。
```vbscript
' File : ivb.vbs
' Usage : CScript ivb.vbs
Do While True
    WScript.StdOut.Write(">>> ")
    ln = Wscript.StdIn.ReadLine
    If LCase(Trim(ln)) = "exit" Then Exit Do
    On Error Resume Next
    Err.Clear
    Execute ln
    If Err.Number <> 0 Then WScript.Echo(Err.Description)
    On Error Goto 0
Loop
```
　コマンドラインより「CScript ivb.vbs」と入力することで実行できます。終了は「exit」です。
```bat
C:\Users\cx20\edu\VBScript\ivb>CScript ivb.vbs
Microsoft (R) Windows Script Host Version 5.7
Copyright (C) Microsoft Corporation 1996-2001. All rights reserved.

>>> WScript.Echo "Hello"
Hello
>>> WScript.Echo 1+1
2
>>> exit
```
## 参考情報

- VBScript ランゲージ リファレンス - MSDN ライブラリ
- Windows Script 5.6 ドキュメント ダウンロード (exe 形式; 1.67 MB)
- VBA の機能で VBScript に含まれていない機能 - MSDN ライブラリ
- VBScript - Wikipedia
