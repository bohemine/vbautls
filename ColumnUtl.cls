VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'+--Overview--------------------------------------------------------+
'Class module:ColumnUtl
'Version:1.0.0

'このクラスモジュールは列操作用のメソッドを提供します。
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+

'シートと列を指定し、その列を削除します。
'
'引数sheetNameのみ省略時はアクティブなシートが選択されます。

Function removeColumns(columnsName As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnsName).Delete
Else
   Worksheets(sheetName).Range(columnsName).Delete
End If
    
End Function



'シートと列を指定し、列を挿入します。
'クリップボードに列が格納されてある場合、それが貼り付けされます。
'クリップボードに列以外の有効でない範囲のデータがある場合エラーが生じる可能性があります。
'
'引数sheetNameのみ省略時はアクティブなシートが選択されます。

Function insertColumns(columnsName As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnsName).Insert
Else
   Worksheets(sheetName).Range(columnsName).Insert
End If
    
End Function



'指定した列のヘッダー名を付けます。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。
'
Function renameHeader(columnName As String, header As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnName + "1").Value = header
Else
   Worksheets(sheetName).Range(columnName + "1").Value = header
End If
    
End Function



'シートと列を指定し、その列を指定された列へコピーします。
'貼り付けも行われます。

Function copyColumns(sheetName_src As String, columnName_src As String, sheetName_dst As String, columnName_dst As String)

    Worksheets(sheetName_src).Range(columnName_src + ":" + columnName_src).copy Worksheets(sheetName_dst).Range(columnName_dst + ":" + columnName_dst)

End Function



'シートと列を指定し、その列にあるデータをクリアします。
'
'引数sheetNameのみ省略時はアクティブなシートが選択されます。

Function clearColumnsContents(columnsName As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnsName).ClearContents
Else
   Worksheets(sheetName).Range(columnsName).ClearContents
End If
    
End Function



'シートと列を指定し、その列にあるデータを切り取ります。
'
'引数sheetNameのみ省略時はアクティブなシートが選択されます。

Function cutColumns(columnsName As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnsName).Cut
Else
   Worksheets(sheetName).Range(columnsName).Cut
End If
    
End Function



'シートと列を指定し、その列を指定された列へ移動します。
'
'切り取った列が貼り付け先よりも左側にある場合、指定した列より１つ左の列に貼り付けられます。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。

Function moveColumns(columnsName_src As String, columnName_dst As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
    cutColumns columnsName_src
    insertColumns columnName_dst
Else
    cutColumns columnsName_src, sheetName
    insertColumns columnName_dst, sheetName
End If
    
End Function



'列名を指定し、その列番号返します。
'return Long
'
'切り取った列が貼り付け先よりも左側にある場合、指定した列より１つ左の列に貼り付けられます。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。

Function getColumnNumber(columnsName As String) As Long

    getColumnNumber = ActiveSheet.Range(columnsName + "1").column

End Function



'列番号を指定し、その列名返します。

'切り取った列が貼り付け先よりも左側にある場合、指定した列より１つ左の列に貼り付けられます。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。
'
Function getColumnName(columnIndex As Integer)
    
    getColumnName = Split(Cells(, columnIndex).Address, "$")(1)

End Function



'先頭行にあるフィールド名を指定し、最初に見つかった列の列名を返します。
'
'切り取った列が貼り付け先よりも左側にある場合、指定した列より１つ左の列に貼り付けられます。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。
'完全に同一の文字列を対象とします。文字列の一部がマッチしても対象になりません。

Function getColumnNameByField(target As String, Optional sheetName As Variant) As String

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Range("1:1").Find(what:=target, LookAt:=xlWhole).Activate
    getColumnNameByField = Split(ActiveCell.Address, "$")(1)

Else
    Worksheets(sheetName).Activate
    Range("1:1").Find(what:=target, LookAt:=xlWhole).Activate
    getColumnNameByField = Split(ActiveCell.Address, "$")(1)
   
End If

End Function

'+------------------------------------------------------------------+




