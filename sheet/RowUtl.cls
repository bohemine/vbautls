VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Class module:RowUtl
'Version:1.0.0

'このクラスモジュールは列操作用のメソッドを提供します。
'
'+------------------------------------------------------------------+





'+--Public Functions------------------------------------------------+


'シートと行番号を指定し削除します。
'
'sheetName省略時はアクティブなシートを対象とします。

Function removeRow(rowNo As Long, Optional sheetName As Variant)
  
   If IsMissing(sheetName) Then
       Rows(rowNo).Delete
   Else
      Sheets(sheetName).Rows(rowNo).Delete
   End If
    
End Function



'シートと行番号を範囲指定し削除します。
'
'sheetName省略時はアクティブなシートを対象とします。

Function removeRows(rowNo_from As Long, rowNo_to As Long, Optional sheetName As Variant)
  
  Dim rowNoRange As String
  rowNoRange = CStr(rowNo_from) + ":" + CStr(rowNo_to)
  
   If IsMissing(sheetName) Then
       Worksheets(ActiveSheet.Name).Range(rowNoRange).Delete
   Else
      Worksheets(sheetName).Range(rowNoRange).Delete
   End If
    
End Function



'指定したシートの先頭行を削除します。
'
'sheetName省略時はアクティブなシートを対象とします。

Function removeHeadRow(Optional sheetName As Variant)
  
   If IsMissing(sheetName) Then
       Worksheets(ActiveSheet.Name).Range("A1").Offset(0, 0).EntireRow.Delete
   Else
       Worksheets(sheetName).Range("A1").Offset(0, 0).EntireRow.Delete
   End If
    
End Function



'指定したシートの重複行を削除します。
'
'columnIndexへ対象とする列番号を渡して下さい。
'headerへ先頭行の扱いを指定して下さい。xlYes[ヘッダとして扱う]　or xlNo[ヘッダとして扱わない]　or xlGuess [Excelに判断させる]
'
'"A1"を基準セルとしたときのCurrentRegionが対象範囲です。
'Excel2007以降でのみ使用可。

Function removeDuplicateRows(columnIndex As Integer, header As Variant, sheetName As String)
  
  Worksheets(sheetName).Range("A1").CurrentRegion.RemoveDuplicates columns:=columnIndex, header:=header

End Function



'指定したシートにて2つの列を元に重複行を削除します。
'
'columnIndexへ対象とする列番号を渡して下さい。
'headerへ先頭行の扱いを指定して下さい。xlYes[ヘッダとして扱う]　or xlNo[ヘッダとして扱わない]　or xlGuess [Excelに判断させる]
'
'"A1"を基準セルとしたときのCurrentRegionが対象範囲です。
'Excel2007以降でのみ使用可。

Function removeDuplicateRows2(columnIndex1 As Integer, columnIndex2 As Integer, header As Variant, sheetName As String)
  
  Worksheets(sheetName).Range("A1").CurrentRegion.RemoveDuplicates columns:=Array(columnIndex1, columnIndex2), header:=header

End Function



'TODO: 数値引数対応
'
'シートと行を指定し、指定された行へコピーします。
'
'貼り付けも行われます。

Function copyRow(sheetName_src As String, rowName_src As String, sheetName_dst As String, rowName_dst As String)

    Worksheets(sheetName_src).Range(rowName_src + ":" + rowName_src).copy Worksheets(sheetName_dst).Range(rowName_dst + ":" + rowName_dst)
 
End Function

'+------------------------------------------------------------------+














