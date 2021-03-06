VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummaryUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Class module:SummaryUtl
'Version:1.0.0

'このクラスモジュールは集計用のメソッドを提供します。
'
'+------------------------------------------------------------------+





'+--Public Functions------------------------------------------------+


'指定したシートの列の数値データを集計します。
'カラム名は"C"のように指定する。

Function sumColumn(columnName As String, Optional sheetName As Variant) As Long

    Dim columnName_forRange As String
    columnName_forRange = columnName + ":" + columnName
 
   If IsMissing(sheetName) Then
    sumColumn = WorksheetFunction.Subtotal(9, Worksheets(ActiveSheet.Name).Range(columnName_forRange))
  
   Else
    sumColumn = WorksheetFunction.Subtotal(9, Worksheets(sheetName).Range(columnName_forRange))
   End If
    
End Function




'+------------------------------------------------------------------+






