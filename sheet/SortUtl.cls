VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'+--Overview--------------------------------------------------------+
'Class module:SortUtl
'Version:1.0.0

'このクラスモジュールは並び替え用のメソッドを提供します。
'
'+------------------------------------------------------------------+



'+--Public Functions------------------------------------------------+

'指定したシートにて３つの条件を元に行を並び替えます。

Function sortRows(key1_v As String, order_key1 As Variant, key2_v As String, order_key2 As Variant, key3_v As String, order_key3 As Variant, header_v As Variant, sheetName As Variant)

    Worksheets(sheetName).Cells.sort key1:=Range(key1_v), Order1:=order_key1, key2:=Range(key2_v), Order2:=order_key2, key3:=Range(key3_v), order3:=order_key3, header:=header_v

End Function


'+------------------------------------------------------------------+
