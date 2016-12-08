VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Class module:WindowUtl
'Version:1.0.0

'このクラスモジュールはウィンドウ操作用のメソッドを提供します。
'
'+------------------------------------------------------------------+


'+--Public Functions------------------------------------------------+


'シート名とセルを指定しウィンドウ枠を固定します。
'
'引数sheetName省略時はアクティブなシートが選択されます。

Function freeze(cell As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range(cell).Select
        ActiveWindow.FreezePanes = True
    Else
        Worksheets(sheetName).Range(cell).Select
        ActiveWindow.FreezePanes = True
    End If

End Function



'シート名を指定し先頭行を固定します。
'
'すでにウィンドウ枠が固定されている場合は固定し直さない。
'ウィンドウ枠をセットし直す場合は１度クリアして下さい。
'引数sheetName省略時はアクティブなシートが選択されます。

Function freezeHeadRow(Optional sheetName As Variant)
    
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Activate
        Worksheets(ActiveSheet.Name).Rows("2").Select
        ActiveWindow.FreezePanes = True
    Else
         Worksheets(sheetName).Activate
        Worksheets(sheetName).Rows("2").Select
        ActiveWindow.FreezePanes = True
    End If
 
End Function



'シート名を指定しウィンドウ枠を解除します。
'
'引数sheetName省略時はアクティブなシートが選択されます。

Function clearFreezePanes(Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        ActiveWindow.FreezePanes = False
    Else
        Worksheets(sheetName).Activate
        ActiveWindow.FreezePanes = False
    End If
  
End Function



'シート名と列を指定し列幅を最適化します。
'
'引数sheetName省略時はアクティブなシートが選択されます。
'
'引数例:A列からI列まで　autoFitColumns "A:I", "Sheet1" 、1列のみ　autoFitColumns "C:C"

Function autoFitColumns(columnsName As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).columns(columnsName).autoFit
    Else
        ''ワークシート指定ありの場合
        Worksheets(sheetName).columns(columnsName).autoFit
    End If
 
End Function



'シート名と行を指定し行幅を最適化します。
'
'引数sheetName省略時はアクティブなシートが選択されます。

Function autoFitRows(rowsName As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Rows(rowsName).autoFit
    Else
        ''ワークシート指定ありの場合
        Worksheets(sheetName).Rows(rowsName).autoFit
    End If
 
End Function



'シート名と列を指定し列幅を変更します。
'
'引数sheetName省略時はアクティブなシートが選択されます。

Function setColumnsWidth(columnsName As String, columsWidth As Integer, Optional sheetName As Variant)
    
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).columns(columnsName).ColumnWidth = columsWidth
    Else
        ''ワークシート指定ありの場合
        Worksheets(sheetName).columns(columnsName).ColumnWidth = columsWidth
    End If
    
End Function



'シート名と行を指定し行幅を変更します。
'引数sheetName省略時はアクティブなシートが選択されます。
'
'引数例:行1から行3まで　setRowsHeight "1:3", "Sheet1" 、行1のみ　setRowsHeight "1:1"

Function setRowsHeight(rowsName As String, rowsHeight As Integer, Optional sheetName As Variant)
    
    If IsMissing(sheetName) Then
         Worksheets(ActiveSheet.Name).Rows(rowsName).rowHeight = rowsHeight
    Else
         Worksheets(sheetName).Rows(rowsName).rowHeight = rowsHeight
    End If
    
End Function




'+------------------------------------------------------------------+





