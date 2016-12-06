VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridLineUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Class module:GridlineUtl
'Version:1.0.0
'
'+------------------------------------------------------------------+



'+--Const List--------------------------------------------------------+
'
'[LineStyle]
'実線 (細): xlContinuous
'破線:      xlDash
'一点鎖線:  xlDashDot
'二点鎖線:  xlDashDotDot
'点線:      xlDot
'二重線:    xlDouble
'斜め斜線:  xlSlantDashDot
'無し:      xlLineStyleNone
'
'[Weight]
'極細:      xlHairline
'細:        xlThin
'中:        xlMedium
'太:        xlThick
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+

'writeGrid

'>シートとセル範囲を指定し、格子状に罫線を引きます。
'引数range_v省略時は、アクティブなセル範囲を対象とします。
'引数range_v省略時は、アクティブシートになるため引数sheetNameも省略する必要があります。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。
'

Function writeGrid(lineStyle_v As Variant, weight_v As Variant, Optional range_v As Variant, Optional sheetName As Variant)

If IsMissing(range_v) Then

    '範囲指定が引数として渡されず、アクティブな領域を対象とする場合
    With Selection.Borders
        .LineStyle = lineStyle_v
        .Weight = weight_v
    End With
Else
    
    '範囲指定が引数として渡される場合、sheetNameが渡されているかどうかで分岐する。

    If IsMissing(sheetName) Then
    
      With Worksheets(ActiveSheet.Name).Range(range_v).Borders
        .LineStyle = lineStyle_v
        .Weight = weight_v
      End With
        
    Else
    
      With Worksheets(sheetName).Range(range_v).Borders
        .LineStyle = lineStyle_v
        .Weight = weight_v
      End With
    
    End If

End If
    
End Function

'writeGridEdge

'>シートとセル範囲を指定し、セル範囲の上下左右の外枠のみに罫線を引きます。
'引数range_v省略時は、アクティブなセル範囲を対象とします。
'引数range_v省略時は、アクティブシートになるため引数sheetNameも省略する必要があります。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。
'
'引数例:writeGridEdge xlContinuous, xlThick, "A1:C4", "Sheet1"

Function writeGridEdge(lineStyle_v As Variant, weight_v As Variant, Optional range_v As Variant, Optional sheetName As Variant)


If IsMissing(range_v) Then

    '範囲指定が引数として渡されず、アクティブな領域を対象とする場合
    Selection.BorderAround LineStyle:=lineStyle_v, Weight:=weight_v
        
Else
    
    '範囲指定が引数として渡される場合、sheetNameが渡されているかどうかで分岐する。

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range(range_v).BorderAround LineStyle:=lineStyle_v, Weight:=weight_v
    Else
        Worksheets(sheetName).Range(range_v).BorderAround LineStyle:=lineStyle_v, Weight:=weight_v
    End If

End If
    
End Function


'writeGridInside

'>シートとセル範囲を指定し、セル範囲の内部に罫線を引きます。
'引数range_v省略時は、アクティブなセル範囲を対象とします。
'引数range_v省略時は、アクティブシートになるため引数sheetNameも省略する必要があります。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。
'
'
'引数例:writeGridInside xlContinuous, xlThick, "A1:C4", "Sheet1"

Function writeGridInside(lineStyle_v As Variant, weight_v As Variant, Optional range_v As Variant, Optional sheetName As Variant)

    If IsMissing(range_v) Then

    '範囲指定が引数として渡されず、アクティブな領域を対象とする場合
        With Selection.Borders(xlInsideHorizontal)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        With Selection.Borders(xlInsideVertical)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
    Else
    
    '範囲指定が引数として渡される場合、sheetNameが渡されているかどうかで分岐する。
    
    
        If IsMissing(sheetName) Then
    
        'sheetName省略した場合
        With Worksheets(ActiveSheet.Name).Range(range_v).Borders(xlInsideHorizontal)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        With Worksheets(ActiveSheet.Name).Range(range_v).Borders(xlInsideVertical)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
    
        Else
        
        
        With Worksheets(sheetName).Range(range_v).Borders(xlInsideHorizontal)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        With Worksheets(sheetName).Range(range_v).Borders(xlInsideVertical)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        End If
    
    End If

End Function


'writeGridSheetHoraizontal

'>シートを指定し、シート全体に水平方向の罫線を引きます。
'引数sheetNameのみ省略時はアクティブなシートが選択されます。
'
'
'引数例:writeGridSheetHoraizontal xlContinuous, xlThick, "Sheet1"

Function writeGridSheetHoraizontal(lineStyle_v As Variant, weight_v As Variant, Optional sheetName As Variant)
  
    If IsMissing(sheetName) Then
    
        'sheetName省略した場合
        With Worksheets(ActiveSheet.Name).Cells.Borders(xlInsideHorizontal)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
    Else
        
        With Worksheets(sheetName).Cells.Borders(xlInsideHorizontal)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
    End If
  
End Function


'writeGridUsedRange

'>シートを指定し、値が入力されている範囲に格子状に罫線を引きます。
'sheetName省略時はアクティブなシートが選択されます。

Function writeGridUsedRange(lineStyle_v As Variant, weight_v As Variant, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
    
        With Worksheets(ActiveSheet.Name).UsedRange.Borders
          .LineStyle = lineStyle_v
          .Weight = weight_v
        End With
      
    Else
    
      With Worksheets(sheetName).UsedRange.Borders
          .LineStyle = lineStyle_v
          .Weight = weight_v
      End With
      
    End If

End Function


'writeGridColumns

'>シートと列を指定し、格子状に罫線を引きます。
'sheetName省略時はアクティブなシートが選択されます。
'
'引数例:writeGridColumns xlContinuous, xlThick, "A:I", "hogeSheet"
'離れた列範囲指定の場合: writeGridColumns xlContinuous, xlThick, "A:D, G:H", "hogeSheet"

Function writeGridColumns(lineStyle_v As Variant, weight_v As Variant, columnsName As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Range(columnsName).Select
        writeGrid lineStyle_v, weight_v
    Else
        Worksheets(sheetName).Activate
        Range(columnsName).Select
        writeGrid lineStyle_v, weight_v
    End If

End Function


'writeGridRows

'>シートと行を指定し、格子状に罫線を引きます。
'sheetName省略時はアクティブなシートが選択されます。
'
'引数例:writeGridRows xlContinuous, xlThick, "A:I", "hogeSheet"
'離れた行範囲指定の場合: writeGridRows xlContinuous, xlThick, "1:3, 14:15", "hogeSheet"

Function writeGridRows(lineStyle_v As Variant, weight_v As Variant, rowsName As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Range(rowsName).Select
        writeGrid lineStyle_v, weight_v
    Else
        Worksheets(sheetName).Activate
        Range(rowsName).Select
        writeGrid lineStyle_v, weight_v
    End If

End Function

'writeGridBottom

'>シートとをセル範囲を指定し、下辺に罫線を引きます。
'
'引数range_v省略時は、アクティブなセル範囲を対象とします。
'引数range_v省略時は、アクティブシートになるため引数sheetNameも省略する必要があります。
'sheetName省略時はアクティブなシートが選択されます。
'
'引数例:writeGridBottom xlContinuous, xlThick, "A2:C3", "hogeSheet"


Function writeGridBottom(lineStyle_v As Variant, weight_v As Variant, Optional range_v As Variant, Optional sheetName As Variant)

    If IsMissing(range_v) Then

    '範囲指定が引数として渡されず、アクティブな領域を対象とする場合
        With Selection.Borders(xlEdgeBottom)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        With Selection.Borders(xlEdgeBottom)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
    Else
    
    '範囲指定が引数として渡される場合、sheetNameが渡されているかどうかで分岐する。
    
    
        If IsMissing(sheetName) Then
    
        'sheetName省略した場合
        With Worksheets(ActiveSheet.Name).Range(range_v).Borders(xlEdgeBottom)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        With Worksheets(ActiveSheet.Name).Range(range_v).Borders(xlEdgeBottom)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
    
        Else
        
        
        With Worksheets(sheetName).Range(range_v).Borders(xlEdgeBottom)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        With Worksheets(sheetName).Range(range_v).Borders(xlEdgeBottom)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        End If
    
    End If

End Function


'showGridLine

'枠線を表示状態にします。
Function showGridline()

      ActiveWindow.DisplayGridlines = True

End Function



'hideGridLine

'枠線を非表示状態にします。
Function hideGridline()

      ActiveWindow.DisplayGridlines = False

End Function




'+------------------------------------------------------------------+


