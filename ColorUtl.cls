VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColorUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'+--Overview--------------------------------------------------------+
'Classmodule:ColorUtl
'Version:1.0.0

'このクラスモジュールは色付け用のメソッドを提供します。
'colorIndexへはRGB()を使用して値を渡すことができます。

'+------------------------------------------------------------------+



'+--Public Functions------------------------------------------------+

'paintRow
'指定した行へ色を付けます。
'
'sheetName省略時はアクティブなシートを対象とします。
Function paintRow(rowNumber As String, colorIndex As Long, Optional sheetName As Variant)
  
    Dim rowNumber_forRange As String
    rowNumber_forRange = rowNumber + ":" + rowNumber
  
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range(rowNumber_forRange).Interior.color = colorIndex
    Else
        Worksheets(sheetName).Range(rowNumber_forRange).Interior.color = colorIndex
    End If

End Function



'指定した行の文字へ色を付けます。
'
'sheetName省略時はアクティブなシートを対象とします。

Function paintRowFont(rowNumber As String, colorIndex As Long, Optional sheetName As Variant)
  
    Dim rowNumber_forRange As String
  
    rowNumber_forRange = rowNumber + ":" + rowNumber
  
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range(rowNumber_forRange).Font.color = colorIndex
    Else
        Worksheets(sheetName).Range(rowNumber_forRange).Font.color = colorIndex
    End If

End Function



'先頭行へ色を付けます。
'
'sheetName省略時はアクティブなシートを対象とします。

Function paintHeadRow(colorIndex As Long, Optional sheetName As Variant)
  
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range("1:1").Interior.color = colorIndex
    Else
        Worksheets(sheetName).Range("1:1").Interior.color = colorIndex
    End If
       
End Function



'指定した列を対象に指定した値をもつセルへ色を付けます。
'
'検索対象は値として指定されます。
'sheetName省略時はアクティブなシートを対象とします。

Function paintVerticalCells(columnNumber As String, target As String, colorIndex As Long, Optional sheetName As Variant)

Dim columnNumber_forRange As String, fCell As Range, fAdr As String
columnNumber_forRange = columnNumber + ":" + columnNumber

If IsMissing(sheetName) Then
'sheetName省略時

    Set fCell = Worksheets(ActiveSheet.Name).Range(columnNumber_forRange).Find(what:=target, LookIn:=xlValues)

    If Not fCell Is Nothing Then
        fAdr = fCell.Address
        Do
            fCell.Interior.color = colorIndex
            Set fCell = Worksheets(ActiveSheet.Name).Range(columnNumber_forRange).FindNext(fCell)
            If fCell.Address = fAdr Then Exit Do
        Loop
    End If

Else
'sheetName指定時
    
    Set fCell = Worksheets(sheetName).Range(columnNumber_forRange).Find(what:=target, LookIn:=xlValues)

    If Not fCell Is Nothing Then
        fAdr = fCell.Address
        Do
            fCell.Interior.color = colorIndex
            Set fCell = Worksheets(sheetName).Range(columnNumber_forRange).FindNext(fCell)
            If fCell.Address = fAdr Then Exit Do
        Loop
    End If

End If

End Function



'指定した列に指定した値をもつ行の文字へ着色
'
'検索対象は値として指定されます。
'sheetName省略時はアクティブなシートを対象とします。

Function paintVerticalCellsFont(columnNumber As String, target As String, colorIndex As Long, Optional sheetName As Variant)

Dim columnNumber_forRange As String, fCell As Range, fAdr As String
columnNumber_forRange = columnNumber + ":" + columnNumber

If IsMissing(sheetName) Then
'sheetName省略時

    Set fCell = Worksheets(ActiveSheet.Name).Range(columnNumber_forRange).Find(what:=target, LookIn:=xlValues)

    If Not fCell Is Nothing Then
        fAdr = fCell.Address
        Do
            fCell.Font.color = colorIndex
            Set fCell = Worksheets(ActiveSheet.Name).Range(columnNumber_forRange).FindNext(fCell)
            If fCell.Address = fAdr Then Exit Do
        Loop
    End If

Else
'sheetName指定時
    
    Set fCell = Worksheets(sheetName).Range(columnNumber_forRange).Find(what:=target, LookIn:=xlValues)

    If Not fCell Is Nothing Then
      fAdr = fCell.Address
      Do
            fCell.Font.color = colorIndex
            Set fCell = Worksheets(sheetName).Range(columnNumber_forRange).FindNext(fCell)
            If fCell.Address = fAdr Then Exit Do
      Loop
    End If

End If

End Function



'指定した列に指定した値をもつ行へ色を付けます。
'
'sheetName省略時はアクティブなシートを対象とします。
'内部的にpaintRow使用

Function paintRowsHave(columnNumber As String, target As String, colorIndex As Long, Optional sheetName As Variant)

Dim columnNumber_forRange As String, fCell As Range, fAdr As String
columnNumber_forRange = columnNumber + ":" + columnNumber


If IsMissing(sheetName) Then
'sheetName省略時
      
    Set fCell = Worksheets(ActiveSheet.Name).Range(columnNumber_forRange).Find(what:=target, LookIn:=xlValues)

    If Not fCell Is Nothing Then
        fAdr = fCell.Address
        
        Do
        paintRow fCell.row, colorIndex, ActiveSheet.Name
        Set fCell = Worksheets(ActiveSheet.Name).Range(columnNumber_forRange).FindNext(fCell)
        If fCell.Address = fAdr Then Exit Do
        Loop
    End If
          
Else
'sheetName指定時

    Set fCell = Worksheets(sheetName).Range(columnNumber_forRange).Find(what:=target, LookIn:=xlValues)

    If Not fCell Is Nothing Then
        fAdr = fCell.Address
        Do
        paintRow fCell.row, colorIndex, sheetName
        Set fCell = Worksheets(sheetName).Range(columnNumber_forRange).FindNext(fCell)
        If fCell.Address = fAdr Then Exit Do
        Loop
    End If

End If

End Function




'指定した列に指定した値をもつ行の文字へ色を付けます。
'
'sheetName省略時はアクティブなシートを対象とします。
'内部的にpaintRowFont使用

Function paintRowsFontHave(columnNumber As String, target As String, colorIndex As Long, Optional sheetName As Variant)

Dim columnNumber_forRange As String, fCell As Range, fAdr As String
columnNumber_forRange = columnNumber + ":" + columnNumber


If IsMissing(sheetName) Then
'sheetName省略時
      
    Set fCell = Worksheets(ActiveSheet.Name).Range(columnNumber_forRange).Find(what:=target, LookIn:=xlValues)

    If Not fCell Is Nothing Then
        fAdr = fCell.Address
        
        Do
        paintRowFont fCell.row, colorIndex, ActiveSheet.Name
        Set fCell = Worksheets(ActiveSheet.Name).Range(columnNumber_forRange).FindNext(fCell)
        If fCell.Address = fAdr Then Exit Do
        Loop
    End If
          
Else
'sheetName指定時

    Set fCell = Worksheets(sheetName).Range(columnNumber_forRange).Find(what:=target, LookIn:=xlValues)

    If Not fCell Is Nothing Then
        fAdr = fCell.Address
        Do
        paintRowFont fCell.row, colorIndex, sheetName
        Set fCell = Worksheets(sheetName).Range(columnNumber_forRange).FindNext(fCell)
        If fCell.Address = fAdr Then Exit Do
        Loop
    End If

End If

End Function



'指定したシートタブを色付けします。
'
'sheetName省略時はアクティブなシートを対象とします。

Function colorSheetTab(colorIndex As Long, Optional sheetName As Variant)

If IsMissing(sheetName) Then
  Worksheets(ActiveSheet.Name).Tab.color = colorIndex
Else
  Worksheets(sheetName).Tab.color = colorIndex
End If
     
End Function



'指定したシートの文字を色付けします。
'
'sheetName省略時はアクティブなシートを対象とします。

Function colorSheetFont(colorIndex As Long, Optional sheetName As Variant)

If IsMissing(sheetName) Then
  Worksheets(ActiveSheet.Name).Cells.Font.color = colorIndex
Else
  Worksheets(sheetName).Cells.Font.color = colorIndex
End If
     
End Function



'指定したシートのタブの色を消します。
'
'sheetName省略時はアクティブなシートを対象とします。

Function clearSheetTabColor(Optional sheetName As Variant)

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Tab.colorIndex = xlColorIndexNone
Else
    Worksheets(sheetName).Tab.colorIndex = xlColorIndexNone
End If
     
End Function



'指定したシートのセルへ色を付けます。
'
'sheetName省略時はアクティブなシートを対象とします。

Function colorSheet(colorIndex As Long, Optional sheetName As Variant)

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Cells.Interior.color = colorIndex
Else
    Worksheets(sheetName).Cells.Interior.color = colorIndex
End If

End Function



'指定したシートの色を白にします。
'
'sheetName省略時はアクティブなシートを対象とします。

Function colorWhiteSheet(Optional sheetName As Variant)

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Cells.Interior.color = RGB(255, 255, 255)
Else
    Worksheets(sheetName).Cells.Interior.color = RGB(255, 255, 255)
End If
End Function



'指定したシートの色を消します。
'
'sheetName省略時はアクティブなシートを対象とします。

Function clearSheetColor(Optional sheetName As Variant)

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Cells.Interior.colorIndex = xlNone
Else
    Worksheets(sheetName).Cells.Interior.colorIndex = xlNone
End If
End Function



'指定したシートの行をストライプ状に色付けします。
'UsedRangeを対象とします。
'先頭行は対象から除きます。
'sheetName省略時はアクティブなシートを対象とします。

Function colorStripeRows(colorIndex_even As Long, colorIndex_odd As Long, Optional sheetName As Variant)

If IsMissing(sheetName) Then
    
    Dim i1 As Integer
    For i1 = 2 To Worksheets(ActiveSheet.Name).UsedRange.Rows.Count
        If isEven(i1) Then
            paintRow CStr(i1), colorIndex_even, ActiveSheet.Name
        Else
            paintRow CStr(i1), colorIndex_odd, ActiveSheet.Name
        End If
    Next

Else

    Dim i2 As Integer
    For i2 = 2 To Worksheets(sheetName).UsedRange.Rows.Count
        If isEven(i2) Then
            paintRow CStr(i2), colorIndex_even, sheetName
        Else
            paintRow CStr(i2), colorIndex_odd, sheetName
        End If
    Next
End If

  

End Function

'+------------------------------------------------------------------+





'+--Private Functions------------------------------------------------+



'isEven

'偶数判定
Private Function isEven(i As Integer) As Boolean
  If i Mod 2 = 0 Then
    isEven = True
  Else
    isEven = False
  End If
End Function




'+------------------------------------------------------------------+
















'+------------------------------------------------------------------+


