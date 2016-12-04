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

'���̃N���X���W���[���͌r������p�̃��\�b�h��񋟂��܂��B
'
'+------------------------------------------------------------------+



'+--Public Functions------------------------------------------------+

'writeGrid

'>�V�[�g�Ɣ͈͂��w�肵�A�i�q��Ɍr���������܂��B
'����range_v�ȗ����́A�A�N�e�B�u�ȃZ���͈͂�ΏۂƂ��܂��B
'���̏ꍇ�V�[�g���A�N�e�B�u�V�[�g�ɂȂ邽�߈���sheetName���ȗ�����K�v������܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'

Function writeGrid(lineStyle_v As Variant, weight_v As Variant, Optional range_v As Variant, Optional sheetName As Variant)

If IsMissing(range_v) Then

    '�͈͎w�肪�����Ƃ��ēn���ꂸ�A�A�N�e�B�u�ȗ̈��ΏۂƂ���ꍇ
    With Selection.Borders
        .LineStyle = lineStyle_v
        .Weight = weight_v
    End With
Else
    
    '�͈͎w�肪�����Ƃ��ēn�����ꍇ�AsheetName���n����Ă��邩�ǂ����ŕ��򂷂�B

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

'>�V�[�g�Ɣ͈͂��w�肵�A�͈͂̏㉺���E�Ɍr���������܂��B
'����range_v�ȗ����́A�A�N�e�B�u�ȃZ���͈͂�ΏۂƂ��܂��B
'���̏ꍇ�V�[�g���A�N�e�B�u�V�[�g�ɂȂ邽�߈���sheetName���ȗ�����K�v������܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
'������:writeGridEdge xlContinuous, xlThick, "A1:C4", "Sheet1"

Function writeGridEdge(lineStyle_v As Variant, weight_v As Variant, Optional range_v As Variant, Optional sheetName As Variant)


If IsMissing(range_v) Then

    '�͈͎w�肪�����Ƃ��ēn���ꂸ�A�A�N�e�B�u�ȗ̈��ΏۂƂ���ꍇ
    Selection.BorderAround LineStyle:=lineStyle_v, Weight:=weight_v
        
Else
    
    '�͈͎w�肪�����Ƃ��ēn�����ꍇ�AsheetName���n����Ă��邩�ǂ����ŕ��򂷂�B

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range(range_v).BorderAround LineStyle:=lineStyle_v, Weight:=weight_v
    Else
        Worksheets(sheetName).Range(range_v).BorderAround LineStyle:=lineStyle_v, Weight:=weight_v
    End If

End If
    
End Function


'writeGridInside

'>�V�[�g�Ɣ͈͂��w�肵�A�͈͂̓����Ɍr���������܂��B
'����range_v�ȗ����́A�A�N�e�B�u�ȃZ���͈͂�ΏۂƂ��܂��B
'���̏ꍇ�V�[�g���A�N�e�B�u�V�[�g�ɂȂ邽�߈���sheetName���ȗ�����K�v������܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
'
'������:writeGridInside xlContinuous, xlThick, "A1:C4", "Sheet1"

Function writeGridInside(lineStyle_v As Variant, weight_v As Variant, Optional range_v As Variant, Optional sheetName As Variant)

    If IsMissing(range_v) Then

    '�͈͎w�肪�����Ƃ��ēn���ꂸ�A�A�N�e�B�u�ȗ̈��ΏۂƂ���ꍇ
        With Selection.Borders(xlInsideHorizontal)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
        With Selection.Borders(xlInsideVertical)
        .LineStyle = lineStyle_v
        .Weight = weight_v
        End With
    
    Else
    
    '�͈͎w�肪�����Ƃ��ēn�����ꍇ�AsheetName���n����Ă��邩�ǂ����ŕ��򂷂�B
    
    
        If IsMissing(sheetName) Then
    
        'sheetName�ȗ������ꍇ
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

'>�V�[�g�Ɣ͈͂��w�肵�A���������̌r���������܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
'
'������:writeGridSheetHoraizontal xlContinuous, xlThick, "Sheet1"

Function writeGridSheetHoraizontal(lineStyle_v As Variant, weight_v As Variant, Optional sheetName As Variant)
  
    If IsMissing(sheetName) Then
    
        'sheetName�ȗ������ꍇ
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

'>�V�[�g���w�肵�A�l�����͂���Ă���͈͂Ɋi�q��Ɍr���������܂��B
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
'
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

'>�V�[�g�Ɨ���w�肵�A�i�q��Ɍr���������܂��B
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
'������:writeGridColumns xlContinuous, xlThick, "A:I", "Sheet1" ���ꂽ��͈͎w�� "A:D, G:H"
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

'>�V�[�g�ƍs���w�肵�A�i�q��Ɍr���������܂��B
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
'������:writeGridRows xlContinuous, xlThick, "A:I", "Sheet1" ���ꂽ��͈͎w�� "1:3, 14:15"
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


'showGridLine

'�g����\����Ԃɂ��܂��B
Function showGridline()

      ActiveWindow.DisplayGridlines = True

End Function



'hideGridLine

'�g����\����Ԃɂ��܂��B
Function hideGridline()

      ActiveWindow.DisplayGridlines = False

End Function




'+------------------------------------------------------------------+

