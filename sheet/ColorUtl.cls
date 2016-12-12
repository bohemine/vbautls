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

'���̃N���X���W���[���͐F�t���p�̃��\�b�h��񋟂��܂��B
'colorIndex�ւ�RGB()���g�p���Ēl��n�����Ƃ��ł��܂��B

'+------------------------------------------------------------------+



'+--Public Functions------------------------------------------------+

'paintRow
'�w�肵���s�֐F��t���܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B
Function paintRow(rowNumber As String, colorIndex As Long, Optional sheetName As Variant)
  
    Dim rowNumber_forRange As String
    rowNumber_forRange = rowNumber + ":" + rowNumber
  
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range(rowNumber_forRange).Interior.color = colorIndex
    Else
        Worksheets(sheetName).Range(rowNumber_forRange).Interior.color = colorIndex
    End If

End Function



'�w�肵���s�̕����֐F��t���܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function paintRowFont(rowNumber As String, colorIndex As Long, Optional sheetName As Variant)
  
    Dim rowNumber_forRange As String
  
    rowNumber_forRange = rowNumber + ":" + rowNumber
  
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range(rowNumber_forRange).Font.color = colorIndex
    Else
        Worksheets(sheetName).Range(rowNumber_forRange).Font.color = colorIndex
    End If

End Function



'�擪�s�֐F��t���܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function paintHeadRow(colorIndex As Long, Optional sheetName As Variant)
  
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range("1:1").Interior.color = colorIndex
    Else
        Worksheets(sheetName).Range("1:1").Interior.color = colorIndex
    End If
       
End Function



'�w�肵�����ΏۂɎw�肵���l�����Z���֐F��t���܂��B
'
'�����Ώۂ͒l�Ƃ��Ďw�肳��܂��B
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function paintVerticalCells(columnNumber As String, target As String, colorIndex As Long, Optional sheetName As Variant)

Dim columnNumber_forRange As String, fCell As Range, fAdr As String
columnNumber_forRange = columnNumber + ":" + columnNumber

If IsMissing(sheetName) Then
'sheetName�ȗ���

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
'sheetName�w�莞
    
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



'�w�肵����Ɏw�肵���l�����s�̕����֒��F
'
'�����Ώۂ͒l�Ƃ��Ďw�肳��܂��B
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function paintVerticalCellsFont(columnNumber As String, target As String, colorIndex As Long, Optional sheetName As Variant)

Dim columnNumber_forRange As String, fCell As Range, fAdr As String
columnNumber_forRange = columnNumber + ":" + columnNumber

If IsMissing(sheetName) Then
'sheetName�ȗ���

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
'sheetName�w�莞
    
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



'�w�肵����Ɏw�肵���l�����s�֐F��t���܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B
'�����I��paintRow�g�p

Function paintRowsHave(columnNumber As String, target As String, colorIndex As Long, Optional sheetName As Variant)

Dim columnNumber_forRange As String, fCell As Range, fAdr As String
columnNumber_forRange = columnNumber + ":" + columnNumber


If IsMissing(sheetName) Then
'sheetName�ȗ���
      
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
'sheetName�w�莞

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




'�w�肵����Ɏw�肵���l�����s�̕����֐F��t���܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B
'�����I��paintRowFont�g�p

Function paintRowsFontHave(columnNumber As String, target As String, colorIndex As Long, Optional sheetName As Variant)

Dim columnNumber_forRange As String, fCell As Range, fAdr As String
columnNumber_forRange = columnNumber + ":" + columnNumber


If IsMissing(sheetName) Then
'sheetName�ȗ���
      
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
'sheetName�w�莞

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



'�w�肵���V�[�g�^�u��F�t�����܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function colorSheetTab(colorIndex As Long, Optional sheetName As Variant)

If IsMissing(sheetName) Then
  Worksheets(ActiveSheet.Name).Tab.color = colorIndex
Else
  Worksheets(sheetName).Tab.color = colorIndex
End If
     
End Function



'�w�肵���V�[�g�̕�����F�t�����܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function colorSheetFont(colorIndex As Long, Optional sheetName As Variant)

If IsMissing(sheetName) Then
  Worksheets(ActiveSheet.Name).Cells.Font.color = colorIndex
Else
  Worksheets(sheetName).Cells.Font.color = colorIndex
End If
     
End Function



'�w�肵���V�[�g�̃^�u�̐F�������܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function clearSheetTabColor(Optional sheetName As Variant)

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Tab.colorIndex = xlColorIndexNone
Else
    Worksheets(sheetName).Tab.colorIndex = xlColorIndexNone
End If
     
End Function



'�w�肵���V�[�g�̃Z���֐F��t���܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function colorSheet(colorIndex As Long, Optional sheetName As Variant)

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Cells.Interior.color = colorIndex
Else
    Worksheets(sheetName).Cells.Interior.color = colorIndex
End If

End Function



'�w�肵���V�[�g�̐F�𔒂ɂ��܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function colorWhiteSheet(Optional sheetName As Variant)

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Cells.Interior.color = RGB(255, 255, 255)
Else
    Worksheets(sheetName).Cells.Interior.color = RGB(255, 255, 255)
End If
End Function



'�w�肵���V�[�g�̐F�������܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function clearSheetColor(Optional sheetName As Variant)

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Cells.Interior.colorIndex = xlNone
Else
    Worksheets(sheetName).Cells.Interior.colorIndex = xlNone
End If
End Function



'�w�肵���V�[�g�̍s���X�g���C�v��ɐF�t�����܂��B
'UsedRange��ΏۂƂ��܂��B
'�擪�s�͑Ώۂ��珜���܂��B
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

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

'��������
Private Function isEven(i As Integer) As Boolean
  If i Mod 2 = 0 Then
    isEven = True
  Else
    isEven = False
  End If
End Function




'+------------------------------------------------------------------+
















'+------------------------------------------------------------------+

