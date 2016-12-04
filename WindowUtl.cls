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

'���̃N���X���W���[���̓E�B���h�E����p�̃��\�b�h��񋟂��܂��B
'
'+------------------------------------------------------------------+


'+--Public Functions------------------------------------------------+

'freeze

'>�V�[�g���ƃZ�����w�肵�E�B���h�E�g���Œ肵�܂��B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B


Function freeze(cell As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Range(cell).Select
        ActiveWindow.FreezePanes = True
    Else
        Worksheets(sheetName).Range(cell).Select
        ActiveWindow.FreezePanes = True
    End If

End Function

'freezeHeadRow

'>�V�[�g�����w�肵�擪�s���Œ肵�܂��B
'���łɃE�B���h�E�g���Œ肳��Ă���ꍇ�͌Œ肵�����Ȃ��B
'�E�B���h�E�g���Z�b�g�������ꍇ�͂P�x�N���A���ĉ������B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

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

'clearFreezePanes

'>�V�[�g�����w�肵�E�B���h�E�g���������܂��B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function clearFreezePanes(Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        ActiveWindow.FreezePanes = False
    Else
        Worksheets(sheetName).Activate
        ActiveWindow.FreezePanes = False
    End If
  
End Function

'autoFitColumns

'>�V�[�g���Ɨ���w�肵�񕝂��œK�����܂��B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

'������:A�񂩂�I��܂Ł@autoFitColumns "A:I", "Sheet1" �A1��̂݁@autoFitColumns "C:C"
Function autoFitColumns(columnsName As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Columns(columnsName).autoFit
    Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
        Worksheets(sheetName).Columns(columnsName).autoFit
    End If
 
End Function

'autoFitRows

'>�V�[�g���ƍs���w�肵�s�����œK�����܂��B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function autoFitRows(rowsName As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Rows(rowsName).autoFit
    Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
        Worksheets(sheetName).Rows(rowsName).autoFit
    End If
 
End Function


'setColumnWidth

'>�V�[�g���Ɨ���w�肵�񕝂�ύX���܂��B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function setColumnsWidth(columnsName As String, columsWidth As Integer, Optional sheetName As Variant)
    
    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Columns(columnsName).ColumnWidth = columsWidth
    Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
        Worksheets(sheetName).Columns(columnsName).ColumnWidth = columsWidth
    End If
    
End Function


'setRowsHeight

'>�V�[�g���ƍs���w�肵�s����ύX���܂��B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

'������:�s1����s3�܂Ł@setRowsHeight "1:3", "Sheet1" �A�s1�̂݁@setRowsHeight "1:1"
Function setRowsHeight(rowsName As String, rowsHeight As Integer, Optional sheetName As Variant)
    
    If IsMissing(sheetName) Then
         Worksheets(ActiveSheet.Name).Rows(rowsName).rowHeight = rowsHeight
    Else
         Worksheets(sheetName).Rows(rowsName).rowHeight = rowsHeight
    End If
    
End Function




'+------------------------------------------------------------------+




