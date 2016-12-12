VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FontUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Class module:FontUtl
'Version:1.0.0

'���̃N���X���W���[���̓t�H���g����p�̃��\�b�h��񋟂��܂��B
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+


'�w�肵���V�[�g�S�̂̃t�H���g��ύX���܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function changeFont(fontName As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Cells.Font.Name = fontName
    Else
        Worksheets(sheetName).Cells.Font.Name = fontName
    End If

End Function



'�w�肵���V�[�g�S�̂̃t�H���g�̃T�C�Y��ύX���܂��B
'
'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B

Function resizeFont(fontSize_v As Integer, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Cells.Font.Size = fontSize_v
    Else
        Worksheets(sheetName).Cells.Font.Size = fontSize_v
    End If

End Function


'+------------------------------------------------------------------+
