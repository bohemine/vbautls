VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SaveUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Classmodule:SaveUtl
'Version:1.0.0

'���̃N���X���W���[���̓��[�N�u�b�N�ۑ��p�̃��\�b�h��񋟂��܂��B

'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+



'�J����
'saveBook

'���O��t���ă��[�N�u�b�N��ۑ����܂��B
'fileName�͊g���q��������������n���ĉ������B
Private Function saveAsFile(fileName As String, filefilter_v As String)

Application.GetSaveAsFilename InitialFileName:=fileName, FileFilter:=filefilter_v

' filePath_v As String, fileFormat_v As Variant

End Function




'�J����
'saveBook

'���[�N�u�b�N��ۑ����܂��B
Private Function saveBook(filePath_v As String, fileFormat_v As Variant)

Dim book As Workbook
Set book = ActiveWorkbook
book.SaveAs fileName:=filePath_v, fileFormat:=fileFormat_v


End Function





























'+------------------------------------------------------------------+


