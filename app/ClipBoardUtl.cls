VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClipBoardUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'+--Overview--------------------------------------------------------+
'Classmodule:ClipboardUtl
'Version:1.0.0

'���̃N���X�̓N���b�v�{�[�h�ւ̑��상�\�b�h��񋟂��܂��
'
'<�ˑ����C�u����>
'Microsoft Forms 2.0 Object Library
'
'+------------------------------------------------------------------+


'��������󂯎��N���b�v�{�[�h�ɕۑ����܂��B

Function saveClip(str As String)
  With New MSForms.DataObject
    .SetText str
    .PutInClipboard
  End With
End Function



'�N���b�v�{�[�h�̒��g���擾���܂��B

Function getClip()
   With New MSForms.DataObject
     .GetFromClipboard
    getClip = .GetText
  End With
End Function


'�N���b�v�{�[�h�̒��g���폜���܂��B
'�J����
Private Function clearClip()
    
    Excel.Application.CutCopyMode = False

End Function


