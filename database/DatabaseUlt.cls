VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseUlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



'+--Overview--------------------------------------------------------+
'Classmodule:DatabaseUtl
'Version:1.0.0

'Dependencies:
'Microsoft ActiveX Data Objects 6.1 Library
'Microsoft.ACE.OLEDB.12.0

'+------------------------------------------------------------------+





'ACCESS�̃t�@�C���p�X�ƃv���o�C�_�[���w�肵ADODB.Connection�I�u�W�F�N�g��Ԃ��܂��B
'accdb�`����ACCESS�t�@�C���ɐڑ�����ꍇ�A�v���o�C�_�Ƃ���[Microsoft.ACE.OLEDB.12.0]���w�肷��B

Function connect(dbPath As String, provider As String) As ADODB.Connection

  Set connect = CreateObject("ADODB.Connection")
  connect.Open "Provider=" & provider & ";" & "Data Source=" & dbPath & ";"

End Function



'ADODB.Connection�I�u�W�F�N�g��n���ؒf���܂��B

Function disConnect(ByRef con As ADODB.Connection)

 con.Close
 Set con = Nothing
 
End Function

