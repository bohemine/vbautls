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





'ACCESSのファイルパスとプロバイダーを指定しADODB.Connectionオブジェクトを返します。
'accdb形式のACCESSファイルに接続する場合、プロバイダとして[Microsoft.ACE.OLEDB.12.0]を指定する。

Function connect(dbPath As String, provider As String) As ADODB.Connection

  Set connect = CreateObject("ADODB.Connection")
  connect.Open "Provider=" & provider & ";" & "Data Source=" & dbPath & ";"

End Function



'ADODB.Connectionオブジェクトを渡し切断します。

Function disConnect(ByRef con As ADODB.Connection)

 con.Close
 Set con = Nothing
 
End Function


