VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutomationUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'+--Overview--------------------------------------------------------+
'Classmodule:AutomationUtl
'Version:1.0.0

'+------------------------------------------------------------------+


'�w�肵���A�v���P�[�V�������A�N�e�B�u�ɂ��܂��B
'�E�B���h�E�̃^�C�g�����w�肷��B
Function activeApp(appName As String)

  AppActivate appName, True

End Function


'�w�肵���L�[�R�[�h�𑗂�܂��B
'
'example
'A��10�񑗂�
'sendKeys "{A 10}"

Function sendKey(keyCode As String)

  sendKeys keyCode, True

End Function

'TAB�L�[�̃L�[�R�[�h�𑗂�܂��B

Function sendTab()

  sendKeys "{TAB}", True

End Function


'ENTER�L�[�̃L�[�R�[�h�𑗂�܂��

Function sendEnter()

  sendKeys "{ENTER}", True

End Function


'�w�肵�����Ԃ̊ԃX���[�v���܂��

Function sleep(second As String)

  Application.Wait Now + TimeValue("0:00:" & second)

End Function














