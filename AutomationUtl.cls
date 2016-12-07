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


'指定したアプリケーションをアクティブにします。
'ウィンドウのタイトルを指定する。
Function activeApp(appName As String)

  AppActivate appName, True

End Function


'指定したキーコードを送ります。
'
'example
'Aを10回送る
'sendKeys "{A 10}"

Function sendKey(keyCode As String)

  sendKeys keyCode, True

End Function

'TABキーのキーコードを送ります。

Function sendTab()

  sendKeys "{TAB}", True

End Function


'ENTERキーのキーコードを送ります｡

Function sendEnter()

  sendKeys "{ENTER}", True

End Function


'指定した時間の間スリープします｡

Function sleep(second As String)

  Application.Wait Now + TimeValue("0:00:" & second)

End Function















