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

'このクラスはクリップボードへの操作メソッドを提供します｡
'
'<依存ライブラリ>
'Microsoft Forms 2.0 Object Library
'
'+------------------------------------------------------------------+


'文字列を受け取りクリップボードに保存します。

Function saveClip(str As String)
  With New MSForms.DataObject
    .SetText str
    .PutInClipboard
  End With
End Function



'クリップボードの中身を取得します。

Function getClip()
   With New MSForms.DataObject
     .GetFromClipboard
    getClip = .GetText
  End With
End Function


'クリップボードの中身を削除します。
'開発中
Private Function clearClip()
    
    Excel.Application.CutCopyMode = False

End Function



