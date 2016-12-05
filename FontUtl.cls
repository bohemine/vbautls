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

'このクラスモジュールはフォント操作用のメソッドを提供します。
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+


'changeFont

'指定したシート全体のフォントを変更します。

'sheetName省略時はアクティブなシートを対象とします。
Function changeFont(fontName As String, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Cells.Font.Name = fontName
    Else
        Worksheets(sheetName).Cells.Font.Name = fontName
    End If

End Function


'resizeFont

'指定したシート全体のフォントを変更します。

'sheetName省略時はアクティブなシートを対象とします。
Function resizeFont(fontSize_v As Integer, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
        Worksheets(ActiveSheet.Name).Cells.Font.Size = fontSize_v
    Else
        Worksheets(sheetName).Cells.Font.Size = fontSize_v
    End If

End Function


'+------------------------------------------------------------------+

