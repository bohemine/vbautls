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
'Version:0.0.1

'このクラスモジュールはワークブック保存用のメソッドを提供します。

'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+



'開発中
'
'名前を付けてワークブックを保存します。
'fileNameは拡張子を除いた部分を渡して下さい。
Private Function saveAsFile(fileName As String, filefilter_v As String)

Application.GetSaveAsFilename InitialFileName:=fileName, FileFilter:=filefilter_v

' filePath_v As String, fileFormat_v As Variant

End Function




'開発中
'
'ワークブックを保存します。
Private Function saveBook(filePath_v As String, fileFormat_v As Variant)

Dim book As Workbook
Set book = ActiveWorkbook
book.SaveAs fileName:=filePath_v, fileFormat:=fileFormat_v


End Function





























'+------------------------------------------------------------------+



