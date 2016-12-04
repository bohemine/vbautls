VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Class module:SheetUtl
'Version:1.0.0

'このクラスモジュールはシート操作用のメソッドを提供します。
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+


'addSheet

'>シートを指定したシートの右に追加します。

'全ての引数を省略した場合、自動的に設定されたシート名のシートが末尾に追加されます。
'引数newSheetNameのみを渡した場合、指定したシート名のシートが末尾に追加されます。
'引数newSheetNameとafterSheetNameを渡した場合、指定したシート名のシートがafterSheetNameの後に追加されます。
'
'afterSheetNameを渡す場合、newSheetNameを省略できないため、afterSheetNameのみを渡してシート名は自動で設定させたい場合、
'afterSheetName:= "Sheet2" のように引数名を指定して値を渡して下さい。
'
'newSheetNameとして既存のシート名を指定した場合エラーが生じます。その場合自動的に設定されたシート名で末尾に追加が行われます。


Function addSheet(Optional newSheetName As Variant, Optional afterSheetName As Variant)

If IsMissing(newSheetName) Then

    If IsMissing(afterSheetName) Then
        Worksheets.Add after:=Worksheets(Worksheets.Count)
    Else
    'newSheetNameを渡さず、afterSheetNameを渡すには引数名を指定して値を渡す。
        Worksheets.Add after:=Worksheets(Worksheets(afterSheetName).index)
    End If

Else
    If IsMissing(afterSheetName) Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = newSheetName
    Else

        Worksheets.Add(after:=Worksheets(Worksheets(afterSheetName).index)).Name = newSheetName
    End If
  
End If
     
End Function

'addSheetBefore

'>シートを指定したシートの左に追加します。

'全ての引数を省略した場合、自動的に設定されたシート名のシートが一番左に追加されます。
'引数newSheetNameのみを渡した場合、指定したシート名のシートが一番左に追加されます。
'引数newSheetNameとbeforeSheetNameを渡した場合、指定したシート名のシートがbeforeSheetNameの前に追加されます。
'
'beforeSheetNameを渡す場合、newSheetNameを省略できないため、beforeSheetNameのみを渡してシート名は自動で設定させたい場合、
'beforeSheetName:= "Sheet2" のように引数名を指定して値を渡して下さい。
'
'newSheetNameとして既存のシート名を指定した場合エラーが生じます。その場合自動的に設定されたシート名で一番左に追加が行われます。


Function addSheetBefore(Optional newSheetName As Variant, Optional beforeSheetName As Variant)

If IsMissing(newSheetName) Then

    If IsMissing(beforeSheetName) Then
        Worksheets.Add before:=Worksheets(1)
    Else
    'newSheetNameを渡さず、beforeSheetNameを渡すには引数名を指定して値を渡す。
        Worksheets.Add before:=Worksheets(Worksheets(beforeSheetName).index)
    End If

Else
    If IsMissing(beforeSheetName) Then
        Worksheets.Add(before:=Worksheets(1)).Name = newSheetName
    Else

        Worksheets.Add(before:=Worksheets(Worksheets(beforeSheetName).index)).Name = newSheetName
    End If
  
End If
     
End Function


'renameSheet

'>指定したしシートの名前を変更します。

'引数oldSheetNameは省略不可です。
'既に存在するシート名を渡した場合エラーが生じます。

Function renameSheet(oldSheetName As String, newSheetName As String)

      Worksheets(oldSheetName).Name = newSheetName
     
End Function

'deleteSheet

'>指定したシートを削除します。

'showAlertにFalseを渡すと削除時のアラートを表示しないようになります。
'showAleart省略もしくはTrueを渡すとアラート有効のまま実行します。
'存在しないシートを指定するとエラーとなります。
'注意! アラート有効にしてもアラートがならない場合があります。再現性のないバグ?

Function deleteSheet(sheetName As String, Optional showAlert As Variant)

 Application.DisplayAlerts = True

If IsMissing(showAlert) Then
 
 Application.DisplayAlerts = True
 Worksheets(sheetName).Delete
 Application.DisplayAlerts = True
   
Else

 Application.DisplayAlerts = showAlert
 Worksheets(sheetName).Delete
 Application.DisplayAlerts = True
   
End If
  Application.DisplayAlerts = True

End Function

'copySheet
'>指定したシートを末尾にコピーします。
'
'引数newSheetName省略時はシート名は自動的に設定されます。
Function copySheet(sheetName As String, Optional newSheetName As Variant)
  
Worksheets(sheetName).Copy after:=Worksheets(Worksheets.Count)
  
If IsMissing(newSheetName) Then
    
Else
    Worksheets(Worksheets.Count).Name = newSheetName
End If
  
End Function

'moveSheet
'>指定したシートを指定したシートの左へ移動します。
'
'引数afterSheetName省略時は末尾に移動します。
Function moveSheet(sheetName As String, Optional afterSheetName As Variant)
  
If IsMissing(afterSheetName) Then
     Worksheets(sheetName).Move after:=Worksheets(Worksheets.Count)
Else
     Worksheets(sheetName).Move after:=Worksheets(Worksheets(afterSheetName).index)
End If
  
End Function

'moveSheetBefore
'>指定したシートを指定したシートの左へ移動します。
'
'引数beforeSheetName省略時は一番左へ移動します。
Function moveSheetBefore(sheetName As String, Optional beforeSheetName As Variant)
   
If IsMissing(beforeSheetName) Then
     Worksheets(sheetName).Move before:=Worksheets(1)
Else
     Worksheets(sheetName).Move before:=Worksheets(Worksheets(beforeSheetName).index)
End If
  
End Function




'+------------------------------------------------------------------+


