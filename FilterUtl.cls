VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'+--Overview--------------------------------------------------------+
'Class module:FilterUtl
'Version:1.0.0

'このクラスモジュールはフィルタ操作用のメソッドを提供します。
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+


'setFilter

'>シートと行番号を指定し、フィルタを表示します。
'第2引数sheetName省略時はアクティブなシートが選択されます。
'
'指定したシート内のいずれかの箇所でフィルタが表示済みであった場合は動作しません。
'新たに表示し直す場合、フィルタを非表示にしてから実行して下さい。
'
'引数例:setFilter "1","Sheet1"  、  setFilter "1"

Function setFilter(rowNumber As String, Optional sheetName As Variant)

    Dim RowNumberStr As String
    RowNumberStr = rowNumber + ":" + rowNumber


    '第2引数sheetNameの明示、省略による条件分岐
    If IsMissing(sheetName) Then
        ''ワークシート指定なしの場合アクティブなシートを選択のため何もしない。
        Else
        ''ワークシート指定ありの場合
            Worksheets(sheetName).Activate
    End If

    
    If ActiveSheet.AutoFilterMode = False Then
        Range(RowNumberStr).Select
        Selection.AutoFilter
    End If
 
End Function


'setHeaderFilter

'>指定したシートの先頭行でフィルタを表示します。
'引数sheetName省略時はアクティブなシートが選択されます。
'
'指定したシート内のいずれかの箇所でフィルタが表示済みであった場合は、表示し直すことはありません。
'新たに表示し直す場合、フィルタを非表示にしてから実行して下さい。

Function setHeadFilter(Optional sheetName As Variant)

    If IsMissing(sheetName) Then
            setFilter "1", ActiveSheet.Name
        Else
        ''ワークシート指定ありの場合
            setFilter "1", sheetName
    End If
    
End Function


'clearFilter

'>指定したシート内のフィルタを非表示にします。
'引数sheetName省略時はアクティブなシートが選択されます。
'シート内全域に対して有効。

Function clearFilter(Optional sheetName As Variant)

  If IsMissing(sheetName) Then
        ''ワークシート指定なしの場合アクティブなシートを選択のため何もしない。
        Else
        ''ワークシート指定ありの場合
              Worksheets(sheetName).Activate
    End If
  
    If ActiveSheet.AutoFilterMode = True Then
        Range("1:1").Select
        Selection.AutoFilter
    End If
    
End Function


'filter

'シートと列番号と条件１つを指定しフィルタを行います。
'第3引数sheetName省略時はアクティブなシートが選択されます。
'フィルタは先頭行で実行されます。

'引数例:filter 3, ">2015/12/16", filterStr 1, ">20"

Function filter(columnIndex As Integer, status As Variant, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
            Worksheets(ActiveSheet.Name).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status
        Else
        ''ワークシート指定ありの場合
            Worksheets(sheetName).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status
    End If
  
End Function


'filterOr

'シートと指定した列へ２つの条件のフィルタを行います。
'第4引数sheetName省略時はアクティブなシートが選択されます。
'フィルタは先頭行で実行されます。
'２つの条件はOr演算となります。

Function filterOr(columnIndex As Integer, status1 As String, status2 As String, Optional sheetName As Variant)

   If IsMissing(sheetName) Then
           
            Worksheets(ActiveSheet.Name).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status1, Operator:=xlOr, Criteria2:=status2

        Else
        ''ワークシート指定ありの場合
            Worksheets(sheetName).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status1, Operator:=xlOr, Criteria2:=status2
    End If
    
End Function

'filterAnd

'シートと指定した列へ２つの条件のフィルタを行います。
'第4引数sheetName省略時はアクティブなシートが選択されます。
'フィルタは先頭行で実行されます。
'２つの条件はAnd演算となります。

'引数例:範囲で絞り込む filterAnd 2, ">=1", "<=10", "Sheet1"

Function filterAnd(columnIndex As Integer, status1 As String, status2 As String, Optional sheetName As Variant)

   If IsMissing(sheetName) Then
           
            Worksheets(ActiveSheet.Name).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status1, Operator:=xlAnd, Criteria2:=status2

        Else
        ''ワークシート指定ありの場合
            Worksheets(sheetName).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status1, Operator:=xlAnd, Criteria2:=status2
    End If
    
End Function


'filterSpan

'シートと指定した列へ２つの条件の範囲でフィルタを行います。
'第4引数sheetName省略時はアクティブなシートが選択されます。
'フィルタは先頭行で実行されます。
'２つの条件はAnd演算となります。

'条件自体の値も検索対象になる。
'引数例:範囲で絞り込む filterSpan 2, "1", "100", "Sheet1"

Function filterSpan(columnIndex As Integer, fromStatus As String, toStatus As String, Optional sheetName As Variant)

   If IsMissing(sheetName) Then
           
            Worksheets(ActiveSheet.Name).Range("A1").AutoFilter field:=columnIndex, Criteria1:=">=" & fromStatus, Operator:=xlAnd, Criteria2:="<=" & toStatus

        Else
        ''ワークシート指定ありの場合
            Worksheets(sheetName).Range("A1").AutoFilter field:=columnIndex, Criteria1:=">=" & fromStatus, Operator:=xlAnd, Criteria2:="<=" & toStatus
    End If
    
End Function




'+------------------------------------------------------------------+







