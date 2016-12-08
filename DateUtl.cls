VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'+--Overview--------------------------------------------------------+
'Class module: DateUtl
'Version:0.0.1
'
'+------------------------------------------------------------------+


'日付を受け取り平日かどうかを判定します。

Function isWeekday(day As Date) As Boolean

  Dim dow As Integer
  dow = Weekday(day)

  If dow = vbSunday Or dow = vbSaturday Then
    isWeekday = False
  Else
    isWeekday = True
  End If

End Function



'日付を受け取り土日かどうかを判定します。
'
Function isSunOrSaturday(day As Date) As Boolean

  Dim dow As Integer
  dow = Weekday(day)

  If dow = vbSunday Or dow = vbSaturday Then
    isSunOrSaturday = True
  Else
    isSunOrSaturday = False
  End If
   
End Function



'日付を受け取り,その次の稼動日を返します。
'
Function nextWorkingDay(day As Date)
 
 If Not isHoliday(day) Then
   nextWorkingDay = DateAdd("d", 1, day)
 Else
  
  Do While isHoliday(day)
    nextWorkingDay = DateAdd("d", 1, day)
  Loop
 
 End If
 
End Function

'
'
Function preWorkingDay()

End Function
'日付を受け取り国民の祝日かどうかを判定します。
'
Function isNationalHoliday(day As Date) As Boolean

  Dim dow As Integer
  dow = Weekday(day)

  If dow = vbSunday Or dow = vbSaturday Then
    isSunOrSaturday = True
  Else
    isSunOrSaturday = False
  End If
   
End Function


'配列に要素が含まれるか判定します。
'
Function testExists(arr() As Variant, val As Date) As Boolean
  
  testExists = exists(arr(), val)
End Function

'配列に要素が含まれるか判定します。
'
Function exists(arr() As Variant, val As Date) As Boolean
  
  Dim i As Integer
  
  i = filter(CStr(arr()), val, True)
  
  
  If UBound(i) <> -1 Then
    exists = True
  Else
    exists = False
  End If

End Function












