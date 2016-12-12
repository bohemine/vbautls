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


'���t���󂯎�蕽�����ǂ����𔻒肵�܂��B

Function isWeekday(day As Date) As Boolean

  Dim dow As Integer
  dow = Weekday(day)

  If dow = vbSunday Or dow = vbSaturday Then
    isWeekday = False
  Else
    isWeekday = True
  End If

End Function



'���t���󂯎��y�����ǂ����𔻒肵�܂��B
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



'���t���󂯎��,���̎��̉ғ�����Ԃ��܂��B
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
'���t���󂯎�荑���̏j�����ǂ����𔻒肵�܂��B
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


'�z��ɗv�f���܂܂�邩���肵�܂��B
'
Function testExists(arr() As Variant, val As Date) As Boolean
  
  testExists = exists(arr(), val)
End Function

'�z��ɗv�f���܂܂�邩���肵�܂��B
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











