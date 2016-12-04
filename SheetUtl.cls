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

'���̃N���X���W���[���̓V�[�g����p�̃��\�b�h��񋟂��܂��B
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+


'addSheet

'>�V�[�g���w�肵���V�[�g�̉E�ɒǉ����܂��B

'�S�Ă̈������ȗ������ꍇ�A�����I�ɐݒ肳�ꂽ�V�[�g���̃V�[�g�������ɒǉ�����܂��B
'����newSheetName�݂̂�n�����ꍇ�A�w�肵���V�[�g���̃V�[�g�������ɒǉ�����܂��B
'����newSheetName��afterSheetName��n�����ꍇ�A�w�肵���V�[�g���̃V�[�g��afterSheetName�̌�ɒǉ�����܂��B
'
'afterSheetName��n���ꍇ�AnewSheetName���ȗ��ł��Ȃ����߁AafterSheetName�݂̂�n���ăV�[�g���͎����Őݒ肳�������ꍇ�A
'afterSheetName:= "Sheet2" �̂悤�Ɉ��������w�肵�Ēl��n���ĉ������B
'
'newSheetName�Ƃ��Ċ����̃V�[�g�����w�肵���ꍇ�G���[�������܂��B���̏ꍇ�����I�ɐݒ肳�ꂽ�V�[�g���Ŗ����ɒǉ����s���܂��B


Function addSheet(Optional newSheetName As Variant, Optional afterSheetName As Variant)

If IsMissing(newSheetName) Then

    If IsMissing(afterSheetName) Then
        Worksheets.Add after:=Worksheets(Worksheets.Count)
    Else
    'newSheetName��n�����AafterSheetName��n���ɂ͈��������w�肵�Ēl��n���B
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

'>�V�[�g���w�肵���V�[�g�̍��ɒǉ����܂��B

'�S�Ă̈������ȗ������ꍇ�A�����I�ɐݒ肳�ꂽ�V�[�g���̃V�[�g����ԍ��ɒǉ�����܂��B
'����newSheetName�݂̂�n�����ꍇ�A�w�肵���V�[�g���̃V�[�g����ԍ��ɒǉ�����܂��B
'����newSheetName��beforeSheetName��n�����ꍇ�A�w�肵���V�[�g���̃V�[�g��beforeSheetName�̑O�ɒǉ�����܂��B
'
'beforeSheetName��n���ꍇ�AnewSheetName���ȗ��ł��Ȃ����߁AbeforeSheetName�݂̂�n���ăV�[�g���͎����Őݒ肳�������ꍇ�A
'beforeSheetName:= "Sheet2" �̂悤�Ɉ��������w�肵�Ēl��n���ĉ������B
'
'newSheetName�Ƃ��Ċ����̃V�[�g�����w�肵���ꍇ�G���[�������܂��B���̏ꍇ�����I�ɐݒ肳�ꂽ�V�[�g���ň�ԍ��ɒǉ����s���܂��B


Function addSheetBefore(Optional newSheetName As Variant, Optional beforeSheetName As Variant)

If IsMissing(newSheetName) Then

    If IsMissing(beforeSheetName) Then
        Worksheets.Add before:=Worksheets(1)
    Else
    'newSheetName��n�����AbeforeSheetName��n���ɂ͈��������w�肵�Ēl��n���B
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

'>�w�肵�����V�[�g�̖��O��ύX���܂��B

'����oldSheetName�͏ȗ��s�ł��B
'���ɑ��݂���V�[�g����n�����ꍇ�G���[�������܂��B

Function renameSheet(oldSheetName As String, newSheetName As String)

      Worksheets(oldSheetName).Name = newSheetName
     
End Function

'deleteSheet

'>�w�肵���V�[�g���폜���܂��B

'showAlert��False��n���ƍ폜���̃A���[�g��\�����Ȃ��悤�ɂȂ�܂��B
'showAleart�ȗ���������True��n���ƃA���[�g�L���̂܂܎��s���܂��B
'���݂��Ȃ��V�[�g���w�肷��ƃG���[�ƂȂ�܂��B
'����! �A���[�g�L���ɂ��Ă��A���[�g���Ȃ�Ȃ��ꍇ������܂��B�Č����̂Ȃ��o�O?

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
'>�w�肵���V�[�g�𖖔��ɃR�s�[���܂��B
'
'����newSheetName�ȗ����̓V�[�g���͎����I�ɐݒ肳��܂��B
Function copySheet(sheetName As String, Optional newSheetName As Variant)
  
Worksheets(sheetName).Copy after:=Worksheets(Worksheets.Count)
  
If IsMissing(newSheetName) Then
    
Else
    Worksheets(Worksheets.Count).Name = newSheetName
End If
  
End Function

'moveSheet
'>�w�肵���V�[�g���w�肵���V�[�g�̍��ֈړ����܂��B
'
'����afterSheetName�ȗ����͖����Ɉړ����܂��B
Function moveSheet(sheetName As String, Optional afterSheetName As Variant)
  
If IsMissing(afterSheetName) Then
     Worksheets(sheetName).Move after:=Worksheets(Worksheets.Count)
Else
     Worksheets(sheetName).Move after:=Worksheets(Worksheets(afterSheetName).index)
End If
  
End Function

'moveSheetBefore
'>�w�肵���V�[�g���w�肵���V�[�g�̍��ֈړ����܂��B
'
'����beforeSheetName�ȗ����͈�ԍ��ֈړ����܂��B
Function moveSheetBefore(sheetName As String, Optional beforeSheetName As Variant)
   
If IsMissing(beforeSheetName) Then
     Worksheets(sheetName).Move before:=Worksheets(1)
Else
     Worksheets(sheetName).Move before:=Worksheets(Worksheets(beforeSheetName).index)
End If
  
End Function




'+------------------------------------------------------------------+

