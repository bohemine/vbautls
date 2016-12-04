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

'���̃N���X���W���[���̓t�B���^����p�̃��\�b�h��񋟂��܂��B
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+


'setFilter

'>�V�[�g�ƍs�ԍ����w�肵�A�t�B���^��\�����܂��B
'��2����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
'�w�肵���V�[�g���̂����ꂩ�̉ӏ��Ńt�B���^���\���ς݂ł������ꍇ�͓��삵�܂���B
'�V���ɕ\���������ꍇ�A�t�B���^���\���ɂ��Ă�����s���ĉ������B
'
'������:setFilter "1","Sheet1"  �A  setFilter "1"

Function setFilter(rowNumber As String, Optional sheetName As Variant)

    Dim RowNumberStr As String
    RowNumberStr = rowNumber + ":" + rowNumber


    '��2����sheetName�̖����A�ȗ��ɂ���������
    If IsMissing(sheetName) Then
        ''���[�N�V�[�g�w��Ȃ��̏ꍇ�A�N�e�B�u�ȃV�[�g��I���̂��߉������Ȃ��B
        Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
            Worksheets(sheetName).Activate
    End If

    
    If ActiveSheet.AutoFilterMode = False Then
        Range(RowNumberStr).Select
        Selection.AutoFilter
    End If
 
End Function


'setHeaderFilter

'>�w�肵���V�[�g�̐擪�s�Ńt�B���^��\�����܂��B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
'�w�肵���V�[�g���̂����ꂩ�̉ӏ��Ńt�B���^���\���ς݂ł������ꍇ�́A�\�����������Ƃ͂���܂���B
'�V���ɕ\���������ꍇ�A�t�B���^���\���ɂ��Ă�����s���ĉ������B

Function setHeadFilter(Optional sheetName As Variant)

    If IsMissing(sheetName) Then
            setFilter "1", ActiveSheet.Name
        Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
            setFilter "1", sheetName
    End If
    
End Function


'clearFilter

'>�w�肵���V�[�g���̃t�B���^���\���ɂ��܂��B
'����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'�V�[�g���S��ɑ΂��ėL���B

Function clearFilter(Optional sheetName As Variant)

  If IsMissing(sheetName) Then
        ''���[�N�V�[�g�w��Ȃ��̏ꍇ�A�N�e�B�u�ȃV�[�g��I���̂��߉������Ȃ��B
        Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
              Worksheets(sheetName).Activate
    End If
  
    If ActiveSheet.AutoFilterMode = True Then
        Range("1:1").Select
        Selection.AutoFilter
    End If
    
End Function


'filter

'�V�[�g�Ɨ�ԍ��Ə����P���w�肵�t�B���^���s���܂��B
'��3����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'�t�B���^�͐擪�s�Ŏ��s����܂��B

'������:filter 3, ">2015/12/16", filterStr 1, ">20"

Function filter(columnIndex As Integer, status As Variant, Optional sheetName As Variant)

    If IsMissing(sheetName) Then
            Worksheets(ActiveSheet.Name).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status
        Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
            Worksheets(sheetName).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status
    End If
  
End Function


'filterOr

'�V�[�g�Ǝw�肵����ւQ�̏����̃t�B���^���s���܂��B
'��4����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'�t�B���^�͐擪�s�Ŏ��s����܂��B
'�Q�̏�����Or���Z�ƂȂ�܂��B

Function filterOr(columnIndex As Integer, status1 As String, status2 As String, Optional sheetName As Variant)

   If IsMissing(sheetName) Then
           
            Worksheets(ActiveSheet.Name).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status1, Operator:=xlOr, Criteria2:=status2

        Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
            Worksheets(sheetName).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status1, Operator:=xlOr, Criteria2:=status2
    End If
    
End Function

'filterAnd

'�V�[�g�Ǝw�肵����ւQ�̏����̃t�B���^���s���܂��B
'��4����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'�t�B���^�͐擪�s�Ŏ��s����܂��B
'�Q�̏�����And���Z�ƂȂ�܂��B

'������:�͈͂ōi�荞�� filterAnd 2, ">=1", "<=10", "Sheet1"

Function filterAnd(columnIndex As Integer, status1 As String, status2 As String, Optional sheetName As Variant)

   If IsMissing(sheetName) Then
           
            Worksheets(ActiveSheet.Name).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status1, Operator:=xlAnd, Criteria2:=status2

        Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
            Worksheets(sheetName).Range("A1").AutoFilter field:=columnIndex, Criteria1:=status1, Operator:=xlAnd, Criteria2:=status2
    End If
    
End Function


'filterSpan

'�V�[�g�Ǝw�肵����ւQ�̏����͈̔͂Ńt�B���^���s���܂��B
'��4����sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'�t�B���^�͐擪�s�Ŏ��s����܂��B
'�Q�̏�����And���Z�ƂȂ�܂��B

'�������̂̒l�������ΏۂɂȂ�B
'������:�͈͂ōi�荞�� filterSpan 2, "1", "100", "Sheet1"

Function filterSpan(columnIndex As Integer, fromStatus As String, toStatus As String, Optional sheetName As Variant)

   If IsMissing(sheetName) Then
           
            Worksheets(ActiveSheet.Name).Range("A1").AutoFilter field:=columnIndex, Criteria1:=">=" & fromStatus, Operator:=xlAnd, Criteria2:="<=" & toStatus

        Else
        ''���[�N�V�[�g�w�肠��̏ꍇ
            Worksheets(sheetName).Range("A1").AutoFilter field:=columnIndex, Criteria1:=">=" & fromStatus, Operator:=xlAnd, Criteria2:="<=" & toStatus
    End If
    
End Function




'+------------------------------------------------------------------+






