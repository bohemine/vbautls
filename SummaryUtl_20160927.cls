VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummaryUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Class module:SummaryUtl
'Version:0.0.1
'Last modified:2016-09-18

'���̃N���X���W���[���͏W�v�p�̃��\�b�h��񋟂��܂��B
'
'+------------------------------------------------------------------+





'+--Release notes---------------------------------------------------+
'2016-09-18 ����
'
'
'+------------------------------------------------------------------+






'+--Public Functions------------------------------------------------+


'removeHeadRow
'�w�肵���V�[�g�̐擪�s���폜���܂��B

'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B
Function sum(Optional sheetName As Variant)
  
  
    
End Function

'total

'�w�肵���V�[�g�̗�̐��l�f�[�^���W�v���܂��B
'�J��������"C"�̂悤�Ɏw�肷��B
'
Function sumColumn(columnName As String, Optional sheetName As Variant) As Long

    Dim columnName_forRange As String
    columnName_forRange = columnName + ":" + columnName
 
 
   If IsMissing(sheetName) Then
    sumColumn = WorksheetFunction.Subtotal(9, Worksheets(ActiveSheet.Name).Range(columnName_forRange))
  
   Else
    sumColumn = WorksheetFunction.Subtotal(9, Worksheets(sheetName).Range(columnName_forRange))
   End If

    
End Function


















'+------------------------------------------------------------------+





