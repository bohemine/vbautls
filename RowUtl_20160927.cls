VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'+--Overview--------------------------------------------------------+
'Class module:RowUtl
'Version:0.0.1
'Last modified:2016-09-17

'���̃N���X���W���[���͗񑀍�p�̃��\�b�h��񋟂��܂��B
'
'+------------------------------------------------------------------+





'+--Release notes---------------------------------------------------+
'2016-09-17 ����
'
'
'+------------------------------------------------------------------+






'+--Public Functions------------------------------------------------+


'removeHeadRow
'�w�肵���V�[�g�̐擪�s���폜���܂��B

'sheetName�ȗ����̓A�N�e�B�u�ȃV�[�g��ΏۂƂ��܂��B
Function removeHeadRow(Optional sheetName As Variant)
  
   If IsMissing(sheetName) Then
       Worksheets(ActiveSheet.Name).Range("A1").Offset(0, 0).EntireRow.Delete
   Else
       Worksheets(sheetName).Range("A1").Offset(0, 0).EntireRow.Delete
   End If
    
End Function

'removeDuplicateRows
'�w�肵���V�[�g�̏d���s���폜���܂��B

'columnIndex�֑ΏۂƂ����ԍ���n���ĉ������B
'header�֐擪�s�̈������w�肵�ĉ������BxlYes[�w�b�_�Ƃ��Ĉ���]�@or xlNo[�w�b�_�Ƃ��Ĉ���Ȃ�]�@or xlGuess [Excel�ɔ��f������]

'"A1"����Z���Ƃ����Ƃ���CurrentRegion���Ώ۔͈͂ł��B
'Excel2007�ȍ~�ł̂ݎg�p�B
Function removeDuplicateRows(columnIndex As Integer, header As Variant, sheetName As String)
  
  Worksheets(sheetName).Range("A1").CurrentRegion.RemoveDuplicates Columns:=columnIndex, header:=header

End Function


'removeDuplicateRows2
'�w�肵���V�[�g�ɂ�2�̗�����ɏd���s���폜���܂��B

'columnIndex�֑ΏۂƂ����ԍ���n���ĉ������B
'header�֐擪�s�̈������w�肵�ĉ������BxlYes[�w�b�_�Ƃ��Ĉ���]�@or xlNo[�w�b�_�Ƃ��Ĉ���Ȃ�]�@or xlGuess [Excel�ɔ��f������]

'"A1"����Z���Ƃ����Ƃ���CurrentRegion���Ώ۔͈͂ł��B
'Excel2007�ȍ~�ł̂ݎg�p�B
Function removeDuplicateRows2(columnIndex1 As Integer, columnIndex2 As Integer, header As Variant, sheetName As String)
  
  Worksheets(sheetName).Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(columnIndex1, columnIndex2), header:=header

End Function


'copyRow

'>�V�[�g�ƍs���w�肵�A�w�肳�ꂽ�s�փR�s�[���܂��B

'�\��t�����s���܂��B
Function copyRow(sheetName_src As String, rowName_src As String, sheetName_dst As String, rowName_dst As String)

    Worksheets(sheetName_src).Range(rowName_src + ":" + rowName_src).Copy Worksheets(sheetName_dst).Range(rowName_dst + ":" + rowName_dst)
 
End Function

'+------------------------------------------------------------------+













