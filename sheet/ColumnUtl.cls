VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnUtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'+--Overview--------------------------------------------------------+
'Class module:ColumnUtl
'Version:1.0.0

'���̃N���X���W���[���͗񑀍�p�̃��\�b�h��񋟂��܂��B
'
'+------------------------------------------------------------------+




'+--Public Functions------------------------------------------------+

'�V�[�g�Ɨ���w�肵�A���̗���폜���܂��B
'
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function removeColumns(columnsName As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnsName).Delete
Else
   Worksheets(sheetName).Range(columnsName).Delete
End If
    
End Function



'�V�[�g�Ɨ���w�肵�A���}�����܂��B
'�N���b�v�{�[�h�ɗ񂪊i�[����Ă���ꍇ�A���ꂪ�\��t������܂��B
'�N���b�v�{�[�h�ɗ�ȊO�̗L���łȂ��͈͂̃f�[�^������ꍇ�G���[��������\��������܂��B
'
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function insertColumns(columnsName As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnsName).Insert
Else
   Worksheets(sheetName).Range(columnsName).Insert
End If
    
End Function



'�w�肵����̃w�b�_�[����t���܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
Function renameHeader(columnName As String, header As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnName + "1").Value = header
Else
   Worksheets(sheetName).Range(columnName + "1").Value = header
End If
    
End Function



'�V�[�g�Ɨ���w�肵�A���̗���w�肳�ꂽ��փR�s�[���܂��B
'�\��t�����s���܂��B

Function copyColumns(sheetName_src As String, columnName_src As String, sheetName_dst As String, columnName_dst As String)

    Worksheets(sheetName_src).Range(columnName_src + ":" + columnName_src).copy Worksheets(sheetName_dst).Range(columnName_dst + ":" + columnName_dst)

End Function



'�V�[�g�Ɨ���w�肵�A���̗�ɂ���f�[�^���N���A���܂��B
'
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function clearColumnsContents(columnsName As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnsName).ClearContents
Else
   Worksheets(sheetName).Range(columnsName).ClearContents
End If
    
End Function



'�V�[�g�Ɨ���w�肵�A���̗�ɂ���f�[�^��؂���܂��B
'
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function cutColumns(columnsName As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
   Worksheets(ActiveSheet.Name).Range(columnsName).Cut
Else
   Worksheets(sheetName).Range(columnsName).Cut
End If
    
End Function



'�V�[�g�Ɨ���w�肵�A���̗���w�肳�ꂽ��ֈړ����܂��B
'
'�؂������񂪓\��t������������ɂ���ꍇ�A�w�肵������P���̗�ɓ\��t�����܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function moveColumns(columnsName_src As String, columnName_dst As String, Optional sheetName As Variant)

If IsMissing(sheetName) Then
    cutColumns columnsName_src
    insertColumns columnName_dst
Else
    cutColumns columnsName_src, sheetName
    insertColumns columnName_dst, sheetName
End If
    
End Function



'�񖼂��w�肵�A���̗�ԍ��Ԃ��܂��B
'return Long
'
'�؂������񂪓\��t������������ɂ���ꍇ�A�w�肵������P���̗�ɓ\��t�����܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B

Function getColumnNumber(columnsName As String) As Long

    getColumnNumber = ActiveSheet.Range(columnsName + "1").column

End Function



'��ԍ����w�肵�A���̗񖼕Ԃ��܂��B

'�؂������񂪓\��t������������ɂ���ꍇ�A�w�肵������P���̗�ɓ\��t�����܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'
Function getColumnName(columnIndex As Integer)
    
    getColumnName = Split(Cells(, columnIndex).Address, "$")(1)

End Function



'�擪�s�ɂ���t�B�[���h�����w�肵�A�ŏ��Ɍ���������̗񖼂�Ԃ��܂��B
'
'�؂������񂪓\��t������������ɂ���ꍇ�A�w�肵������P���̗�ɓ\��t�����܂��B
'����sheetName�̂ݏȗ����̓A�N�e�B�u�ȃV�[�g���I������܂��B
'���S�ɓ���̕������ΏۂƂ��܂��B������̈ꕔ���}�b�`���Ă��ΏۂɂȂ�܂���B

Function getColumnNameByField(target As String, Optional sheetName As Variant) As String

If IsMissing(sheetName) Then
    Worksheets(ActiveSheet.Name).Range("1:1").Find(what:=target, LookAt:=xlWhole).Activate
    getColumnNameByField = Split(ActiveCell.Address, "$")(1)

Else
    Worksheets(sheetName).Activate
    Range("1:1").Find(what:=target, LookAt:=xlWhole).Activate
    getColumnNameByField = Split(ActiveCell.Address, "$")(1)
   
End If

End Function

'+------------------------------------------------------------------+



