VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelBookSheetCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************************
' �ϐ�
'**************************************************
Private ExcelFullPath      As String  ' �擾����Excel�̃t���p�X
Private ExcelSheetName     As String  ' �擾����Excel�̃V�[�g��
Private ExcelStartColPoint As Double  ' �擾����Excel�̎擾�J�n��
Private ExcelStartRowPoint As Double  ' �擾����Excel�̎擾�J�n�s
Private ExcelGetColCount   As Double  ' �擾����Excel���牽��܂Ŏ擾���邩
Private ExcelGetRowCount   As Double  ' �擾����Excel���牽�s�܂Ŏ擾���邩
Private ExcelStopColFlg    As Boolean ' ����̗񂪋󔒂̏ꍇ�f�[�^�擾���~�߂邩
Private ExcelStopColPoint  As Double  ' ����̗�(stopColFlg��True�̏ꍇ�L��)
Private SheetItems         As Object  ' �擾�����G�N�Z���t�@�C���̃f�[�^

'**************************************************
' �R���X�g���N�^�ł�
'**************************************************
Public Sub Class_Initialize()
    Set SheetItems = CreateObject("System.Collections.ArrayList")
End Sub

'**************************************************
' �f�X�g���N�^�ł�
'**************************************************
Public Sub Class_Terminate()
End Sub

'*****************************************************************************
'* Public    �FGetExcelData                                                  *
'*           �F�w�肵���擾����G�N�Z���������Ƀf�[�^���擾���܂��B        *
'* ����      �FfullPath      �G�N�Z���t�@�C���̃t���p�X                      *
'*           �FsheetName�@�@ �擾����V�[�g��                                *
'*           �FstartColPoint �擾�J�n��                                      *
'*           �FstartRowPoint �擾�J�n�s                                      *
'*           �FgetColCount   ����܂Ŏ擾���邩    �@                        *
'*           �FgetRowCount�@ ���s�܂Ŏ擾���邩                              *
'*           �FstopColFlg    ����̗񂪋󔒂̏ꍇ�f�[�^�擾���~�߂邩        *
'*           �F              True :�~�߂�                                    *
'*           �F              False:�~�߂Ȃ�                                  *
'*           �FstopColPoint  ����̗�                                        *
'*           �F              stopColFlg��True�̏ꍇ�L��                      *
'* �ԋp�l    �F�Ȃ�                                                          *
'* ���l      �F�o�^�����f�[�^�͑��݂��Ă��鎖��O��Ƃ��Ă��܂�              *
'*****************************************************************************
Public Sub GetExcelData(fullPath As String, sheetName As String, startColPoint As Double, startRowPoint As Double, getColCount As Double, getRowCount As Double, stopColFlg As Boolean, stopColPoint As Double)
    ExcelFullPath = fullPath
    ExcelSheetName = sheetName
    ExcelStartColPoint = startColPoint
    ExcelStartRowPoint = startRowPoint
    ExcelGetColCount = getColCount
    ExcelGetRowCount = getRowCount
    ExcelStopColFlg = stopColFlg
    ExcelStopColPoint = stopColPoint
    
    Call GetExcelData_Start
    
End Sub

'*****************************************************************************
'* Public    �FGetExcelData_Start                                            *
'* ����      �F�Ȃ�                                                          *
'* �ԋp�l    �F�Ȃ�                                                          *
'*****************************************************************************
Private Sub GetExcelData_Start()
    Dim exBook As Workbook ' �I�u�W�F�N�g
    Dim tCol As Double     ' ��
    Dim tRow As Double     ' �s
    Dim itemCnt As Integer ' ��擾�J�E���g

    Set exBook = Workbooks.Open(Filename:=ExcelFullPath, UpdateLinks:=False, ReadOnly:=True) ' Excel��ǂݎ���p�ŊJ��
    
    For tRow = ExcelStartRowPoint To ExcelStartRowPoint + ExcelGetRowCount - 1 ' �擾�f�[�^���̃f�[�^��z��Ɋi�[
        
        ' ����̗񂪋󔒂̏ꍇ�f�[�^�擾���~�߂�
        If ExcelStopColFlg = True And exBook.Sheets(ExcelSheetName).Cells(tRow, ExcelStopColPoint).Value = "" Then
            Exit For
        End If
    
        ' �f�[�^������
        ReDim itemAry(0 To ExcelGetColCount - 1) ' �擾�f�[�^�ꎞ�i�[�p
        itemCnt = 0                              ' �f�[�^������
        
        For tCol = ExcelStartColPoint To ExcelStartColPoint + ExcelGetColCount - 1  ' �擾�f�[�^���̃f�[�^��z��Ɋi�[
            itemAry(itemCnt) = exBook.Sheets(ExcelSheetName).Cells(tRow, tCol).Value
            itemCnt = itemCnt + 1
        Next
        SheetItems.Add (itemAry)
    Next

    exBook.Close ' Excel�����

End Sub

'*****************************************************************************
'* Public    �FSetSheetDataCopy                                              *
'*           �F�V�[�g�̎w��ʒu�Ɏ擾�����f�[�^���R�s�[���܂��B              *
'* ����      �FsheetName�@�@ �R�s�[����V�[�g��                              *
'*           �FstartColPoint �J�n��                                          *
'*           �FstartRowPoint �J�n�s                                          *
'* �ԋp�l    �F�Ȃ�                                                          *
'*****************************************************************************
Public Sub SetSheetDataCopy(ByVal sheetName As String, startColPoint As Double, startRowPoint As Double)
    Dim item          ' ���[�v�p
    Dim cellItem      ' ���[�v�p
    Dim row As Double ' �s�J�E���g�p
    Dim col As Double ' ��J�E���g�p
    
    ActiveWorkbook.Worksheets(sheetName).Activate ' �Ώۂ̃V�[�g���A�N�e�B�u�ɂ���
    
    row = startRowPoint ' �J�n�s
    
    For Each item In SheetItems ' �擾�����f�[�^�����[�v����
        col = startColPoint ' �J�n�񏉊���
        For Each cellItem In item
            Cells(row, col).Value = cellItem
            col = col + 1
        Next
        row = row + 1
    Next
    
End Sub

'**************************************************
' count
'**************************************************
Property Get count() As Integer
    count = SheetItems.count
End Property

'**************************************************
' ToArray
'**************************************************
Function ToArray() As Variant
    ToArray = SheetItems.ToArray
End Function

'**************************************************
' �T���v�����\�b�h �擾�����f�[�^�����[�v�Ŏ擾����
'**************************************************
Public Sub SampleGetData()
    Dim item               ' arraylist���[�v�p
    Dim itemCnt As Double ' �f�[�^�o�͗�J�E���g�p
    
    For Each item In Me.ToArray ' �擾�����f�[�^�����[�v����
        For itemCnt = 0 To 2
            Debug.Print item(itemCnt)
        Next
    Next
End Sub
