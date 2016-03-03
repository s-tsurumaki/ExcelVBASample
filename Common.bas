Attribute VB_Name = "Common"
Option Explicit

'�����R�[�h�̎��
Public Enum charsetType
    Shift_JIS = 0
    UTF8 = 1
End Enum

'�����R�[�h�̎��
Public Enum SetPathType
    Desktop
    AppData
    StartMenu
    SendTo
    MyDocuments
    Workbook
End Enum


'*****************************************************************************
'* Public    �FPathSelect                                                    *
'*           �F�p�X���擾                                                    *
'* ����      �FSetPath   �p�X�̎��                                          *
'* �ԋp�l    �FString    �p�X��                                              *
'* ���l      �F������@�@�@�@����t�H���_                                    *
'*           �FDesktop�@�@�@ �l�p�̃f�X�N�g�b�v�t�H���_                    *
'*           �FAppData�@�@�@ �l�p��Application Data �t�H���_               *
'*           �FStartMenu�@�@ �l�p�̃X�^�[�g���j���[�t�H���_                *
'*           �FSendTo�@�@�@�@�l�p��SendTo�t�H���_                          *
'*           �FMyDocuments   �l�p�̃}�C�h�L�������g�t�H���_                *
'*           �FWorkbook      ���̃t�@�C��������t�H���_                      *
'*****************************************************************************
Public Function PathSelect(ByVal SetPath As SetPathType) As String
    Dim ShellObject As Object
    Dim SetPathStr As String
    
    Select Case SetPath
        Case SetPathType.Desktop
            SetPathStr = "Desktop"
        Case SetPathType.AppData
            SetPathStr = "AppData"
        Case SetPathType.StartMenu
            SetPathStr = "StartMenu"
        Case SetPathType.SendTo
            SetPathStr = "SendTo"
        Case SetPathType.MyDocuments
            SetPathStr = "MyDocuments"
        Case SetPathType.Workbook
            SetPathStr = "Workbook"
        Case Else
            SetPathStr = "Workbook"
    End Select
    
    
    If SetPathStr = "Desktop" Or SetPathStr = "AppData" Or _
       SetPathStr = "StartMenu" Or SetPathStr = "SendTo" Or _
       SetPathStr = "MyDocuments" Then
        Set ShellObject = CreateObject("IWshRuntimeLibrary.WshShell")
        PathSelect = ShellObject.SpecialFolders(SetPathStr)
    ElseIf SetPathStr = "Workbook" Then
        PathSelect = ThisWorkbook.path
    Else
        PathSelect = ThisWorkbook.path
    End If
End Function

'*****************************************************************************
'* Public    �FPathSelect                                                    *
'*           �F�e�L�X�g�t�@�C���쐬                                          *
'* ����      �FsaveFilePath �t�@�C���̃t���p�X                               *
'*           �FoutputText   �o�͂���t�@�C���p�X                             *
'*           �Fcode         �����R�[�h�̎�� Shift_JIS�܂���UTF8             *
'*           �FoverWrite    1:�t�@�C���L�莞�㏑�����Ȃ��A2:�㏑������       *
'* �ԋp�l    �F�Ȃ�                                                          *
'*****************************************************************************
Public Sub CreateText(saveFilePath As String, outputText As String, code As charsetType, overWrite As Integer)

    Dim fso
    Dim codeStr As String
    Dim dirPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    dirPath = fso.GetParentFolderName(saveFilePath)

    If (Not fso.FolderExists(dirPath)) Then
        Call fso.CreateFolder(dirPath)
    End If
    
    If code = Shift_JIS Then
        codeStr = "Shift_JIS"
    Else
        codeStr = "UTF-8"
    End If
    
    With CreateObject("ADODB.Stream")
        .Charset = codeStr
        .Open
        .WriteText outputText
        .SaveToFile saveFilePath, overWrite '1:�t�@�C���L�莞�㏑�����Ȃ��A2:�㏑������
        .Close
    End With

End Sub

'*****************************************************************************
'* Public    �FGetPath                                                       *
'*           �F�t�@�C���p�X����f�B���N�g���p�X�̂ݎ擾                      *
'* ����      �FFilePathDir �t���p�X                                          *
'* �ԋp�l    �FString      �t�@�C����                                        *
'* ���l      �F��F�t���p�X��C:\aaa\bbb\ccc\test000.txt�Ƃ���                *
'*           �FC:\aaa\bbb\ccc                                                *
'*****************************************************************************
Public Function GetPath(ByVal FilePathDir As String) As String
    Dim w_num As Integer
    w_num = InStrRev(FilePathDir, "\")
    GetPath = Left(FilePathDir, w_num - 1)
End Function

'*****************************************************************************
'* Public    �FGetFileName                                                   *
'*           �F�t���p�X����t�@�C�����̂ݎ擾                                *
'* ����      �FFilePathDir �t���p�X                                          *
'*           �FExtTrim     �g���q�L��                                        *
'* �ԋp�l    �FString      �t�@�C����                                        *
'* ���l      �F��F�t���p�X��C:\aaa\bbb\ccc\test000.txt�Ƃ���                *
'*           �F�g���q�����     ExtTrim = True  ���� test000                 *
'*           �F�g���q�����Ȃ� ExtTrim = False ���� test000.txt             *
'*           �F���t���p�X��.�������ꍇ�̓t�@�C���������Ƃ݂Ȃ�               *
'*****************************************************************************
Public Function GetFileName(ByVal FilePathDir As String, _
                            ByVal ExtTrim As Boolean) As String
    If 0 = InStr(1, FilePathDir, ".") Then '.�������ꍇ�̓t�@�C���������Ƃ݂Ȃ�
        GetFileName = "False"
    Else
        If ExtTrim = True Then
            GetFileName = Mid(FilePathDir, InStrRev(FilePathDir, "\") + 1, Len(FilePathDir) - InStrRev(FilePathDir, "\") - (Len(FilePathDir) - InStrRev(FilePathDir, ".")) - 1)
        Else
            GetFileName = Mid(FilePathDir, InStrRev(FilePathDir, "\") + 1, Len(FilePathDir) - InStrRev(FilePathDir, "\"))
        End If
    End If
End Function

'*****************************************************************************
'* Public    �FRunBatFile                                                    *
'*           �F�o�b�`�t�@�C�����N�����܂�                                    *
'* ����      �Fpath     �o�b�`�t�@�C���p�X                                   *
'*           �FbatType  �o�b�`�̋N��                                         *
'*           �F  0: ��\��                                                   *
'*           �F  1: �ʏ�\��                                                 *
'*           �F  2: �ŏ���                                                   *
'*           �F  3: �ő剻                                                   *
'*           �FexeType                                                       *
'*           �F  True:�o�b�`�t�@�C���̏������I������܂ő҂�                 *
'*           �F  False:�o�b�`�t�@�C���̏����I�����܂����Ɏ��s�̃R�[�h�����s  *
'* �ԋp�l    �F�Ȃ�                                                          *
'*****************************************************************************
Public Sub RunBatFile(path As String, batType As Integer, exeType As Boolean)
    
    Dim ShellObject As Object
    Dim MsgBoxRet As String

    Set ShellObject = CreateObject("WScript.Shell")
    ShellObject.Run """" & path & """", batType, exeType

End Sub

'*****************************************************************************
'* Public    �FRepEE4M                                                       *
'*           �FExecuteExcel4Macro�̃��b�p�[�֐�                              *
'*           �FExecuteExcel4Macro�𗘗p���ăZ���̒l���擾����Ƌ󔒃Z����0�� *
'*           �F�擾����̂ŁA�����񂩂��r���ċ󔒂�0�����ʂ��Ă��܂�       *
'* ����      �Fee4m ExecuteExcel4Macro�̈���                                 *
'* �ԋp�l    �FExecuteExcel4Macro�Ŏ擾�����l                                *
'*****************************************************************************
Public Function RepEE4M(ee4m As String) As String
    Dim ret As String
    
    If Application.ExecuteExcel4Macro("LEN( " & ee4m & " )") > 0 Then
        ret = Application.ExecuteExcel4Macro(ee4m)
    End If

    RepEE4M = ret
End Function

'*****************************************************************************
'* Public    �FCreateEE4MPath                                                *
'*           �FExecuteExcel4Macro�̃p�X���쐬���܂��B                        *
'* ����      �FexcelBookFullPath Excel�t�@�C���̃t���p�X                     *
'*           �Fsheet             �V�[�g��                                    *
'* �ԋp�l    �FExecuteExcel4Macro�̃p�X                                      *
'* ���l      �F�Z���ʒu�iR1C1�j�͒ǉ����Ă��܂���                            *
'*****************************************************************************
Public Function CreateEE4MPath(excelBookFullPath As String, sheet As String) As String
    Dim ret As String
    Dim excelBookPath As String
    Dim excelFileName As String
    
    excelBookPath = Common.GetPath(file)
    excelFileName = Common.GetFileName(file, False)
    
    ret = "'" & excelBookPath & "\[" & excelFileName & "]" & sheet & "'!"
    
    CreateEE4MPath = ret
End Function


