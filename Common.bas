Attribute VB_Name = "Common"
Option Explicit

'文字コードの種類
Public Enum charsetType
    Shift_JIS = 0
    UTF8 = 1
End Enum

'文字コードの種類
Public Enum SetPathType
    Desktop
    AppData
    StartMenu
    SendTo
    MyDocuments
    Workbook
End Enum


'*****************************************************************************
'* Public    ：PathSelect                                                    *
'*           ：パスを取得                                                    *
'* 引数      ：SetPath   パスの種類                                          *
'* 返却値    ：String    パス名                                              *
'* 備考      ：文字列　　　　特殊フォルダ                                    *
'*           ：Desktop　　　 個人用のデスクトップフォルダ                    *
'*           ：AppData　　　 個人用のApplication Data フォルダ               *
'*           ：StartMenu　　 個人用のスタートメニューフォルダ                *
'*           ：SendTo　　　　個人用のSendToフォルダ                          *
'*           ：MyDocuments   個人用のマイドキュメントフォルダ                *
'*           ：Workbook      このファイルがあるフォルダ                      *
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
'* Public    ：PathSelect                                                    *
'*           ：テキストファイル作成                                          *
'* 引数      ：saveFilePath ファイルのフルパス                               *
'*           ：outputText   出力するファイルパス                             *
'*           ：code         文字コードの種類 Shift_JISまたはUTF8             *
'*           ：overWrite    1:ファイル有り時上書きしない、2:上書きする       *
'* 返却値    ：なし                                                          *
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
        .SaveToFile saveFilePath, overWrite '1:ファイル有り時上書きしない、2:上書きする
        .Close
    End With

End Sub

'*****************************************************************************
'* Public    ：GetPath                                                       *
'*           ：ファイルパスからディレクトリパスのみ取得                      *
'* 引数      ：FilePathDir フルパス                                          *
'* 返却値    ：String      ファイル名                                        *
'* 備考      ：例：フルパスはC:\aaa\bbb\ccc\test000.txtとする                *
'*           ：C:\aaa\bbb\ccc                                                *
'*****************************************************************************
Public Function GetPath(ByVal FilePathDir As String) As String
    Dim w_num As Integer
    w_num = InStrRev(FilePathDir, "\")
    GetPath = Left(FilePathDir, w_num - 1)
End Function

'*****************************************************************************
'* Public    ：GetFileName                                                   *
'*           ：フルパスからファイル名のみ取得                                *
'* 引数      ：FilePathDir フルパス                                          *
'*           ：ExtTrim     拡張子有無                                        *
'* 返却値    ：String      ファイル名                                        *
'* 備考      ：例：フルパスはC:\aaa\bbb\ccc\test000.txtとする                *
'*           ：拡張子を取る     ExtTrim = True  結果 test000                 *
'*           ：拡張子を取らない ExtTrim = False 結果 test000.txt             *
'*           ：※フルパスに.が無い場合はファイルが無いとみなす               *
'*****************************************************************************
Public Function GetFileName(ByVal FilePathDir As String, _
                            ByVal ExtTrim As Boolean) As String
    If 0 = InStr(1, FilePathDir, ".") Then '.が無い場合はファイルが無いとみなす
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
'* Public    ：RunBatFile                                                    *
'*           ：バッチファイルを起動します                                    *
'* 引数      ：path     バッチファイルパス                                   *
'*           ：batType  バッチの起動                                         *
'*           ：  0: 非表示                                                   *
'*           ：  1: 通常表示                                                 *
'*           ：  2: 最小化                                                   *
'*           ：  3: 最大化                                                   *
'*           ：exeType                                                       *
'*           ：  True:バッチファイルの処理が終了するまで待つ                 *
'*           ：  False:バッチファイルの処理終了をまたずに次行のコードを実行  *
'* 返却値    ：なし                                                          *
'*****************************************************************************
Public Sub RunBatFile(path As String, batType As Integer, exeType As Boolean)
    
    Dim ShellObject As Object
    Dim MsgBoxRet As String

    Set ShellObject = CreateObject("WScript.Shell")
    ShellObject.Run """" & path & """", batType, exeType

End Sub

'*****************************************************************************
'* Public    ：RepEE4M                                                       *
'*           ：ExecuteExcel4Macroのラッパー関数                              *
'*           ：ExecuteExcel4Macroを利用してセルの値を取得すると空白セルも0で *
'*           ：取得するので、文字列から比較して空白か0か識別しています       *
'* 引数      ：ee4m ExecuteExcel4Macroの引数                                 *
'* 返却値    ：ExecuteExcel4Macroで取得した値                                *
'*****************************************************************************
Public Function RepEE4M(ee4m As String) As String
    Dim ret As String
    
    If Application.ExecuteExcel4Macro("LEN( " & ee4m & " )") > 0 Then
        ret = Application.ExecuteExcel4Macro(ee4m)
    End If

    RepEE4M = ret
End Function

'*****************************************************************************
'* Public    ：CreateEE4MPath                                                *
'*           ：ExecuteExcel4Macroのパスを作成します。                        *
'* 引数      ：excelBookFullPath Excelファイルのフルパス                     *
'*           ：sheet             シート名                                    *
'* 返却値    ：ExecuteExcel4Macroのパス                                      *
'* 備考      ：セル位置（R1C1）は追加していません                            *
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


