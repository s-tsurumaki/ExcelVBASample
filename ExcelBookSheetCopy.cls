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
' 変数
'**************************************************
Private ExcelFullPath      As String  ' 取得するExcelのフルパス
Private ExcelSheetName     As String  ' 取得するExcelのシート名
Private ExcelStartColPoint As Double  ' 取得するExcelの取得開始列
Private ExcelStartRowPoint As Double  ' 取得するExcelの取得開始行
Private ExcelGetColCount   As Double  ' 取得するExcelから何列まで取得するか
Private ExcelGetRowCount   As Double  ' 取得するExcelから何行まで取得するか
Private ExcelStopColFlg    As Boolean ' 特定の列が空白の場合データ取得を止めるか
Private ExcelStopColPoint  As Double  ' 特定の列(stopColFlgがTrueの場合有効)
Private SheetItems         As Object  ' 取得したエクセルファイルのデータ

'**************************************************
' コンストラクタです
'**************************************************
Public Sub Class_Initialize()
    Set SheetItems = CreateObject("System.Collections.ArrayList")
End Sub

'**************************************************
' デストラクタです
'**************************************************
Public Sub Class_Terminate()
End Sub

'*****************************************************************************
'* Public    ：GetExcelData                                                  *
'*           ：指定した取得するエクセル情報を元にデータを取得します。        *
'* 引数      ：fullPath      エクセルファイルのフルパス                      *
'*           ：sheetName　　 取得するシート名                                *
'*           ：startColPoint 取得開始列                                      *
'*           ：startRowPoint 取得開始行                                      *
'*           ：getColCount   何列まで取得するか    　                        *
'*           ：getRowCount　 何行まで取得するか                              *
'*           ：stopColFlg    特定の列が空白の場合データ取得を止めるか        *
'*           ：              True :止める                                    *
'*           ：              False:止めない                                  *
'*           ：stopColPoint  特定の列                                        *
'*           ：              stopColFlgがTrueの場合有効                      *
'* 返却値    ：なし                                                          *
'* 備考      ：登録したデータは存在している事を前提としています              *
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
'* Public    ：GetExcelData_Start                                            *
'* 引数      ：なし                                                          *
'* 返却値    ：なし                                                          *
'*****************************************************************************
Private Sub GetExcelData_Start()
    Dim exBook As Workbook ' オブジェクト
    Dim tCol As Double     ' 列
    Dim tRow As Double     ' 行
    Dim itemCnt As Integer ' 列取得カウント

    Set exBook = Workbooks.Open(Filename:=ExcelFullPath, UpdateLinks:=False, ReadOnly:=True) ' Excelを読み取り専用で開く
    
    For tRow = ExcelStartRowPoint To ExcelStartRowPoint + ExcelGetRowCount - 1 ' 取得データ分のデータを配列に格納
        
        ' 特定の列が空白の場合データ取得を止める
        If ExcelStopColFlg = True And exBook.Sheets(ExcelSheetName).Cells(tRow, ExcelStopColPoint).Value = "" Then
            Exit For
        End If
    
        ' データ初期化
        ReDim itemAry(0 To ExcelGetColCount - 1) ' 取得データ一時格納用
        itemCnt = 0                              ' データ初期化
        
        For tCol = ExcelStartColPoint To ExcelStartColPoint + ExcelGetColCount - 1  ' 取得データ分のデータを配列に格納
            itemAry(itemCnt) = exBook.Sheets(ExcelSheetName).Cells(tRow, tCol).Value
            itemCnt = itemCnt + 1
        Next
        SheetItems.Add (itemAry)
    Next

    exBook.Close ' Excelを閉じる

End Sub

'*****************************************************************************
'* Public    ：SetSheetDataCopy                                              *
'*           ：シートの指定位置に取得したデータをコピーします。              *
'* 引数      ：sheetName　　 コピーするシート名                              *
'*           ：startColPoint 開始列                                          *
'*           ：startRowPoint 開始行                                          *
'* 返却値    ：なし                                                          *
'*****************************************************************************
Public Sub SetSheetDataCopy(ByVal sheetName As String, startColPoint As Double, startRowPoint As Double)
    Dim item          ' ループ用
    Dim cellItem      ' ループ用
    Dim row As Double ' 行カウント用
    Dim col As Double ' 列カウント用
    
    ActiveWorkbook.Worksheets(sheetName).Activate ' 対象のシートをアクティブにする
    
    row = startRowPoint ' 開始行
    
    For Each item In SheetItems ' 取得したデータ分ループする
        col = startColPoint ' 開始列初期化
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
' サンプルメソッド 取得したデータをループで取得する
'**************************************************
Public Sub SampleGetData()
    Dim item               ' arraylistループ用
    Dim itemCnt As Double ' データ出力列カウント用
    
    For Each item In Me.ToArray ' 取得したデータ分ループする
        For itemCnt = 0 To 2
            Debug.Print item(itemCnt)
        Next
    Next
End Sub

