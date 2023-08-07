'''''''''''''''UserForm1'''''''''''''''''''

Private Sub cmdSubmit_Click()

    Dim selectedItems As Collection
    Dim item As Variant
    Dim i As Long

    Set selectedItems = New Collection

    For i = 0 To lstLookupValues.ListCount - 1
        If lstLookupValues.Selected(i) Then
            selectedItems.Add lstLookupValues.List(i)
        End If
    Next i

    If selectedItems.Count = 0 Then
        MsgBox "찾을 값은 한개 이상 입력해야 됨."
    ElseIf txtHeaderRowSource.Value = "" Or txtHeaderRowTarget.Value = "" Or txtKeyValue.Value = "" Then
        MsgBox "해더 행 번호와 키 값 다시 확인"
    Else
        UpdateTargetWorksheet selectedItems, CLng(txtHeaderRowSource.Value), CLng(txtHeaderRowTarget.Value), txtKeyValue.Value
        Unload Me
    End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub lstLookupValues_Click()

End Sub

Private Sub txtHeaderRowSource_Change()

End Sub

Private Sub UserForm_Initialize()
    Dim tgtWb As Workbook
    Dim tgtWs As Worksheet
    Dim tgtHeaderRow As Long
    Dim i As Long
    Dim lastCol As Long

    ' Set the target workbook and worksheet
    Set tgtWb = ActiveWorkbook
    Set tgtWs = tgtWb.ActiveSheet
    tgtHeaderRow = InputBox("해더 행 번호 : ")

    lastCol = tgtWs.Cells(tgtHeaderRow, tgtWs.Columns.Count).End(xlToLeft).Column

    For i = 1 To lastCol
        lstLookupValues.AddItem tgtWs.Cells(tgtHeaderRow, i).Value
    Next i
End Sub




'''''''''''''''''''''MODULE1''''''''''''''''''''''''''''''''

Function GetWorkbook(ByVal sFullName As String) As Workbook
    Dim sFile As String
    Dim wbReturn As Workbook

    sFile = Dir(sFullName)

    On Error Resume Next
    Set wbReturn = Workbooks(sFile)

    If wbReturn Is Nothing Then
        Set GetWorkbook = Nothing
    Else
        Set GetWorkbook = wbReturn
    End If

    On Error GoTo 0
End Function
Function MatchKeyValue(keyValue As Variant, headerRow As Range) As Long
Dim cell As Range
    Dim i As Long
    
    For Each cell In headerRow.Cells
        If IsNumeric(keyValue) And IsNumeric(cell.Value) Then
            If CDbl(keyValue) = CDbl(cell.Value) Then
                MatchKeyValue = cell.Column
                Exit Function
            End If
        Else
            If CStr(keyValue) = CStr(cell.Value) Then
                MatchKeyValue = cell.Column
                Exit Function
            End If
        End If
    Next cell

    MatchKeyValue = 0
End Function
Function MatchValue(lookupValue As Variant, headerRow As Range) As Long
    Dim cell As Range

    For Each cell In headerRow.Cells
        If IsNumeric(lookupValue) And IsNumeric(cell.Value) Then
            If CDbl(lookupValue) = CDbl(cell.Value) Then
                MatchValue = cell.Column
                Exit Function
            End If
        Else
            If CStr(lookupValue) = CStr(cell.Value) Then
                MatchValue = cell.Column
                Exit Function
            End If
        End If
    Next cell

    MatchValue = 0
End Function

Sub RunUpdateTargetWorksheet()
    ' 유저폼에서 값을 선택할 수 있게 보여줌
    UserForm1.Show
End Sub


Sub UpdateTargetWorksheet(selectedItems As Collection, srcHeaderRow As Long, tgtHeaderRow As Long, keyValueText As String)
    
    ' 매크로 작업 중 스크린 활성화 정지
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWb As Workbook
    Dim tgtWs As Worksheet
    Dim tgtFile As Variant
    Dim srcKeyValueCol As Long
    Dim srcLookupValueCol() As Long
    Dim tgtKeyValueCol As Long
    Dim tgtLookupValueCol() As Long
    Dim srcKeyCell As Range
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim startTime As Double
    Dim finishTime As Double
    Dim tgtKeyCells As Object
    Dim rowNumber As Variant

    ' source worksheet 세팅
    Set srcWb = ActiveWorkbook
    Set srcWs = ActiveSheet

    ' target workbook 세팅 명령창 실행
    tgtFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", Title:="불러올 파일 선택", MultiSelect:=False)
    If tgtFile = False Then
        MsgBox "불러올 파일이 선택되지 않아서 매크로를 종료합니다"
        Exit Sub
    End If

    ' target workbook이 이미 열려있을 경우
    Set tgtWb = GetWorkbook(tgtFile)

    ' target workbook이 열려 있지 않을 경우 해당 파일을 열음
    If tgtWb Is Nothing Then
        Set tgtWb = Workbooks.Open(tgtFile)
    End If

    ' target worksheet 세팅
    Set tgtWs = tgtWb.ActiveSheet

    startTime = Timer

    ' source worksheet와 target worksheet의 키 값이 해더 행 몇열에 있는지 확인
    srcKeyValueCol = MatchKeyValue(keyValueText, srcWs.Rows(srcHeaderRow))
    tgtKeyValueCol = MatchKeyValue(keyValueText, tgtWs.Rows(srcHeaderRow))
    ReDim srcLookupValueCol(selectedItems.Count - 1)
    ReDim tgtLookupValueCol(selectedItems.Count - 1)

    For j = 1 To selectedItems.Count
        srcLookupValueCol(j - 1) = MatchValue(selectedItems.item(j), srcWs.Rows(srcHeaderRow))
        tgtLookupValueCol(j - 1) = MatchValue(selectedItems.item(j), tgtWs.Rows(tgtHeaderRow))
    Next j

    If IsError(srcKeyValueCol) Or IsError(tgtKeyValueCol) Then
        MsgBox "해더 행에서 키 값을 찾지 못해 매크로를 종료합니다."
        Exit Sub
    End If


    ' targetworksheet의 마지막 행 찾기(주의 - sr no가 다 채워지지 않고 중간에 끊기면 에러가 생길 수 있음)
    lastRow = tgtWs.Cells(tgtWs.Rows.Count, tgtKeyValueCol).End(xlUp).Row

    ' 딕셔너리 생성
    Set tgtKeyCells = New Collection
    On Error Resume Next
    For i = tgtHeaderRow + 1 To lastRow
        tgtKeyCells.Add i, CStr(tgtWs.Cells(i, tgtKeyValueCol).Value)
    Next i
    On Error GoTo 0

    ' loop
    For Each srcKeyCell In srcWs.Range(srcWs.Cells(srcHeaderRow + 1, srcKeyValueCol), srcWs.Cells(srcWs.Rows.Count, srcKeyValueCol).End(xlUp))
        ' 필터값만 보이게 반영함
        If srcWs.Rows(srcKeyCell.Row).Hidden = False Then
            ' 키 값을 기준으로 target worksheet의 값들이 일치하는지를 비교함
            On Error Resume Next
            rowNumber = tgtKeyCells(CStr(srcKeyCell.Value))
            On Error GoTo 0
            If Not IsEmpty(rowNumber) Then
                ' 필터값만 보이게 반영함
                If tgtWs.Rows(rowNumber).Hidden = False Then
                    ' source worksheets와 target worksheets의 값들을 비교하여 변경함
                    For j = LBound(srcLookupValueCol) To UBound(srcLookupValueCol)
                        If tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).Value <> srcWs.Cells(srcKeyCell.Row, srcLookupValueCol(j)).Value Then
                            tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).Value = srcWs.Cells(srcKeyCell.Row, srcLookupValueCol(j)).Value
                            tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).Interior.Color = RGB(255, 165, 0) ' 바뀐 값을 주황색으로 만듬(rgb코드로 원하는 색 설정 가능)
                        End If
                    Next j
                End If
            End If
        End If
    Next srcKeyCell

    ' 매크로 작업 중 스크린 활성화 정지
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

    ' 작업시간 측정
    finishTime = Timer - startTime

    ' 바뀐 값으로 파일을 자동으로 저장함 (현재는 저장 기능 오프 상태 파일이 열려있지 않거나 값을 자동으로 저장하려면 아래 줄 코드에 ' 표시를 빼면 됨
    'tgtWb.Save
    MsgBox "끝 " & Format(Int(finishTime / 60), "0") & " min " & Format(finishTime Mod 60, "0.00") & " sec"

End Sub
