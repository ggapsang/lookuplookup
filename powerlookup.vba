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
        MsgBox "No lookup values selected. Please select at least one lookup value."
    ElseIf txtHeaderRowSource.Value = "" Or txtHeaderRowTarget.Value = "" Or txtKeyValue.Value = "" Then
        MsgBox "Please fill in all the header row numbers and key value fields."
    Else
        UpdateTargetWorksheet selectedItems, CLng(txtHeaderRowSource.Value), CLng(txtHeaderRowTarget.Value), txtKeyValue.Value
        Unload Me
    End If

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
    tgtHeaderRow = InputBox("Enter the header row number_target : ")

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

Sub RunUpdateTargetWorksheet()
    ' Show the UserForm to select multiple lookup values
    UserForm1.Show
End Sub

Sub UpdateTargetWorksheet(selectedItems As Collection, srcHeaderRow As Long, tgtHeaderRow As Long, keyValueText As String)
    
    ' Turn off screen updating and calculation
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

    ' Set the source range and workbook
    Set srcWb = ActiveWorkbook
    Set srcWs = ActiveSheet

    ' Prompt the user to select the target workbook
    tgtFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", Title:="Select the target workbook", MultiSelect:=False)
    If tgtFile = False Then
        MsgBox "No target workbook selected. Exiting the macro."
        Exit Sub
    End If

    ' Check if the target workbook is already open
    Set tgtWb = GetWorkbook(tgtFile)

    ' If the workbook is not open, open it
    If tgtWb Is Nothing Then
        Set tgtWb = Workbooks.Open(tgtFile)
    End If

    ' Set the target worksheet
    Set tgtWs = tgtWb.ActiveSheet

    startTime = Timer

    ' Find the column numbers for the key value in the source and target header rows
    srcKeyValueCol = Application.Match(keyValueText, srcWs.Rows(srcHeaderRow), 0)
    tgtKeyValueCol = Application.Match(keyValueText, tgtWs.Rows(tgtHeaderRow), 0)
    ReDim srcLookupValueCol(selectedItems.Count - 1)
    ReDim tgtLookupValueCol(selectedItems.Count - 1)

    For j = 1 To selectedItems.Count
        srcLookupValueCol(j - 1) = Application.Match(selectedItems.item(j), srcWs.Rows(srcHeaderRow), 0)
        tgtLookupValueCol(j - 1) = Application.Match(selectedItems.item(j), tgtWs.Rows(tgtHeaderRow), 0)
    Next j

    If IsError(srcKeyValueCol) Or IsError(tgtKeyValueCol) Then
        MsgBox "The specified key value text was not found in the header rows. Exiting the macro."
        Exit Sub
    End If


    ' Determine the last row with data in the target worksheet
    lastRow = tgtWs.Cells(tgtWs.Rows.Count, tgtKeyValueCol).End(xlUp).Row

    ' Create a Collection to store the key values and their corresponding row numbers from the target worksheet
    Set tgtKeyCells = New Collection
    On Error Resume Next
    For i = tgtHeaderRow + 1 To lastRow
        tgtKeyCells.Add i, CStr(tgtWs.Cells(i, tgtKeyValueCol).Value)
    Next i
    On Error GoTo 0

    ' Loop through the source range
    For Each srcKeyCell In srcWs.Range(srcWs.Cells(srcHeaderRow + 1, srcKeyValueCol), srcWs.Cells(srcWs.Rows.Count, srcKeyValueCol).End(xlUp))
        ' Check if the source row is visible (not filtered out)
        If srcWs.Rows(srcKeyCell.Row).Hidden = False Then
            ' Find the corresponding row in the target worksheet based on the key value using the collection
            On Error Resume Next
            rowNumber = tgtKeyCells(CStr(srcKeyCell.Value))
            On Error GoTo 0
            If Not IsEmpty(rowNumber) Then
                ' Check if the target row is visible (not filtered out)
                If tgtWs.Rows(rowNumber).Hidden = False Then
                    ' Compare the lookup values in the source and target worksheets for each specified lookup value
                    For j = LBound(srcLookupValueCol) To UBound(srcLookupValueCol)
                        If tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).Value <> srcWs.Cells(srcKeyCell.Row, srcLookupValueCol(j)).Value Then
                            tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).Value = srcWs.Cells(srcKeyCell.Row, srcLookupValueCol(j)).Value
                            tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).Interior.Color = RGB(255, 165, 0) ' Set the cell background color to orange
                        End If
                    Next j
                End If
            End If
        End If
    Next srcKeyCell

    ' Turn screen updating and calculation back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

    ' Check process run time
    finishTime = Timer - startTime

    ' Save the target workbook without closing it
    tgtWb.Save
    MsgBox "Target worksheet updated successfully. " & Format(Int(finishTime / 60), "0") & " min " & Format(finishTime Mod 60, "0.00") & " sec"

End Sub
