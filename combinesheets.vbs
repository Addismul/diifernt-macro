Sub CopySheets()
    Application.ScreenUpdating = False
    Dim fd As FileDialog, lRow As Long, vSelectedItem As Variant, srcWB As Workbook, desWB As Workbook
    Set desWB = ThisWorkbook
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = True
        If .Show = -1 Then
            For Each vSelectedItem In .SelectedItems
                Set srcWB = Workbooks.Open(vSelectedItem)
                Sheets(1).Copy after:=desWB.Sheets(desWB.Sheets.Count)
                srcWB.Close False
            Next
            Else
        End If
    End With
    Application.ScreenUpdating = True
End Sub
