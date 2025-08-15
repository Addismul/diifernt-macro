Public Sub ExtractTrackedChangesToNewDoc()

    '=========================
    'Macro created 2007 by Lene Fredborg, DocTools - www.thedoctools.com
    'THIS MACRO IS COPYRIGHT. YOU ARE WELCOME TO USE THE MACRO BUT YOU MUST KEEP THE LINE ABOVE.
    'YOU ARE NOT ALLOWED TO PUBLISH THE MACRO AS YOUR OWN, IN WHOLE OR IN PART.
    '=========================
    'The macro creates a new document
    'and extracts insertions and deletions
    'marked as tracked changes from the active document
    'NOTE: Other types of changes are skipped
    '(e.g. formatting changes or inserted/deleted footnotes and endnotes)
    'Only insertions and deletions in the main body of the document will be extracted
    'The document will also include metadata
    'Inserted text will be applied black font color
    'Deleted text will be applied red font color
    
    'Minor adjustments are made to the styles used
    'You may need to change the style settings and table layout to fit your needs
    '=========================

    Dim oDoc As Document
    Dim oNewDoc As Document
    Dim oTable As Table
    Dim oRow As Row
    Dim oCol As Column
    Dim oRange As Range
    Dim oRevision As Revision
    Dim strText As String
    Dim n As Long
    Dim i As Long
    Dim Title As String
    
    Title = "Extract Tracked Changes to New Document"
    n = 0 'use to count extracted changes
    
    Set oDoc = ActiveDocument
    
    If oDoc.Revisions.Count = 0 Then
        MsgBox "The active document contains no tracked changes.", vbOKOnly, Title
        GoTo ExitHere
    Else
        'Stop if user does not click Yes
        If MsgBox("Do  you want to extract tracked changes to a new document?" & vbCr & vbCr & _
                "NOTE: Only insertions and deletions will be included. " & _
                "All other types of changes will be skipped.", _
                vbYesNo + vbQuestion, Title) <> vbYes Then
            GoTo ExitHere
        End If
    End If
        
    Application.ScreenUpdating = False
    'Create a new document for the tracked changes, base on Normal.dot
    Set oNewDoc = Documents.Add
    'Set to landscape
    oNewDoc.PageSetup.Orientation = wdOrientLandscape
    With oNewDoc
        'Make sure any content is deleted
        .Content = ""
        'Set appropriate margins
        With .PageSetup
            .LeftMargin = CentimetersToPoints(2)
            .RightMargin = CentimetersToPoints(2)
            .TopMargin = CentimetersToPoints(2.5)
        End With
        'Insert a 6-column table for the tracked changes and metadata
        Set oTable = .Tables.Add _
            (Range:=Selection.Range, _
            numrows:=1, _
            NumColumns:=6)
    End With
    
    'Insert info in header - change date format as you wish
    oNewDoc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = _
        "Tracked changes extracted from: " & oDoc.FullName & vbCr & _
        "Created by: " & Application.UserName & vbCr & _
        "Creation date: " & Format(Date, "MMMM d, yyyy")
            
    'Adjust the Normal style and Header style
    With oNewDoc.Styles(wdStyleNormal)
        With .Font
            .Name = "Arial"
            .Size = 9
            .Bold = False
        End With
        With .ParagraphFormat
            .LeftIndent = 0
            .SpaceAfter = 6
        End With
    End With
    
    With oNewDoc.Styles(wdStyleHeader)
        .Font.Size = 8
        .ParagraphFormat.SpaceAfter = 0
    End With
    
    'Format the table appropriately
    With oTable
        .Range.Style = wdStyleNormal
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        For Each oCol In .Columns
            oCol.PreferredWidthType = wdPreferredWidthPercent
        Next oCol
        .Columns(1).PreferredWidth = 5  'Page
        .Columns(2).PreferredWidth = 5  'Line
        .Columns(3).PreferredWidth = 10 'Type of change
        .Columns(4).PreferredWidth = 55 'Inserted/deleted text
        .Columns(5).PreferredWidth = 15 'Author
        .Columns(6).PreferredWidth = 10 'Revision date
    End With

    'Insert table headings
    With oTable.Rows(1)
        .Cells(1).Range.Text = "Page"
        .Cells(2).Range.Text = "Line"
        .Cells(3).Range.Text = "Type"
        .Cells(4).Range.Text = "What has been inserted or deleted"
        .Cells(5).Range.Text = "Author"
        .Cells(6).Range.Text = "Date"
    End With
    
    'Get info from each tracked change (insertion/deletion) from oDoc and insert in table
    For Each oRevision In oDoc.Revisions
        Select Case oRevision.Type
            'Only include insertions and deletions
            Case wdRevisionInsert, wdRevisionDelete
                'In case of footnote/endnote references (appear as Chr(2)),
                'insert "[footnote reference]"/"[endnote reference]"
                With oRevision
                    'Get the changed text
                    strText = .Range.Text
                
                    Set oRange = .Range
                    Do While InStr(1, oRange.Text, Chr(2)) > 0
                        'Find each Chr(2) in strText and replace by appropriate text
                        i = InStr(1, strText, Chr(2))
                        
                        If oRange.Footnotes.Count = 1 Then
                            strText = Replace(Expression:=strText, _
                                    Find:=Chr(2), Replace:="[footnote reference]", _
                                    Start:=1, Count:=1)
                            'To keep track of replace, adjust oRange to start after i
                            oRange.Start = oRange.Start + i
                    
                        ElseIf oRange.Endnotes.Count = 1 Then
                            strText = Replace(Expression:=strText, _
                                    Find:=Chr(2), Replace:="[endnote reference]", _
                                    Start:=1, Count:=1)
                            'To keep track of replace, adjust oRange to start after i
                            oRange.Start = oRange.Start + i
                        End If
                   Loop
                End With
                'Add 1 to counter
                n = n + 1
                'Add row to table
                Set oRow = oTable.Rows.Add
                
                'Insert data in cells in oRow
                With oRow
                    'Page number
                    .Cells(1).Range.Text = _
                        oRevision.Range.Information(wdActiveEndPageNumber)
                    
                    'Line number - start of revision
                    .Cells(2).Range.Text = _
                        oRevision.Range.Information(wdFirstCharacterLineNumber)
                    
                    'Type of revision
                    If oRevision.Type = wdRevisionInsert Then
                        .Cells(3).Range.Text = "Inserted"
                        'Apply automatic color (black on white)
                        oRow.Range.Font.Color = wdColorAutomatic
                    Else
                        .Cells(3).Range.Text = "Deleted"
                        'Apply red color
                        oRow.Range.Font.Color = wdColorRed
                    End If
                    
                    'The inserted/deleted text
                    .Cells(4).Range.Text = strText
                    
                    'The author
                    .Cells(5).Range.Text = oRevision.Author
                    
                    'The revision date
                    .Cells(6).Range.Text = Format(oRevision.Date, "mm-dd-yyyy")
                End With
        End Select
    Next oRevision
    
    'If no insertions/deletions were found, show message and close oNewDoc
    If n = 0 Then
        MsgBox "No insertions or deletions were found.", vbOKOnly, Title
        oNewDoc.Close savechanges:=wdDoNotSaveChanges
        GoTo ExitHere
    End If
    
    'Apply bold formatting and heading format to row 1
    With oTable.Rows(1)
        .Range.Font.Bold = True
        .HeadingFormat = True
    End With
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
        
    oNewDoc.Activate
    MsgBox n & " tracked changed have been extracted. " & _
        "Finished creating document.", vbOKOnly, Title

ExitHere:
    Set oDoc = Nothing
    Set oNewDoc = Nothing
    Set oTable = Nothing
    Set oRow = Nothing
    Set oRange = Nothing
    
End Sub

