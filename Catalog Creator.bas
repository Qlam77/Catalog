Attribute VB_Name = "Module1"
Private Sub HeaderCreater(textMsg As String)
    'Creates a header with heading 1 style then moves cursor to a new line with normal style
    Selection.Style = wdStyleHeading1
    Selection.InsertBefore (textMsg & vbNewLine)
    Selection.EndOf
    Selection.Style = wdStyleNormal
End Sub

Sub InsertProductPhotos()
    Dim fd As FileDialog
    Dim oTable As Table
    Dim iRow As Integer
    Dim iCol As Integer
    Dim oCell As Range
    Dim y As Long
    Dim i As Long
    Dim sNoDoc As String
    Dim picName As String
    Dim scale_Factor As Long
    Dim max_height As Single
    Dim rng As Word.Range
    
    'define resize constraints
    max_height = 174  ' 2 inches = 144 pt
    max_columns = 3
    numTables = 0
    
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select image files and click OK"
        .Filters.Add "Images", "*.gif; *.jpg; *.jpeg; *.bmp; *.tif; *.png; *.wmf"
        .FilterIndex = 2
        
        If .Show = -1 Then
            'add a 1 row 3 column table to take the images
            Set oTable = Selection.Tables.Add(Selection.Range, 1, max_columns)
            '+++++++++++++++++++++++++++++++++++++++++++++
            'oTable.AutoFitBehavior (wdAutoFitFixed)
            oTable.Rows.Height = InchesToPoints(2.1)
            oTable.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
            '++++++++++++++++++++++++++++++++++++++++++++++
            oTable.Borders.Enable = True
            '++++++++++++++++++++++++++++++++++++++++++++++
            'starting column location
            iCol = 1
            iRow = 1
            numTables = ActiveDocument.Tables.Count
            
            For i = 1 To .SelectedItems.Count
                picName = WordBasic.FilenameInfo(.SelectedItems(i), 3)
                
                'select cell
                'Set oCell = ActiveDocument.Tables(numTables).Cell(iRow, iCol).Range
                Set oCell = Selection.Tables(1).Cell(iRow, iCol).Range
                
                'insert image
                oCell.InlineShapes.AddPicture FileName:= _
                .SelectedItems(i), LinkToFile:=False, _
                SaveWithDocument:=True, Range:=oCell
                
                'resize image
                If oCell.InlineShapes(1).Height > max_height Then
                    scale_Factor = oCell.InlineShapes(1).ScaleHeight * (max_height / oCell.InlineShapes(1).Height)
                    oCell.InlineShapes(1).ScaleHeight = scale_Factor
                    oCell.InlineShapes(1).ScaleWidth = scale_Factor
                End If
                
                'center content
                oCell.ParagraphFormat.Alignment = wdAlignParagraphCenter
                
                If i < .SelectedItems.Count Then  'add another row, more to go
                    iCol = iCol + 1
                    If iCol > max_columns Then
                        oTable.Rows.Add
                        iRow = iRow + 1
                        iCol = 1
                    End If
                End If
            Next i
            UpdateTOC
        Else
            End
        End If
    End With
    Set fd = Nothing
End Sub

Private Sub UpdateTOC()
    'updates the table of contents
    Dim TOC As TableOfContents
    For Each TOC In ActiveDocument.TablesOfContents
        TOC.Update
    Next TOC
End Sub

Private Sub EndCreater()
    'Creates the end portion of the section if the user wants to add more
    Dim rng As Range
    Dim numTables As Long
    numTables = 0
    numTables = ActiveDocument.Tables.Count
    Set rng = ActiveDocument.Tables(numTables).Range
    If numTables > 1 Then
        Selection.Tables(1).Select
        Selection.Collapse WdCollapseDirection.wdCollapseEnd
        Selection.InsertBreak Type:=wdPageBreak
        Selection.MoveLeft Unit:=wdCharacter, Count:=3
        Selection.Delete Unit:=wdCharacter, Count:=1
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.Delete Unit:=wdCharacter, Count:=1
        Selection.EndOf
    End If
End Sub

Sub CatalogCreater()
    'the main program that will loop through catalog creation and ends when the users stops
    Dim Check As Boolean
    Dim i As String
    Check = False
    Do
        i = InputBox("Enter Catalog Title", "Catalog Creater")
        If Not (i = vbNullString Or StrPtr(i) = 0) Then
                HeaderCreater (i)
                InsertProductPhotos
                Check = (MsgBox("Enter More Products?", vbYesNo) = vbYes) ' Stop when user click's on No
                If Check Then
                    EndCreater
                End If
        Else
            End
        End If
    Loop Until Check = False
End Sub

Sub ClearContents()
    'clears the catalog excluding the title page and toc
    Dim rng As Range
    Dim pge As Integer
    
    'Define the start page
    pge = 3
    
    Selection.GoTo wdGoToPage, wdGoToAbsolute, Count:=pge
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Selection.Delete
    Selection.Style = wdStyleNormal
    UpdateTOC
End Sub





