' PowerPoint_To_Word_02_16.bas
' GLOBAL VARIABLES
Dim AbsoluteDocumentPath As String
Dim DocumentFilename As String
Dim WorkingFolderPath As String
Dim PresentationFilename As String
Dim PresentationName As String
Dim SlideCount As Integer
' --------------- --------------- --------------- --------------- ---------------
Function Get_AbsoluteDocumentPath() As String
    Get_AbsoluteDocumentPath = Application.ActiveDocument.FullName
End Function
Function Get_DocumentFilename()
    Get_DocumentFilename = Application.ActiveDocument.name
End Function
Function Get_WorkingFolderPath()
    Dim WorkingFolderPathLength As Integer

    WorkingFolderPathLength = Len(AbsoluteDocumentPath) - Len(DocumentFilename)
    Get_WorkingFolderPath = Left(AbsoluteDocumentPath, WorkingFolderPathLength)
End Function
Sub Open_ThePresentation()
    Set ThePresentation = New PowerPoint.Application
    
    ' create file dialog
    Dim TheFileDialog As FileDialog
    Set TheFileDialog = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
    
    ' configure file dialog
    TheFileDialog.InitialFileName = WorkingFolderPath
    TheFileDialog.Title = "Select PowerPoint Presentation"
    TheFileDialog.Filters.Add "PowerPoint Presentations", "*.pptx", 1
    
    ' open file dialog
    TheFileDialog.Show

    ' open selected presentation
    ThePresentation.Presentations.Open TheFileDialog.SelectedItems.Item(1)
End Sub
Function Get_PresentationFilename()
    Set ThePresentation = New PowerPoint.Application
    Get_PresentationFilename = ThePresentation.ActivePresentation.name
End Function
Function Get_PresentationName()
    Set ThePresentation = New PowerPoint.Application
    Dim DotPosition As Integer

    DotPosition = InStr(PresentationFilename, ".")
    Get_PresentationName = Left(PresentationFilename, DotPosition - 1)
End Function
Sub Add_DocumentBoilerplate()
    ' add document title
    ActiveDocument.Content.InsertAfter PresentationName & vbLf & vbLf
    
    ' format document title
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.Font.Size = 12
    Selection.Font.Bold = True
    Selection.Collapse
      
    ' add note to translator
    ActiveDocument.Content.InsertAfter "Note to translator: You should only translate the green text." & vbLf
End Sub
Function Get_SlideCount()
    Set ThePresentation = New PowerPoint.Application
    Get_SlideCount = ThePresentation.ActivePresentation.Slides.Count
End Function
Sub Test_Add_Table()
    Call Add_Table(1, 3, 5)
End Sub
Sub Add_Table(SlideNumber As Integer, ShapesCount As Integer, ColumnCount As Integer)
    ' add slide sub-title
    ActiveDocument.Content.InsertAfter "Slide " & SlideNumber
    Selection.EndKey Unit:=wdStory
        
    ' add table
    ActiveDocument.Tables.Add _
        Range:=Selection.Range, _
        NumRows:=ShapesCount, _
        NumColumns:=ColumnCount, _
        DefaultTableBehavior:=wdWord9TableBehavior, _
        AutoFitBehavior:=wdAutoFitFixed
End Sub
Sub Test_Format_Table()
    Call Format_Table(1)
End Sub
Sub Format_Table(TableNumber As Integer)
    ' select table
    ActiveDocument.Tables(TableNumber).Select
    Selection.Font.Size = 9

    ' format table paragraphs
    With Selection.ParagraphFormat
        .LeftIndent = 0
        .RightIndent = 1
        .SpaceBefore = 3
        .SpaceAfter = 3
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 11
        .Alignment = wdAlignParagraphLeft
    End With
       
    ' size column widths
    With ActiveDocument.Tables(TableNumber)
        .PreferredWidthType = wdPreferredWidthPoints
        .Columns(1).SetWidth ColumnWidth:=20, RulerStyle:=wdAdjustFirstColumn
        .Columns(2).SetWidth ColumnWidth:=140, RulerStyle:=wdAdjustFirstColumn
        .Columns(3).SetWidth ColumnWidth:=90, RulerStyle:=wdAdjustFirstColumn
        .Columns(4).SetWidth ColumnWidth:=110, RulerStyle:=wdAdjustFirstColumn
    End With
        
    ' make 5th column text green
    ActiveDocument.Tables(TableNumber).Columns(5).Select
    Selection.Font.TextColor.RGB = RGB(110, 175, 70)
    Selection.Collapse
End Sub
' --------------- --------------- --------------- --------------- ---------------
Sub Test_Get_ShapeCount()
    MsgBox Get_ShapeCount(1), , "Get_ShapeCount()"
End Sub
Function Get_ShapeCount(SlideNumber As Integer) As Integer
    Set ThePresentation = New PowerPoint.Application
    Get_ShapeCount = ThePresentation.ActivePresentation.Slides(SlideNumber).Shapes.Count
End Function
' --------------- --------------- --------------- --------------- ---------------
Sub Test_Get_ShapeName()
    MsgBox Get_ShapeName(1, 1), , "Get_ShapeName()"
End Sub
Function Get_ShapeName(SlideNumber As Integer, ShapeNumber As Integer) As String
    Set ThePresentation = New PowerPoint.Application
    Get_ShapeName = ThePresentation.ActivePresentation.Slides(SlideNumber).Shapes(ShapeNumber).name
End Function
' --------------- --------------- --------------- --------------- ---------------
Sub Test_Get_ShapeType()
    MsgBox Get_ShapeType(2, 1), , "Get_ShapeType()"
End Sub
Function Get_ShapeType(SlideNumber As Integer, ShapeNumber As Integer) As String
    Set ThePresentation = New PowerPoint.Application
    Dim ShapeType As MsoShapeType
    ShapeType = ThePresentation.ActivePresentation.Slides(SlideNumber).Shapes(ShapeNumber).Type

    Select Case ShapeType
        Case msoPlaceholder
            Get_ShapeType = "msoPlaceholder"
        Case msoTextBox
            Get_ShapeType = "msoTextBox"
        Case msoAutoShape
            Get_ShapeType = "msoAutoShape"
        Case msoPicture
            Get_ShapeType = "msoPicture"
        Case msoGraphic
            Get_ShapeType = "msoGraphic"
        Case msoGroup
            Get_ShapeType = "msoGroup"
        Case Else
            Get_ShapeType = "[UNKNOWN SHAPE TYPE]"
    End Select

End Function
' --------------- --------------- --------------- --------------- ---------------
Sub Test_Has_Text_Frame()
    MsgBox Has_Text_Frame(3, 1), , "Has_Text_Frame()"
    MsgBox Has_Text_Frame(3, 3), , "Has_Text_Frame()"
End Sub
Function Has_Text_Frame(SlideNumber As Integer, ShapeNumber As Integer) As Boolean
    Set ThePresentation = New PowerPoint.Application
    Has_Text_Frame = ThePresentation.ActivePresentation.Slides(SlideNumber).Shapes(ShapeNumber).HasTextFrame
End Function
' --------------- --------------- --------------- --------------- ---------------
Sub Test_Get_ShapeText()
    MsgBox Get_ShapeText(1, 1), , "Get_ShapeText()"
End Sub
Function Get_ShapeText(SlideNumber As Integer, ShapeNumber As Integer) As String
    Set ThePresentation = New PowerPoint.Application
    Get_ShapeText = ThePresentation.ActivePresentation.Slides(SlideNumber).Shapes(ShapeNumber).TextFrame.TextRange.Text
End Function
' --------------- --------------- --------------- --------------- ---------------
Sub Add_Tables()
    
    Dim SlideNumber As Integer
    Dim ShapeNumber As Integer
    Dim ShapeCount As Integer
    Dim ShapeText As String
        
    For SlideNumber = 1 To 2 'SlideCount
    
        ShapeCount = Get_ShapeCount(SlideNumber)
        Call Add_Table(SlideNumber, ShapeCount, 5)
        Call Format_Table(SlideNumber)
               
        ' add shape information to table
        For ShapeNumber = 1 To ShapeCount
        
            With ActiveDocument.Tables(SlideNumber)
                ' shape-number to column-one
                .Cell(ShapeNumber, 1).Range.InsertAfter ShapeNumber
                
                ' shape-name to column-two
                .Cell(ShapeNumber, 2).Range.InsertAfter Get_ShapeName(SlideNumber, ShapeNumber)
                
                ' shape-type to column-three
                .Cell(ShapeNumber, 3).Range.InsertAfter Get_ShapeType(SlideNumber, ShapeNumber)

                ' columns 4 & 5 - shape text
                ' test for non-text shape
                If Has_Text_Frame(SlideNumber, ShapeNumber) Then
                
                    ' get shape-text
                    ShapeText = Get_ShapeText(SlideNumber, ShapeNumber)
            
                    If Len(ShapeText) > 0 Then
                        ' shape-text to column-five
                        ActiveDocument.Tables(SlideNumber).Cell(ShapeNumber, 5).Range.InsertAfter ShapeText
                    Else
                        ' message to column-four
                        ActiveDocument.Tables(SlideNumber).Cell(ShapeNumber, 4).Range.InsertAfter "[NO TEXT]"
                    End If
                    
                Else
                    ' message to column-four
                    ActiveDocument.Tables(SlideNumber).Cell(ShapeNumber, 4).Range.InsertAfter "[NON-TEXT SHAPE]"
                End If
                
            End With
            
        Next ShapeNumber
        
    Next SlideNumber

End Sub
Sub Main()
    AbsoluteDocumentPath = Get_AbsoluteDocumentPath()
    DocumentFilename = Get_DocumentFilename()
    WorkingFolderPath = Get_WorkingFolderPath()
    
    ' open the presentation
    Call Open_ThePresentation

    PresentationFilename = Get_PresentationFilename()
    PresentationName = Get_PresentationName()
    Call Add_DocumentBoilerplate
    SlideCount = Get_SlideCount()
    Call Add_Tables
End Sub
