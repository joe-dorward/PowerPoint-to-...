' PowerPoint_To_Word_03.bas
' GLOBAL VARIABLES
Dim AbsoluteDocumentPath As String
Dim DocumentFilename As String
Dim WorkingFolderPath As String
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
Sub Main()
    AbsoluteDocumentPath = Get_AbsoluteDocumentPath()
    DocumentFilename = Get_DocumentFilename()
    WorkingFolderPath = Get_WorkingFolderPath()
    MsgBox WorkingFolderPath, , "Get_WorkingFolderPath()"
End Sub
