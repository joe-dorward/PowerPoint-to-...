' PowerPoint_To_Word_02.bas
' GLOBAL VARIABLES
Dim AbsoluteDocumentPath As String
Dim DocumentFilename As String
' --------------- --------------- --------------- --------------- ---------------
Function Get_AbsoluteDocumentPath() As String
    Get_AbsoluteDocumentPath = Application.ActiveDocument.FullName
End Function
Function Get_DocumentFilename()
    Get_DocumentFilename = Application.ActiveDocument.name
End Function
Sub Main()
    AbsoluteDocumentPath = Get_AbsoluteDocumentPath()
    ' MsgBox AbsoluteDocumentPath, , "Get_AbsoluteDocumentPath()"
    
    DocumentFilename = Get_DocumentFilename()
    MsgBox DocumentFilename, , "Get_DocumentFilename()"
End Sub
