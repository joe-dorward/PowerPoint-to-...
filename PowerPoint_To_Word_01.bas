' PowerPoint_To_Word_01.bas
' GLOBAL VARIABLES
Dim AbsoluteDocumentPath As String
' --------------- --------------- --------------- --------------- ---------------
Function Get_AbsoluteDocumentPath() As String
    Get_AbsoluteDocumentPath = Application.ActiveDocument.FullName
End Function
Sub Main()
    AbsoluteDocumentPath = Get_AbsoluteDocumentPath()
    MsgBox AbsoluteDocumentPath, , "Get_AbsoluteDocumentPath()"
End Sub
