Attribute VB_Name = "ExportAllVBA"
Sub ExportAllVBA()
    Dim cmp As Object
    Dim sPath As String
    sPath = ThisWorkbook.path & "\vba_src\"
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath

    For Each cmp In ThisWorkbook.VBProject.VBComponents
        Dim ext As String
        Select Case cmp.Type
            Case 1: ext = ".bas"   ' standard module
            Case 2: ext = ".cls"   ' class module
            Case 3: ext = ".frm"   ' userform
            Case 100: ext = ".cls" ' document module (ThisWorkbook / sheets)
            Case Else: ext = ".txt"
        End Select
        cmp.Export sPath & cmp.Name & ext
    Next cmp

    MsgBox "Exported " & ThisWorkbook.VBProject.VBComponents.count & _
           " components to " & sPath
End Sub

