'Split a tif file, each page to a new tif.

Sub tif_split(byVal source As String, byVal newName As String)
    Dim doc As New MODI.Document
    doc.Creat source
    n = doc.Images.Count
    For i = 0 to n - 1
        doc.PrintOut i, i, 1, "Microsoft Office Document Image Writer", newName & i & ".tif", FitMode = 0
    Next i
    Set doc = Nothing
End sub
