Sub ProcessAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        Call stockSorter
    Next ws
End Sub