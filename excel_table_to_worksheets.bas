    Const COL = "XFD"
    
    mainSheet = ActiveSheet.Name
    tabl = ActiveSheet.ListObjects(1).Name
    Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range(COL & 1), Unique:=True
    
    i = 2
    While Not IsEmpty(Range(COL & i))
        neoSheet = Range(COL & i).Value
    
        Range(tabl).AutoFilter Field:=1, Criteria1:=neoSheet
        Range(tabl & "[#All]").Select
        Selection.Copy
        
        neoSheet = Replace(neoSheet, ".xlsx", "")
        Sheets.Add.Name = neoSheet
        Sheets(neoSheet).Select
        ActiveSheet.Paste
        Columns("A:A").Select
        Application.CutCopyMode = False
        Selection.Delete Shift:=xlToLeft
        
        Sheets(mainSheet).Select
        i = i + 1
    Wend
