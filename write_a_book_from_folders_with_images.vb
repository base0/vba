Sub ListFilesInSecondDepth()
    Dim mainFolder As String
    Dim subFolder As String
    Dim subSubFolder As String
    Dim fso As Object
    Dim mainFolderObj As Object
    Dim subFolderObj As Object
    Dim subSubFolderObj As Object
    Dim fileObj As Object
    
    ' Set the main folder path
    mainFolder = "C:\Users\name\Desktop\folder"
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get the main folder
    Set mainFolderObj = fso.GetFolder(mainFolder)
    
    ' Loop through each subfolder in the main folder
    For Each subFolderObj In mainFolderObj.SubFolders
        ' Print the name of the subfolder
        ' Debug.Print "Subfolder: " & subFolderObj.Path
        Call InsertTextAsHeading(getFilenameInFullPath(subFolderObj.Path), 1)
        
        ' Loop through each sub-subfolder within the subfolder
        For Each subSubFolderObj In subFolderObj.SubFolders
            ' Print the name of the sub-subfolder
            ' Debug.Print "  Sub-subfolder: " & subSubFolderObj.Path
            Call InsertTextAsHeading(getFilenameInFullPath(subSubFolderObj.Path), 2)
            
            ' Loop through each file in the sub-subfolder
            For Each fileObj In subSubFolderObj.Files
                ' Print the name of the file
                ' Debug.Print subSubFolderObj.Path & "\" & fileObj.Name
                 If Right(fileObj.Name, 3) = "png" Then
                    InsertImageIntoDocument (subSubFolderObj.Path & "\" & fileObj.Name)
                End If
            Next fileObj
        Next subSubFolderObj
    Next subFolderObj
    
    ' Clean up
    Set fileObj = Nothing
    Set subSubFolderObj = Nothing
    Set subFolderObj = Nothing
    Set mainFolderObj = Nothing
    Set fso = Nothing
    
End Sub

Function getFilenameInFullPath(fullpath As String)
    lastDelimiter = InStrRev(fullpath, "\")
    getFilenameInFullPath = Mid(fullpath, lastDelimiter + 1)
End Function

Sub InsertTextAsHeading(textToInsert, level)
    Dim wordDoc As Document
    Dim wordRange As Range
    
    ' Reference the active document
    Set wordDoc = ActiveDocument
    
    ' Set the range where the text will be inserted (for example, at the end of the document)
    Set wordRange = wordDoc.Content
    wordRange.Collapse Direction:=wdCollapseEnd
    
    ' Insert the text
    wordRange.Text = textToInsert
    
    ' Apply the Heading 1 style
    If level = 1 Then
       wordRange.Style = wordDoc.Styles(wdStyleHeading1)
    Else
       wordRange.Style = wordDoc.Styles(wdStyleHeading2)
    End If
    
    ' Move the cursor to the end of the inserted text
    wordRange.Collapse Direction:=wdCollapseEnd
    
    wordRange.InsertParagraphAfter
End Sub


Sub InsertImageIntoDocument(imgPath)
    Dim wordDoc As Document
    Dim wordRange As Range
    
    ' Reference the active document
    Set wordDoc = ActiveDocument
    
    ' Set the range where the image will be inserted (for example, at the end of the document)
    Set wordRange = wordDoc.Content
    wordRange.Collapse Direction:=wdCollapseEnd
    
    ' Insert the image at the specified range
    wordDoc.InlineShapes.AddPicture FileName:=imgPath, _
        LinkToFile:=False, SaveWithDocument:=True, Range:=wordRange
    
    ' Optionally, you can format the image (e.g., resize it)
    Const SIZE = 220
    With wordDoc.InlineShapes(wordDoc.InlineShapes.Count)
        .LockAspectRatio = msoTrue
        .WIDTH = SIZE ' Set width in points
        .Height = SIZE ' Set height in points
        .Range.InsertAfter vbCr
    End With
    
    wordRange.Collapse Direction:=wdCollapseEnd
End Sub
