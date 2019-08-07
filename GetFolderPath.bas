Function GetFolderPath(ByVal FolderPath As String) As Object

    Dim oFolder As Object
    Dim olApp As Object
    Dim FoldersArray As Variant
    Dim i As Integer
         
    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    Application.DisplayAlerts = False
    FoldersArray = Split(FolderPath, "\")

    Set olApp = CreateObject("Outlook.Application")
    Set oFolder = olApp.session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Object
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Application.DisplayAlerts = True

    Exit Function
         
GetFolderPath_Error:
    'MsgBox Err.Description
    'Stop
    'Resume
    
    Set GetFolderPath = Nothing
    Exit Function
    
End Function
