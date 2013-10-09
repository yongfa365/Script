Function FilesTree(sPath)
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFso.GetFolder(sPath)
    Set oSubFolders = oFolder.SubFolders
    
    Set oFiles = oFolder.Files
    For Each oFile In oFiles
        WScript.Echo oFile.Path
        'oFile.Delete
    Next
    
    For Each oSubFolder In oSubFolders
        WScript.Echo oSubFolder.Path
        'oSubFolder.Delete
        FilesTree(oSubFolder.Path)'µÝ¹é
    Next
    
    Set oFolder = Nothing
    Set oSubFolders = Nothing
    Set oFso = Nothing
End Function

FilesTree("F:\deltest\deltest") '±éÀú
