Case "deletedir"
        Set nameNode = xmlDoc.selectSingleNode("//name")
        If nameNode Is Nothing Then
            WScript.Echo "No <name> node found in the XML for directory name."
            WScript.Quit
        End If

        Dim dirToDelete
        dirToDelete = defaultPath & "\files\" & dirPath & "\" & nameNode.text
        If fso.FolderExists(dirToDelete) Then
            fso.DeleteFolder dirToDelete
            WScript.Echo "Directory deleted: " & dirToDelete & "<script>window.location.href='explorer.vbs?action=view&dir=" & dirPath & "';</script>"
        Else
            WScript.Echo "Directory does not exist: " & dirToDelete
        End If