Case "deletefile"
        Set nameNode = xmlDoc.selectSingleNode("//name")
        If nameNode Is Nothing Then
            WScript.Echo "No <name> node found in the XML for file name."
            WScript.Quit
        End If

        Dim fileToDelete
        fileToDelete = defaultPath & "\files\" & dirPath & "\" & nameNode.text
        If fso.FileExists(fileToDelete) Then
            fso.DeleteFile fileToDelete
            WScript.Echo "File deleted: " & fileToDelete & "<script>window.location.href='explorer.vbs?action=view&dir=" & dirPath & "';</script>"
        Else
            WScript.Echo "File does not exist: " & fileToDelete
        End If
