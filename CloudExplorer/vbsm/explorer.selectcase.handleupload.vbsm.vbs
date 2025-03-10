Case "handleupload"
        ' Handle file upload and move it to the specified directory
        Set nameNode = xmlDoc.selectSingleNode("//filename")
        Set sessionIdNode = xmlDoc.selectSingleNode("//nhttp-session-id")
        
        If nameNode Is Nothing Or sessionIdNode Is Nothing Then
            DisplayError "Missing required nodes in the XML (filename or nhttp-session-id)."
            WScript.Quit
        End If
        
        Dim sourceFilePath, destFilePath
        sourceFilePath = "Sessions/" & sessionIdNode.text & "/" & nameNode.text
        destFilePath = defaultPath & "\files\" & dirPath & "\" & nameNode.text
        
        On Error Resume Next
        If fso.FileExists(sourceFilePath) Then
            ' Delete the destination file if it exists
            If fso.FileExists(destFilePath) Then
                fso.DeleteFile destFilePath
                If Err.Number <> 0 Then
                    DisplayError "Error deleting existing file: " & Err.Description
                    WScript.Quit
                End If
            End If

            ' Move the source file to the destination
            fso.MoveFile sourceFilePath, destFilePath
            If Err.Number <> 0 Then
                DisplayError "Error moving file: " & Err.Description
            Else
                WScript.Echo "File moved to: " & destFilePath & "<script>window.location.href='explorer.vbs?action=view&dir=" & dirPath & "';</script>"
            End If
        Else
            DisplayError "Source file does not exist: " & sourceFilePath
        End If
        On Error GoTo 0