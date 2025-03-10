Case "createdir"
        Set nameNode = xmlDoc.selectSingleNode("//name")
        If nameNode Is Nothing Then
            WScript.Echo "No <name> node found in the XML for directory name."
            WScript.Quit
        End If
        
        Dim newDirPath
        newDirPath = defaultPath & "\files\" & dirPath & "\" & nameNode.text
        If Not fso.FolderExists(newDirPath) Then
            fso.CreateFolder(newDirPath)
            WScript.Echo "Directory created: " & newDirPath & "<script>window.location.href='explorer.vbs?action=view&dir=" & dirPath & "/" & nameNode.text & "';</script>"
        Else
            WScript.Echo "<html><head><style>"
            WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; }"
            WScript.Echo ".error-box { background-color: #2e2e2e; color: #e57373; padding: 20px; border: 1px solid #444444; margin-top: 20px; }"
            WScript.Echo "button { padding: 10px 15px; margin-top: 10px; background-color: #007bff; color: #ffffff; border: none; cursor: pointer; }"
            WScript.Echo "button:hover { background-color: #0056b3; }"
            WScript.Echo "</style></head><body>"
            WScript.Echo "<div class='error-box'><h2>Directory already exists</h2>"
            WScript.Echo "<p>The directory you are trying to create already exists: " & newDirPath & "</p>"
            WScript.Echo "<button onclick='window.history.back()'>Go Back</button></div>"
            WScript.Echo "</body></html>"
        End If
