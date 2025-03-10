Case "checkpw"
        Dim pwNode, inputPW
        Set pwNode = xmlDoc.selectSingleNode("//pw")
        If pwNode Is Nothing Then
            WScript.Echo "No <pw> node found in the XML."
            WScript.Quit
        End If
        inputPW = pwNode.text
        If inputPW = accessPW Then
            ' Update session file marking signed in (folder auto‑managed)
            Dim sessionContent2
            sessionContent2 = "<session><signedin>true</signedin></session>"
            Set sessionFile = fso.CreateTextFile(sessionFilePath, True)
            sessionFile.WriteLine(sessionContent2)
            sessionFile.Close
            WScript.Echo "Password accepted.<script>window.location.href='explorer.vbs?action=view&dir=" & dirPath & "';</script>"
        Else
            WScript.Echo "<html><head><style>"
            WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; text-align: center; padding-top: 50px; }"
            WScript.Echo "button { padding: 10px 20px; font-size: 16px; background-color: #f44336; color: #ffffff; border: none; cursor: pointer; margin-top: 10px; }"
            WScript.Echo "</style></head><body>"
            WScript.Echo "<h2>Incorrect Password</h2>"
            WScript.Echo "<button onclick='window.history.back()'>Try Again</button>"
            WScript.Echo "</body></html>"
        End If
        WScript.Quit