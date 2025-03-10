Dim configDoc, configFilePath, accessPW
configFilePath = defaultPath & "\config.xml"
Set configDoc = CreateObject("Microsoft.XMLDOM")
configDoc.async = "false"
If fso.FileExists(configFilePath) Then
    configDoc.load(configFilePath)
    accessPW = ""
    Dim securityNode
    Set securityNode = configDoc.selectSingleNode("//security")
    If Not securityNode Is Nothing Then
        Dim accessNode
        Set accessNode = securityNode.selectSingleNode("accesspassword")
        If Not accessNode Is Nothing Then
            accessPW = accessNode.text
        End If
    End If
End If

If accessPW <> "" Then
    ' Determine session file path using nhttp-session-id (folder auto‑managed by server)
    Dim sessionFilePath, sessionDoc, signedIn, nhttpSessionId
    Set nhttpSessionId = xmlDoc.selectSingleNode("//nhttp-session-id")
    If nhttpSessionId Is Nothing Then
        nhttpSessionId = "default"
    Else
        nhttpSessionId = nhttpSessionId.text
    End If
    sessionFilePath = "Sessions\" & nhttpSessionId & "\session.xml"

    ' If session file doesn't exist, recreate it marking user as not signed in
    If Not fso.FileExists(sessionFilePath) Then
        Dim sessionContent
        sessionContent = "<session><signedin>false</signedin></session>"
        Set sessionFile = fso.CreateTextFile(sessionFilePath, True)
        sessionFile.WriteLine(sessionContent)
        sessionFile.Close
        signedIn = "false"
    Else
        Set sessionDoc = CreateObject("Microsoft.XMLDOM")
        sessionDoc.async = "false"
        sessionDoc.load(sessionFilePath)
        Dim signedinNode
        Set signedinNode = sessionDoc.selectSingleNode("//signedin")
        If Not signedinNode Is Nothing Then
            signedIn = signedinNode.text
        Else
            signedIn = "false"
        End If
    End If
    
    ' If not signed in and action is not checkpw, prompt for password.
    If LCase(actionNode.text) <> "checkpw" And signedIn <> "true" Then
        WScript.Echo "<html><head><style>"
        WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; text-align: center; padding-top: 50px; }"
        WScript.Echo "input { padding: 10px; font-size: 16px; }"
        WScript.Echo "button { padding: 10px 20px; font-size: 16px; background-color: #4caf50; color: #ffffff; border: none; cursor: pointer; margin-top: 10px; }"
        WScript.Echo "</style></head><body>"
        WScript.Echo "<h2>Please Enter Password</h2>"
        WScript.Echo "<form method='get' action='explorer.vbs'>"
        WScript.Echo "<input type='hidden' name='action' value='checkpw'>"
        WScript.Echo "<input type='hidden' name='dir' value='" & dirPath & "'>"
        WScript.Echo "<input type='password' name='pw' required><br>"
        WScript.Echo "<button type='submit'>Submit</button>"
        WScript.Echo "</form>"
        WScript.Echo "</body></html>"
        WScript.Quit
    End If
End If