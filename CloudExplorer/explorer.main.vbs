Dim fso, folder, subfolder, file
Dim rowColor, rowNum
Dim xmlDoc, actionNode, dirNode, nameNode, dirPath, defaultPath
Dim validActions

' Load incoming XML
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.async = "false"
xmlDoc.loadXML WScript.Arguments(0)

Set actionNode = xmlDoc.selectSingleNode("//action")
If actionNode Is Nothing Then
    WScript.Echo "<script>window.location.href='explorer.vbs?action=view&dir=';</script>"
    WScript.Quit
End If

validActions = Array("view", "downloadfiles", "viewimage", "viewvideo", "handleupload", "createdir", "deletefile", "deletedir", "uploadfile", "checkpw")
If Not IsInArray(actionNode.text, validActions) Then
    WScript.Echo "Invalid action specified."
    WScript.Quit
End If

Set dirNode = xmlDoc.selectSingleNode("//dir")
If Not dirNode Is Nothing Then
    dirPath = dirNode.text
Else
    DisplayError "No <dir> node found in the XML."
    WScript.Quit
End If

' Clean up dirPath by removing any instances of "..", "../", or "./"
dirPath = Replace(dirPath, "..\", "")
dirPath = Replace(dirPath, "../", "")
dirPath = Replace(dirPath, "./", "")

Set fso = CreateObject("Scripting.FileSystemObject")
defaultPath = fso.GetParentFolderName(WScript.ScriptFullName)

On Error Resume Next
Set folder = fso.GetFolder(defaultPath & "\files\" & dirPath)
If Err.Number <> 0 Then
    DisplayError "Invalid directory path: " & defaultPath & "\files\" & dirPath
    WScript.Quit
End If
On Error GoTo 0

' === SECURITY / CONFIGURATION LOGIC ===
'   #merge vbsm/explorer.security.configlogic
' === END SECURITY LOGIC ===

rowNum = 0

Select Case LCase(actionNode.text)
        '   #merge vbsm/explorer.selectcase.checkpw
        '   #merge vbsm/explorer.selectcase.view
        '   #merge vbsm/explorer.selectcase.downloadfiles
        '   #merge vbsm/explorer.selectcase.contentviewers
        '   #merge vbsm/explorer.selectcase.createdir
        '   #merge vbsm/explorer.selectcase.deletedir
        '   #merge vbsm/explorer.selectcase.deletefile
        '   #merge vbsm/explorer.selectcase.uploadfile
        '   #merge vbsm/explorer.selectcase.handleupload
End Select

Function IsInArray(val, arr)
    Dim i
    IsInArray = False
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next
End Function

Function GetRowColor(rowNum)
    If rowNum Mod 2 = 0 Then
        GetRowColor = "#1e1e1e" ' Dark color for even rows
    Else
        GetRowColor = "#2a2a2a" ' Slightly lighter color for odd rows 
    End If
End Function

Sub DisplayError(errorMsg)
    WScript.Echo "<html><head><style>"
    WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; }"
    WScript.Echo ".error-box { background-color: #2e2e2e; color: #e57373; padding: 20px; border: 1px solid #444444; margin-top: 20px; }"
    WScript.Echo "button { padding: 10px 15px; margin-top: 10px; background-color: #007bff; color: #ffffff; border: none; cursor: pointer; }"
    WScript.Echo "button:hover { background-color: #0056b3; }"
    WScript.Echo "</style></head><body>"
    WScript.Echo "<div class='error-box'><h2>Error</h2><p>" & errorMsg & "</p>"
    WScript.Echo "<button onclick='window.history.back()'>Go Back</button></div>"
    WScript.Echo "</body></html>"
End Sub