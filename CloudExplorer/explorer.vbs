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

validActions = Array("view", "viewimage", "viewvideo", "handleupload", "createdir", "deletefile", "deletedir", "uploadfile", "checkpw")
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
' === END SECURITY LOGIC ===

rowNum = 0

Select Case LCase(actionNode.text)
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

    Case "view"
        ' --- Sorting Logic ---
        Dim sortBy, sortOrder, sortNode, orderNode
        sortBy = "name" ' default sort by name
        sortOrder = "asc" ' default ascending
        Set sortNode = xmlDoc.selectSingleNode("//sort")
        If Not sortNode Is Nothing Then sortBy = LCase(sortNode.text)
        Set orderNode = xmlDoc.selectSingleNode("//order")
        If Not orderNode Is Nothing Then sortOrder = LCase(orderNode.text)

        ' --- Build file array for sorting ---
        Dim fileArray(), fileCount, i, j, temp, comp
        fileCount = 0
        For Each file In folder.Files
            ReDim Preserve fileArray(fileCount)
            Set fileArray(fileCount) = file
            fileCount = fileCount + 1
        Next

        ' Bubble sort for files based on criteria
        For i = 0 To fileCount - 2
            For j = i + 1 To fileCount - 1
                comp = 0
                If sortBy = "size" Then
                    If fileArray(i).Size > fileArray(j).Size Then comp = 1
                    If fileArray(i).Size < fileArray(j).Size Then comp = - 1
                Else
                    comp = StrComp(fileArray(i).Name, fileArray(j).Name, vbTextCompare)
                End If
                If sortOrder = "desc" Then comp = - comp
                If comp > 0 Then
                    Set temp = fileArray(i)
                    Set fileArray(i) = fileArray(j)
                    Set fileArray(j) = temp
                End If
            Next
        Next

        ' Output HTML with a table
        WScript.Echo "<html><head><style>"
        WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; margin: 0; padding: 0; }"
        WScript.Echo "table { width: 100%; border-collapse: collapse; }"
        WScript.Echo "thead { top: 0; background-color: #1e1e1e; }"
        WScript.Echo "th, td { padding: 10px; border-bottom: 1px solid #333; text-align: left; }"
        WScript.Echo "tr:nth-child(even) td { background-color: #1e1e1e; }"
        WScript.Echo "tr:nth-child(odd) td { background-color: #2a2a2a; }"
        WScript.Echo "a { color: #80aaff; text-decoration: none; }"
        WScript.Echo ".menu-bar { background-color: #1e1e1e; padding: 10px; }"
        WScript.Echo ".menu-bar button { padding: 5px 10px; margin-right: 10px; background-color: #333333; color: #e0e0e0; border: none; }"
        WScript.Echo ".menu-bar button:hover { background-color: #444444; }"
        WScript.Echo ".up-button { float: right; }"
        WScript.Echo "th.actions-column, td.actions-column { width: 20px; text-align: right; }"
        WScript.Echo "</style></head><body>"

        ' Menu bar with additional controls
        WScript.Echo "<div class='menu-bar'>"
        WScript.Echo "<button onclick='createNewDirectory()'>New Directory</button>"
        WScript.Echo "<a href=""explorer.vbs?action=uploadfile&dir=" & dirPath & """><button>Upload file to this directory</button></a>"
        If dirPath <> "" Then
            WScript.Echo "<button class='up-button' onclick='goUp()'>&#x1F53C; Go Up</button>"
        End If
        WScript.Echo "</div>"

        ' JavaScript functions for directory creation, going up a directory, and deleting files/folders
        WScript.Echo "<script>"
        WScript.Echo "function createNewDirectory() {"
        WScript.Echo "    var name = prompt('Enter directory name:');"
        WScript.Echo "    var dir = '" & dirPath & "';"
        WScript.Echo "    if (name) {"
        WScript.Echo "        window.location.href = 'explorer.vbs?action=createdir&name=' + encodeURIComponent(name) + '&dir=' + encodeURIComponent(dir);"
        WScript.Echo "    }"
        WScript.Echo "}"
        WScript.Echo "function goUp() {"
        WScript.Echo "    var dir = '" & dirPath & "';"
        WScript.Echo "    var parentDir = dir.substring(0, dir.lastIndexOf('/'));"
        WScript.Echo "    window.location.href = 'explorer.vbs?action=view&dir=' + encodeURIComponent(parentDir);"
        WScript.Echo "}"
        WScript.Echo "function deleteItem(url) {"
        WScript.Echo "    var overlay = document.createElement('div');"
        WScript.Echo "    overlay.style.position = 'fixed';"
        WScript.Echo "    overlay.style.top = '0';"
        WScript.Echo "    overlay.style.left = '0';"
        WScript.Echo "    overlay.style.width = '100%';"
        WScript.Echo "    overlay.style.height = '100%';"
        WScript.Echo "    overlay.style.backgroundColor = 'rgba(0, 0, 0, 0.8)';"
        WScript.Echo "    overlay.style.display = 'flex';"
        WScript.Echo "    overlay.style.alignItems = 'center';"
        WScript.Echo "    overlay.style.justifyContent = 'center';"
        WScript.Echo "    overlay.style.zIndex = '1000';"
        WScript.Echo "    var popup = document.createElement('div');"
        WScript.Echo "    popup.style.backgroundColor = '#333';"
        WScript.Echo "    popup.style.color = 'white';"
        WScript.Echo "    popup.style.padding = '20px';"
        WScript.Echo "    popup.style.borderRadius = '5px';"
        WScript.Echo "    popup.innerHTML = 'Please wait';"
        WScript.Echo "    overlay.appendChild(popup);"
        WScript.Echo "    document.body.appendChild(overlay);"
        WScript.Echo "    fetch(url).then(response => {"
        WScript.Echo "        if (response.ok) {"
        WScript.Echo "            location.reload();"
        WScript.Echo "        } else {"
        WScript.Echo "            alert('Error deleting item.');"
        WScript.Echo "        }"
        WScript.Echo "    }).catch(error => {"
        WScript.Echo "        alert('Error deleting item: ' + error.message);"
        WScript.Echo "    }).finally(() => {"
        WScript.Echo "        "
        WScript.Echo "    });"
        WScript.Echo "}"
        WScript.Echo "</script>"

        ' Begin table output
        WScript.Echo "<table>"
        WScript.Echo "<thead>"
        ' First row: integrated filter (sort) options spanning all columns
        WScript.Echo "<tr>"
        WScript.Echo "<th colspan='4' style='text-align:left;'>"
        WScript.Echo "Sort Files: "
        WScript.Echo "<a href='explorer.vbs?action=view&dir=" & dirPath & "&sort=name&order=asc'>Name &uarr;</a> "
        WScript.Echo "<a href='explorer.vbs?action=view&dir=" & dirPath & "&sort=name&order=desc'>Name &darr;</a> "
        WScript.Echo "<a href='explorer.vbs?action=view&dir=" & dirPath & "&sort=size&order=asc'>Size &uarr;</a> "
        WScript.Echo "<a href='explorer.vbs?action=view&dir=" & dirPath & "&sort=size&order=desc'>Size &darr;</a>"
        WScript.Echo "</th>"
        WScript.Echo "</tr>"
        ' Second row: column headers
        WScript.Echo "<tr>"
        WScript.Echo "<th>Type</th>"
        WScript.Echo "<th>Name</th>"
        WScript.Echo "<th>Size</th>"
        WScript.Echo "<th>Actions</th>"
        WScript.Echo "</tr>"
        WScript.Echo "</thead>"
        WScript.Echo "<tbody>"

        rowNum = 0 ' Initialize rowNum

        ' List subfolders first (unsorted)
        For Each subfolder In folder.SubFolders
            rowColor = GetRowColor(rowNum)

            ' Calculate folder size and item count recursively
            folderSize = 0
            itemCount = 0
            For Each subfolderIn In subfolder.SubFolders
                itemCount = itemCount + 1
                For Each fileIn In subfolderIn.Files
                    folderSize = folderSize + fileIn.Size
                    itemCount = itemCount + 1
                Next
            Next
            For Each file In subfolder.Files
                folderSize = folderSize + file.Size
                itemCount = itemCount + 1
            Next
            
            ' Convert size to appropriate unit
            If folderSize < 1048576 Then ' Less than 1 MB
                displaySize = Round(folderSize / 1024, 2) & " KB"
            ElseIf folderSize < 1073741824 Then ' Less than 1 GB
                displaySize = Round(folderSize / 1048576, 2) & " MB"
            Else
                displaySize = Round(folderSize / 1073741824, 2) & " GB"
            End If
            
            WScript.Echo "<tr style='background-color:" & rowColor & ";'>"
            WScript.Echo "<td style='width:20px'>&#x1F4C1;</td>" ' Folder icon
            WScript.Echo "<td><a href=""explorer.vbs?action=view&dir=" & dirPath & "/" & subfolder.Name & """ style='color:#80aaff;'>" & subfolder.Name & "</a></td>"
            WScript.Echo "<td>" & displaySize & " (" & itemCount & " items)</td>" ' Display size with item count
            WScript.Echo "<td class='actions-column'><a href='#' onclick=""deleteItem('explorer.vbs?action=deletedir&name=" & subfolder.Name & "&dir=" & dirPath & "'); return false;"" style='color:#ff6666;'>&#x1F5D1;</a></td>"
            WScript.Echo "</tr>"
            rowNum = rowNum + 1
        Next
        
        ' List files with file size (sorted)
        For i = 0 To fileCount - 1
            rowColor = GetRowColor(rowNum)
            Dim fSize, displaySize
            fSize = fileArray(i).Size

            ' Convert size to appropriate unit
            If fSize < 1048576 Then ' Less than 1 MB
                displaySize = Round(fSize / 1024, 2) & " KB"
            ElseIf fSize < 1073741824 Then ' Less than 1 GB
                displaySize = Round(fSize / 1048576, 2) & " MB"
            Else
                displaySize = Round(fSize / 1073741824, 2) & " GB"
            End If
            
            WScript.Echo "<tr style='background-color:" & rowColor & ";'>"
            WScript.Echo "<td style='width:20px'>&#x1F4C4;</td>" ' File icon
            Dim fileLink
            If InStr(LCase(fileArray(i).Name), ".jpg") > 0 Or InStr(LCase(fileArray(i).Name), ".jpeg") > 0 Or InStr(LCase(fileArray(i).Name), ".png") > 0 Then
                fileLink = "explorer.vbs?action=viewimage&name=" & fileArray(i).Name & "&dir=" & dirPath
            ElseIf InStr(LCase(fileArray(i).Name), ".mp4") > 0 Or InStr(LCase(fileArray(i).Name), ".avi") > 0 Or InStr(LCase(fileArray(i).Name), ".mov") > 0 Then
                fileLink = "explorer.vbs?action=viewvideo&name=" & fileArray(i).Name & "&dir=" & dirPath
            Else
                fileLink = "Files/" & dirPath & "/" & fileArray(i).Name
            End If
            WScript.Echo "<td><a href=""" & fileLink & """ style='color:#80aaff;'>" & fileArray(i).Name & "</a></td>"
            WScript.Echo "<td>" & displaySize & "</td>" ' Display file size
            WScript.Echo "<td class='actions-column'>"
            WScript.Echo "  <a href='#' onclick=""deleteItem('explorer.vbs?action=deletefile&name=" & fileArray(i).Name & "&dir=" & dirPath & "'); return false;"" style='color:#ff6666;'>&#x1F5D1;</a> "
            WScript.Echo "  <a href='#' onclick=""navigator.clipboard.writeText(window.location.hostname+'/CloudExplorer/Files/" & dirPath & "/" & fileArray(i).Name & "'); return false;"" title='Share' style='color:#80aaff;'>&#x1F517;</a>"
            WScript.Echo "</td>"
            WScript.Echo "</tr>"
            rowNum = rowNum + 1
        Next

        WScript.Echo "</tbody></table>"
        WScript.Echo "</body></html>"
        
    Case "viewimage"
        Dim fileName, filePath
        Set nameNode = xmlDoc.selectSingleNode("//name")
        If nameNode Is Nothing Then
            WScript.Echo "No <name> node found in the XML."
            WScript.Quit
        End If
        fileName = nameNode.text
        filePath = defaultPath & "\files\" & dirPath & "\" & fileName

        ' Output HTML to display the image
        WScript.Echo "<html><head><style>"
        WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; text-align: center; padding-top: 20px; }"
        WScript.Echo ".img-container { margin: 20px auto; max-width: 90%; }"
        WScript.Echo ".download-button { display: inline-block; padding: 10px 20px; background-color: #33FF33; color: #121212; border: none; cursor: pointer; margin-top: 20px; margin-right: 10px; }"
        WScript.Echo ".download-button:hover { background-color: #28cc28; }"
        WScript.Echo ".back-button { display: inline-block; padding: 10px 20px; background-color: #FF3333; color: #121212; border: none; cursor: pointer; margin-top: 20px; }"
        WScript.Echo ".back-button:hover { background-color: #28cccc; }"
        WScript.Echo "</style></head><body>"

        ' Image display
        WScript.Echo "<div class='img-container'>"
        WScript.Echo "<img style=""width:100%""src='files/" & dirPath & "/" & fileName & "' alt='" & fileName & "' style='max-width: 100%; height: auto;' />"
        WScript.Echo "</div>"

        ' Download button
        WScript.Echo "<a href='files/" & dirPath & "/" & fileName & "' download><button class='download-button'>Download</button></a>"

        ' Go Back button
        WScript.Echo "<button class='back-button' onclick='window.history.back()'>Go Back</button>"

        WScript.Echo "</body></html>"

    Case "viewvideo"
        Dim videoName, videoPath
        Set nameNode = xmlDoc.selectSingleNode("//name")
        If nameNode Is Nothing Then
            WScript.Echo "No <name> node found in the XML."
            WScript.Quit
        End If
        videoName = nameNode.text
        videoPath = defaultPath & "\files\" & dirPath & "\" & videoName

        ' Output HTML to display the video
        WScript.Echo "<html><head><style>"
        WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; text-align: center; padding-top: 20px; }"
        WScript.Echo ".video-container { margin: 20px auto; max-width: 90%; }"
        WScript.Echo ".download-button { display: inline-block; padding: 10px 20px; background-color: #33FF33; color: #121212; border: none; cursor: pointer; margin-top: 20px; margin-right: 10px; }"
        WScript.Echo ".download-button:hover { background-color: #28cc28; }"
        WScript.Echo ".back-button { display: inline-block; padding: 10px 20px; background-color: #FF3333; color: #121212; border: none; cursor: pointer; margin-top: 20px; }"
        WScript.Echo ".back-button:hover { background-color: #cc2828; }"
        WScript.Echo "</style></head><body>"

        ' Video display
        WScript.Echo "<div class='video-container'>"
        WScript.Echo "<video controls style='width: 100%; height: auto;'>"
        WScript.Echo "<source src='files/" & dirPath & "/" & videoName & "' type='video/mp4'>"
        WScript.Echo "Your browser does not support the video tag."
        WScript.Echo "</video>"
        WScript.Echo "</div>"

        ' Download button
        WScript.Echo "<a href='files/" & dirPath & "/" & videoName & "' download><button class='download-button'>Download</button></a>"

        ' Go Back button
        WScript.Echo "<button class='back-button' onclick='window.history.back()'>Go Back</button>"

        WScript.Echo "</body></html>"

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

    Case "uploadfile"
        ' Output upload form with dark mode styling, multi-file support, live upload feed, and auto-scroll
        WScript.Echo "<html><head><style>"
        WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; }"
        WScript.Echo ".menu-bar { background-color: #1e1e1e; color: #e0e0e0; padding: 10px; text-align: center; }"
        WScript.Echo ".menu-bar button { padding: 10px 20px; margin-right: 10px; background-color: #333333; color: #e0e0e0; border: none; cursor: pointer; border-radius: 5px; }"
        WScript.Echo ".menu-bar button:hover { background-color: #444444; }"
        WScript.Echo "#uploadFeed { margin-top: 20px; text-align: left; max-height: 150px; overflow-y: auto; font-size: 14px; border: 1px solid #444; padding: 5px; }"
        WScript.Echo "</style></head><body>"
        WScript.Echo "<div class='menu-bar'><button onclick='window.location.href=""explorer.vbs?action=view&dir=" & dirPath & """'>Back to Directory</button></div>"
        WScript.Echo "<div style='margin:20px auto; padding:20px; background-color: #1e1e1e; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.5); width:500px; text-align: center;'>"
        WScript.Echo "<h2>Upload Files</h2>"
        WScript.Echo "<form id='uploadForm' method='post' enctype='multipart/form-data' style='margin:20px auto; padding:20px; background-color: #1e1e1e; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.5); width:300px;'>"
        WScript.Echo "<input type='file' id='fileInput' name='file' multiple style='margin:10px 0; display:block; width:100%;'><br>" ' Removed required
        WScript.Echo "<input type='file' id='folderInput' name='folder' webkitdirectory directory style='margin:10px 0; display:block; width:100%;'><br>"
        WScript.Echo "<div style='margin-top:10px;'>"
        WScript.Echo "<button type='submit' style='display:block; width:100px; padding:10px 20px; background-color: #4caf50; color:#ffffff; border:none; cursor:pointer; border-radius:5px; margin:10px auto;'>Upload</button>"
        WScript.Echo "<button type='button' style='display:block; width:100px; padding:10px 20px; background-color: #f44336; color:#ffffff; border:none; cursor:pointer; border-radius:5px; margin:10px auto;' onclick='window.history.back()'>Go Back</button>"
        WScript.Echo "</div>"
        WScript.Echo "</form>"
        ' Live upload feed container
        WScript.Echo "<div id='uploadFeed'></div>"
        WScript.Echo "</div>"
        WScript.Echo "<script>"
        WScript.Echo "document.getElementById('uploadForm').addEventListener('submit', function(event) {"
        WScript.Echo "    event.preventDefault();"
        WScript.Echo "    var fileInput = document.getElementById('fileInput');"
        WScript.Echo "    var folderInput = document.getElementById('folderInput');"
        WScript.Echo "    var files = fileInput.files;"
        WScript.Echo "    var folders = folderInput.files;"
        WScript.Echo "    var dir = '" & dirPath & "';"
        WScript.Echo "    var feed = document.getElementById('uploadFeed');"
        WScript.Echo "    feed.innerHTML = '';"
        WScript.Echo "    function uploadNext(index, items) {"
        WScript.Echo "        if (index >= items.length) {"
        WScript.Echo "            window.location.href = 'explorer.vbs?action=view&dir=' + encodeURIComponent(dir);"
        WScript.Echo "            return;"
        WScript.Echo "        }"
        WScript.Echo "        var item = items[index];"
        WScript.Echo "        var entry = document.createElement('div');"
        WScript.Echo "        entry.id = 'entry_' + index;"
        WScript.Echo "        entry.innerHTML = 'Uploading ' + item.name + '...';"
        WScript.Echo "        feed.appendChild(entry);"
        WScript.Echo "        feed.scrollTop = feed.scrollHeight;" ' auto-scroll to bottom
        WScript.Echo "        var formData = new FormData();"
        WScript.Echo "        formData.append('file', item);"
        WScript.Echo "        var xhr = new XMLHttpRequest();"
        WScript.Echo "        xhr.open('POST', '/upload?filename=' + encodeURIComponent(item.name) + '&redirectto=cloudexplorer/explorer.vbs%3Faction%3Dhandleupload%26dir%3D' + encodeURIComponent(dir));"
        WScript.Echo "        xhr.onreadystatechange = function() {"
        WScript.Echo "            if (xhr.readyState == 4) {"
        WScript.Echo "                if(xhr.status == 200) {"
        WScript.Echo "                    entry.innerHTML = 'Uploaded ' + item.name;"
        WScript.Echo "                } else {"
        WScript.Echo "                    entry.innerHTML = 'Error uploading ' + item.name;"
        WScript.Echo "                }"
        WScript.Echo "                feed.scrollTop = feed.scrollHeight;"
        WScript.Echo "                uploadNext(index + 1, items);"
        WScript.Echo "            }"
        WScript.Echo "        };"
        WScript.Echo "        xhr.send(formData);"
        WScript.Echo "    }"
        WScript.Echo "    function checkConflicts(items) {"
        WScript.Echo "        var conflict = false;"
        WScript.Echo "        var uploadedFiles = [];"
        WScript.Echo "        for (var i = 0; i < items.length; i++) {"
        WScript.Echo "            if (uploadedFiles.includes(items[i].name)) {"
        WScript.Echo "                alert('Naming conflict detected: ' + items[i].name);"
        WScript.Echo "                conflict = true;"
        WScript.Echo "                break;"
        WScript.Echo "            }"
        WScript.Echo "            uploadedFiles.push(items[i].name);"
        WScript.Echo "        }"
        WScript.Echo "        return conflict;"
        WScript.Echo "    }"
        WScript.Echo "    var allItems = []"
        WScript.Echo "    if (files.length > 0) {"
        WScript.Echo "        allItems = [...files];"
        WScript.Echo "    }"
        WScript.Echo "    if (folders.length > 0) {"
        WScript.Echo "        allItems = [...allItems, ...folders];"
        WScript.Echo "    }"
        WScript.Echo "    allItems = allItems.filter(function(item) { return item.webkitRelativePath.split('/').length <= 2; });" ' Filter out subfolder items
        WScript.Echo "    if (checkConflicts(allItems)) {"
        WScript.Echo "        return;"
        WScript.Echo "    }"
        WScript.Echo "    uploadNext(0, allItems);"
        WScript.Echo "});"
        WScript.Echo "</script>"
        WScript.Echo "</body></html>"

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