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
Dim vSortBy, vSortOrder, vSortNode, vOrderNode
vSortBy = "name" ' default sort by name
vSortOrder = "asc" ' default ascending
Set vSortNode = xmlDoc.selectSingleNode("//sort")
If Not vSortNode Is Nothing Then vSortBy = LCase(vSortNode.text)
Set vOrderNode = xmlDoc.selectSingleNode("//order")
If Not vOrderNode Is Nothing Then vSortOrder = LCase(vOrderNode.text)

' --- Build file array for sorting ---
Dim vFileArray(), vFileCount, vI, vJ, vTemp, vComp
vFileCount = 0
For Each file In folder.Files
    ReDim Preserve vFileArray(vFileCount)
    Set vFileArray(vFileCount) = file
    vFileCount = vFileCount + 1
Next

' Bubble sort for files based on criteria
For vI = 0 To vFileCount - 2
    For vJ = vI + 1 To vFileCount - 1
        vComp = 0
        If vSortBy = "size" Then
            If vFileArray(vI).Size > vFileArray(vJ).Size Then vComp = 1
            If vFileArray(vI).Size < vFileArray(vJ).Size Then vComp = - 1
        Else
            vComp = StrComp(vFileArray(vI).Name, vFileArray(vJ).Name, vbTextCompare)
        End If
        If vSortOrder = "desc" Then vComp = - vComp
        If vComp > 0 Then
            Set vTemp = vFileArray(vI)
            Set vFileArray(vI) = vFileArray(vJ)
            Set vFileArray(vJ) = vTemp
        End If
    Next
Next

' Output HTML with a table and include the file checkbox in the actions column.
WScript.Echo "<html><head><meta charset='UTF-8'><title>Directory View</title><style>"
WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; margin: 0; padding: 0; }"
WScript.Echo "table { width: 100%; border-collapse: collapse; }"
WScript.Echo "thead { position: sticky; top: 0; background-color: #1e1e1e; }"
WScript.Echo "th, td { padding: 10px; border-bottom: 1px solid #333; text-align: left; }"
WScript.Echo "tr:nth-child(even) td { background-color: #1e1e1e; }"
WScript.Echo "tr:nth-child(odd) td { background-color: #2a2a2a; }"
WScript.Echo "a { color: #80aaff; text-decoration: none; }"
WScript.Echo ".menu-bar { background-color: #1e1e1e; padding: 10px; }"
WScript.Echo ".menu-bar button { padding: 5px 10px; margin-right: 10px; background-color: #333333; color: #e0e0e0; border: none; cursor: pointer; }"
WScript.Echo ".menu-bar button:hover { background-color: #444444; }"
WScript.Echo ".up-button { float: right; }"
WScript.Echo "th.actions-column, td.actions-column { width: 90px; text-align: right; }"
WScript.Echo "</style></head><body>"

' Menu bar with additional controls and hidden "Download Selected" and "Delete Selected" buttons.
WScript.Echo "<div class='menu-bar'>"
WScript.Echo "<button onclick='createNewDirectory()'>New Directory</button>"
WScript.Echo "<a href=""explorer.vbs?action=uploadfile&dir=" & dirPath & """ onclick=""showPleaseWait()""><button>Upload file to this directory</button></a>"
If dirPath <> "" Then
    WScript.Echo "<button class='up-button' onclick='goUp()'>&#x1F53C; Go Up</button>"
End If
WScript.Echo "<button id='downloadButton' style='display:none;' onclick='showPleaseWait();downloadSelected()'>Download Selected</button>"
WScript.Echo "<button id='deleteButton' style='display:none;' onclick='showPleaseWait();deleteSelected()'>Delete Selected</button>"
WScript.Echo "</div>"

' JavaScript functions.
WScript.Echo "<script>"
' Create New Directory
WScript.Echo "function createNewDirectory() {"
WScript.Echo "    var name = prompt('Enter directory name:');"
WScript.Echo "    var dir = '" & dirPath & "';"
WScript.Echo "    if (name) {"
WScript.Echo "        showPleaseWait();"
WScript.Echo "        window.location.href = 'explorer.vbs?action=createdir&name=' + encodeURIComponent(name) + '&dir=' + encodeURIComponent(dir);"
WScript.Echo "    }"
WScript.Echo "}"
' Go Up Directory
WScript.Echo "function goUp() {"
WScript.Echo "    var dir = '" & dirPath & "';"
WScript.Echo "    var parentDir = dir.substring(0, dir.lastIndexOf('/'));"
WScript.Echo "    showPleaseWait();"
WScript.Echo "    window.location.href = 'explorer.vbs?action=view&dir=' + encodeURIComponent(parentDir);"
WScript.Echo "}"
' Standard delete with Please Wait overlay.
WScript.Echo "function showPleaseWait() {"
WScript.Echo "  var overlay = document.createElement('div');"
WScript.Echo "  overlay.id = 'pleaseWaitOverlay';"
WScript.Echo "  overlay.style.position = 'fixed';"
WScript.Echo "  overlay.style.top = '0';"
WScript.Echo "  overlay.style.left = '0';"
WScript.Echo "  overlay.style.width = '100%';"
WScript.Echo "  overlay.style.height = '100%';"
WScript.Echo "  overlay.style.backgroundColor = 'rgba(0,0,0,0.8)';"
WScript.Echo "  overlay.style.display = 'flex';"
WScript.Echo "  overlay.style.alignItems = 'center';"
WScript.Echo "  overlay.style.justifyContent = 'center';"
WScript.Echo "  overlay.style.zIndex = '1000';"
WScript.Echo "  var popup = document.createElement('div');"
WScript.Echo "  popup.style.backgroundColor = '#333';"
WScript.Echo "  popup.style.color = 'white';"
WScript.Echo "  popup.style.padding = '20px';"
WScript.Echo "  popup.style.borderRadius = '5px';"
WScript.Echo "  popup.innerHTML = 'Please wait...';"
WScript.Echo "  overlay.appendChild(popup);"
WScript.Echo "  document.body.appendChild(overlay);"
WScript.Echo "}"

' Add an event listener for detecting history navigation
WScript.Echo "window.addEventListener('pageshow', function(event) {"
WScript.Echo "  if (event.persisted) {"
WScript.Echo "    var overlay = document.getElementById('pleaseWaitOverlay');"
WScript.Echo "    if (overlay) {"
WScript.Echo "      overlay.parentNode.removeChild(overlay);"
WScript.Echo "    }"
WScript.Echo "  }"
WScript.Echo "});"

WScript.Echo "function hidePleaseWait() {"
WScript.Echo "  var overlay = document.getElementById('pleaseWaitOverlay');"
WScript.Echo "  if(overlay) { overlay.parentNode.removeChild(overlay); }"
WScript.Echo "}"
' Existing deleteItem function remains, but used here only for single deletions.
WScript.Echo "function deleteItem(url) {"
WScript.Echo "    showPleaseWait();"
WScript.Echo "    fetch(url).then(response => {"
WScript.Echo "        if (response.ok) {"
WScript.Echo "            location.reload();"
WScript.Echo "        } else {"
WScript.Echo "            alert('Error deleting item.');"
WScript.Echo "            hidePleaseWait();"
WScript.Echo "        }"
WScript.Echo "    }).catch(error => {"
WScript.Echo "        alert('Error deleting item: ' + error.message);"
WScript.Echo "        hidePleaseWait();"
WScript.Echo "    });"
WScript.Echo "}"
' Function to compute ranges.
WScript.Echo "function computeRanges(arr) {"
WScript.Echo "  arr.sort(function(a, b){return a-b});"
WScript.Echo "  var ranges = [];"
WScript.Echo "  var start = arr[0], end = arr[0];"
WScript.Echo "  for(var i=1;i<arr.length;i++){"
WScript.Echo "    if(arr[i] == end+1){"
WScript.Echo "      end = arr[i];"
WScript.Echo "    } else {"
WScript.Echo "      ranges.push(start == end ? start.toString() : start + '-' + end);"
WScript.Echo "      start = arr[i]; end = arr[i];"
WScript.Echo "    }"
WScript.Echo "  }"
WScript.Echo "  ranges.push(start == end ? start.toString() : start + '-' + end);"
WScript.Echo "  return ranges.join(',');"
WScript.Echo "}"
' Update both action buttons visibility.
WScript.Echo "function updateActionButtons() {"
WScript.Echo "  var checkboxes = document.querySelectorAll('.file-checkbox');"
WScript.Echo "  var anyChecked = false;"
WScript.Echo "  for (var i=0; i<checkboxes.length; i++) {"
WScript.Echo "      if(checkboxes[i].checked) { anyChecked = true; break; }"
WScript.Echo "  }"
WScript.Echo "  document.getElementById('downloadButton').style.display = anyChecked ? 'inline-block' : 'none';"
WScript.Echo "  document.getElementById('deleteButton').style.display = anyChecked ? 'inline-block' : 'none';"
WScript.Echo "}"
' Function for the header checkbox.
WScript.Echo "function toggleSelectAll(source) {"
WScript.Echo "  var checkboxes = document.querySelectorAll('.file-checkbox');"
WScript.Echo "  for(var i=0;i<checkboxes.length;i++){"
WScript.Echo "      checkboxes[i].checked = source.checked;"
WScript.Echo "  }"
WScript.Echo "  updateActionButtons();"
WScript.Echo "}"
' Download Selected: gathers indices and redirects to downloadfiles action.
WScript.Echo "function downloadSelected() {"
WScript.Echo "  var checkboxes = document.querySelectorAll('.file-checkbox');"
WScript.Echo "  var selected = [];"
WScript.Echo "  for(var i=0;i<checkboxes.length;i++){"
WScript.Echo "      if(checkboxes[i].checked) { selected.push(parseInt(checkboxes[i].getAttribute('data-index'))); }"
WScript.Echo "  }"
WScript.Echo "  if(selected.length === 0){ return; }"
WScript.Echo "  var ranges = computeRanges(selected);"
WScript.Echo "  var url = 'explorer.vbs?action=downloadfiles&dir=' + encodeURIComponent('" & dirPath & "') + '&files=' + encodeURIComponent(ranges);"
WScript.Echo "  window.location.href = url;"
WScript.Echo "}"
' Delete Selected: gathers indices and sends delete GET requests.
WScript.Echo "function deleteSelected() {"
WScript.Echo "  var checkboxes = document.querySelectorAll('.file-checkbox');"
WScript.Echo "  var selected = [];"
WScript.Echo "  for(var i=0;i<checkboxes.length;i++){"
WScript.Echo "      if(checkboxes[i].checked) { selected.push(parseInt(checkboxes[i].getAttribute('data-index'))); }"
WScript.Echo "  }"
WScript.Echo "  if(selected.length === 0){ return; }"
WScript.Echo "  showPleaseWait();"
WScript.Echo "  var deletePromises = [];"
WScript.Echo "  for(var i=0;i<selected.length;i++){"
WScript.Echo "    var idx = selected[i];"
WScript.Echo "    var url = 'explorer.vbs?action=deletefile&dir=' + encodeURIComponent('" & dirPath & "') + '&name=' + encodeURIComponent(fileNames[idx]);"
WScript.Echo "    deletePromises.push(fetch(url));"
WScript.Echo "  }"
WScript.Echo "  Promise.all(deletePromises).then(function(results){"
WScript.Echo "      location.reload();"
WScript.Echo "  }).catch(function(err){"
WScript.Echo "      alert('Error deleting files: ' + err.message);"
WScript.Echo "      hidePleaseWait();"
WScript.Echo "  });"
WScript.Echo "}"
' Attach updateActionButtons() to each file checkbox change event.
WScript.Echo "document.addEventListener('DOMContentLoaded', function(){"
WScript.Echo "  var checkboxes = document.querySelectorAll('.file-checkbox');"
WScript.Echo "  for(var i=0;i<checkboxes.length;i++){"
WScript.Echo "      checkboxes[i].addEventListener('change', updateActionButtons);"
WScript.Echo "  }"
WScript.Echo "});"
' Provide an array of file names from the sorted file list.
WScript.Echo "var fileNames = ["
For vI = 0 To vFileCount - 1
    WScript.Echo "  '" & Replace(vFileArray(vI).Name, "'", "\'") & "',"
Next
WScript.Echo "];"
WScript.Echo "var dirPath = '" & dirPath & "';"
WScript.Echo "</script>"

' Begin table output.
WScript.Echo "<table>"
WScript.Echo "<thead>"
' First row: sort options.
WScript.Echo "<tr>"
WScript.Echo "<th colspan='4' style='text-align:left;'>"
WScript.Echo "Sort Files: "
WScript.Echo "<a onclick=""showPleaseWait()"" href='explorer.vbs?action=view&dir=" & dirPath & "&sort=name&order=asc'>Name &uarr;</a> "
WScript.Echo "<a onclick=""showPleaseWait()"" href='explorer.vbs?action=view&dir=" & dirPath & "&sort=name&order=desc'>Name &darr;</a> "
WScript.Echo "<a onclick=""showPleaseWait()"" href='explorer.vbs?action=view&dir=" & dirPath & "&sort=size&order=asc'>Size &uarr;</a> "
WScript.Echo "<a onclick=""showPleaseWait()"" href='explorer.vbs?action=view&dir=" & dirPath & "&sort=size&order=desc'>Size &darr;</a>"
WScript.Echo "</th>"
WScript.Echo "</tr>"
' Second row: headers (with header checkbox in the Actions column).
WScript.Echo "<tr>"
WScript.Echo "<th>Type</th>"
WScript.Echo "<th>Name</th>"
WScript.Echo "<th>Size</th>"
WScript.Echo "<th>Actions <input type='checkbox' onclick='toggleSelectAll(this)' title='Select/Deselect all files'></th>"
WScript.Echo "</tr>"
WScript.Echo "</thead>"
WScript.Echo "<tbody>"

Dim vRowNum, vRowColor, vFolderSize, vItemCount, vDisplaySize
vRowNum = 0 ' Initialize rowNum

' List subfolders (unsorted).
For Each subfolder In folder.SubFolders
    vRowColor = GetRowColor(vRowNum)
    vFolderSize = 0
    vItemCount = 0
    For Each subfolderIn In subfolder.SubFolders
        vItemCount = vItemCount + 1
        For Each fileIn In subfolderIn.Files
            vFolderSize = vFolderSize + fileIn.Size
            vItemCount = vItemCount + 1
        Next
    Next
    For Each file In subfolder.Files
        vFolderSize = vFolderSize + file.Size
        vItemCount = vItemCount + 1
    Next
    If vFolderSize < 1048576 Then
        vDisplaySize = Round(vFolderSize / 1024, 2) & " KB"
    ElseIf vFolderSize < 1073741824 Then
        vDisplaySize = Round(vFolderSize / 1048576, 2) & " MB"
    Else
        vDisplaySize = Round(vFolderSize / 1073741824, 2) & " GB"
    End If
    WScript.Echo "<tr style='background-color:" & vRowColor & ";'>"
    WScript.Echo "<td style='width:20px'>&#x1F4C1;</td>"
    WScript.Echo "<td><a onclick=""showPleaseWait()"" href=""explorer.vbs?action=view&dir=" & dirPath & "/" & subfolder.Name & """ style='color:#80aaff;'>" & subfolder.Name & "</a></td>"
    WScript.Echo "<td>" & vDisplaySize & " (" & vItemCount & " items)</td>"
    WScript.Echo "<td class='actions-column'>"
    WScript.Echo "<a href='#' onclick=""showPleaseWait();deleteItem('explorer.vbs?action=deletedir&name=" & subfolder.Name & "&dir=" & dirPath & "'); return false;"" style='color:#ff6666;'>&#x1F5D1;</a>"
    WScript.Echo "</td>"
    WScript.Echo "</tr>"
    vRowNum = vRowNum + 1
Next

' List files (sorted) – with download/delete checkbox in the actions column.
For vI = 0 To vFileCount - 1
    vRowColor = GetRowColor(vRowNum)
    Dim vFSize, vFileLink
    vFSize = vFileArray(vI).Size
    If vFSize < 1048576 Then
        vDisplaySize = Round(vFSize / 1024, 2) & " KB"
    ElseIf vFSize < 1073741824 Then
        vDisplaySize = Round(vFSize / 1048576, 2) & " MB"
    Else
        vDisplaySize = Round(vFSize / 1073741824, 2) & " GB"
    End If
    WScript.Echo "<tr style='background-color:" & vRowColor & ";'>"
    WScript.Echo "<td style='width:20px'>&#x1F4C4;</td>"
    ' Determine file link based on file type.
    Dim viewerFile
    viewerFile = False
    If InStr(LCase(vFileArray(vI).Name), ".jpg") > 0 Or InStr(LCase(vFileArray(vI).Name), ".jpeg") > 0 Or InStr(LCase(vFileArray(vI).Name), ".png") > 0 Or InStr(LCase(vFileArray(vI).Name), ".gif") > 0 Or InStr(LCase(vFileArray(vI).Name), ".webp") > 0 Or InStr(LCase(vFileArray(vI).Name), ".bmp") > 0 Or InStr(LCase(vFileArray(vI).Name), ".tiff") > 0 Then
        vFileLink = "explorer.vbs?action=viewimage&name=" & vFileArray(vI).Name & "&dir=" & dirPath
        viewerFile = True
    ElseIf InStr(LCase(vFileArray(vI).Name), ".mp4") > 0 Or InStr(LCase(vFileArray(vI).Name), ".avi") > 0 Or InStr(LCase(vFileArray(vI).Name), ".mov") > 0 Or InStr(LCase(vFileArray(vI).Name), ".webm") > 0 Or InStr(LCase(vFileArray(vI).Name), ".mkv") > 0 Or InStr(LCase(vFileArray(vI).Name), ".flv") > 0 Or InStr(LCase(vFileArray(vI).Name), ".wmv") > 0 Then
        vFileLink = "explorer.vbs?action=viewvideo&name=" & vFileArray(vI).Name & "&dir=" & dirPath
        viewerFile = True
    Else
        vFileLink = "Files/" & dirPath & "/" & vFileArray(vI).Name
        viewerFile = False
    End If
    If viewerFile Then
        WScript.Echo "<td><a onclick=""showPleaseWait()"" href=""" & vFileLink & """ style='color:#80aaff;'>" & vFileArray(vI).Name & "</a></td>"
    Else
        WScript.Echo "<td><a href=""" & vFileLink & """ style='color:#80aaff;'>" & vFileArray(vI).Name & "</a></td>"
    End If
    WScript.Echo "<td>" & vDisplaySize & "</td>"
    ' In the actions column, include delete/share links AND the file checkbox.
    WScript.Echo "<td class='actions-column'>"
    WScript.Echo "<a href='#' onclick=""showPleaseWait();deleteItem('explorer.vbs?action=deletefile&name=" & vFileArray(vI).Name & "&dir=" & dirPath & "'); return false;"" style='color:#ff6666;'>&#x1F5D1;</a> "
    WScript.Echo "<a href='#' onclick=""navigator.clipboard.writeText(window.location.hostname+'/CloudExplorer/Files/" & dirPath & "/" & vFileArray(vI).Name & "'); return false;"" title='Share' style='color:#80aaff;'>&#x1F517;</a> "
    WScript.Echo "<input type='checkbox' class='file-checkbox' data-index='" & vI & "' onclick='updateActionButtons()'>"
    WScript.Echo "</td>"
    WScript.Echo "</tr>"
    vRowNum = vRowNum + 1
Next

WScript.Echo "</tbody></table>"
WScript.Echo "</body></html>"
Case "downloadfiles"
        ' Retrieve the file ranges string from the XML
        Dim dFilesNode, dFilesParam, dSelectedIndices, dIndices, dToken, dParts, dK, dSelCount
        Set dFilesNode = xmlDoc.selectSingleNode("//files")
        If dFilesNode Is Nothing Then
            DisplayError "No <files> node provided."
            WScript.Quit
        End If
        dFilesParam = dFilesNode.text ' e.g. "0-2,4,6-8"
        
        ' Build an array of individual indices from the range string
        dSelectedIndices = Array()
        dIndices = Split(dFilesParam, ",")
        dSelCount = 0
        For dK = 0 To UBound(dIndices)
            dToken = Trim(dIndices(dK))
            If InStr(dToken, "-") > 0 Then
                dParts = Split(dToken, "-")
                Dim dStartIdx, dEndIdx, dZ
                dStartIdx = CInt(dParts(0))
                dEndIdx = CInt(dParts(1))
                For dZ = dStartIdx To dEndIdx
                    ReDim Preserve dSelectedIndices(dSelCount)
                    dSelectedIndices(dSelCount) = dZ
                    dSelCount = dSelCount + 1
                Next
            Else
                ReDim Preserve dSelectedIndices(dSelCount)
                dSelectedIndices(dSelCount) = CInt(dToken)
                dSelCount = dSelCount + 1
            End If
        Next
        
        ' Rebuild the sorted file list
        Dim dAllFiles(), dFileCount2, dF, dM
        dFileCount2 = 0
        For Each dF In folder.Files
            ReDim Preserve dAllFiles(dFileCount2)
            Set dAllFiles(dFileCount2) = dF
            dFileCount2 = dFileCount2 + 1
        Next
        
        ' Sort files by name (Bubble Sort)
        For dM = 0 To dFileCount2 - 2
            For dJ = dM + 1 To dFileCount2 - 1
                If StrComp(dAllFiles(dM).Name, dAllFiles(dJ).Name, vbTextCompare) > 0 Then
                    Set dTemp = dAllFiles(dM)
                    Set dAllFiles(dM) = dAllFiles(dJ)
                    Set dAllFiles(dJ) = dTemp
                End If
            Next
        Next
        
        ' Gather the files to be downloaded
        Dim dFilesToDownload(), dCountDownload, dIdx
        dCountDownload = 0
        For dM = 0 To UBound(dSelectedIndices)
            dIdx = dSelectedIndices(dM)
            If dIdx >= 0 And dIdx < dFileCount2 Then
                ReDim Preserve dFilesToDownload(dCountDownload)
                Set dFilesToDownload(dCountDownload) = dAllFiles(dIdx)
                dCountDownload = dCountDownload + 1
            End If
        Next
        
        ' Output an HTML page with progress indication
        WScript.Echo "<html><head><meta charset='UTF-8'><title>Downloading Files</title>"
        WScript.Echo "<style>"
        WScript.Echo "body { font-family: Arial, sans-serif; background-color: #121212; color: #e0e0e0; padding: 20px; }"
        WScript.Echo ".menu-bar { background-color: #1e1e1e; color: #e0e0e0; padding: 10px; text-align: center; }"
        WScript.Echo ".menu-bar button { padding: 10px 20px; margin-right: 10px; background-color: #333333; color: #e0e0e0; border: none; cursor: pointer; border-radius: 5px; }"
        WScript.Echo ".menu-bar button:hover { background-color: #444444; }"
        WScript.Echo "#downloadFeed { max-height: 150px; overflow-y: auto; border: 1px solid #444; padding: 10px; background: #1e1e1e; }"
        WScript.Echo ".progress-bar { width: 100%; background-color: #444; border-radius: 5px; margin-top: 10px; }"
        WScript.Echo ".progress { height: 20px; background-color: #4caf50; width: 0%; border-radius: 5px; }"
        WScript.Echo ".progress-text { text-align: center; margin-top: 5px; }"
        WScript.Echo "</style></head><body>"
        
        WScript.Echo "<div class='menu-bar'>"
        WScript.Echo "<button onclick='window.location.href=""explorer.vbs?action=view&dir=" & dirPath & """'>Back to Directory</button>"
        WScript.Echo "</div>"
        WScript.Echo "<h2>Downloading Files</h2>"
        WScript.Echo "<p>Downloads will start automatically. Progress will be shown below:</p>"
        
        ' Progress bar container
        WScript.Echo "<div class='progress-bar'><div class='progress' id='progressBar'></div></div>"
        WScript.Echo "<div class='progress-text' id='progressText'>0%</div>"
        
        ' Download log box
        WScript.Echo "<div id='downloadFeed'></div>"
        
        WScript.Echo "<script>"
        WScript.Echo "let wakeLock = null;"
        WScript.Echo "async function requestWakeLock(){"
        WScript.Echo "try {"
        WScript.Echo "wakeLock = await navigator.wakeLock.request('screen');"
        WScript.Echo "console.log('Wake Lock active.');"
        WScript.Echo "wakeLock.addEventListener('release', () => { console.log('Wake Lock was released'); }"
        WScript.Echo ");"
        WScript.Echo "} catch (err) { console.error(`${err.name}, ${err.message}`);}}"
        WScript.Echo "document.addEventListener('visibilitychange', () => {    if (wakeLock !== null && document.visibilityState === 'visible') {        requestWakeLock();    }});"
        WScript.Echo "requestWakeLock();"

        ' JavaScript functions for handling progress
        WScript.Echo "var files = ["
        For dM = 0 To UBound(dFilesToDownload)
            dFileURL = "Files/" & dirPath & "/" & dFilesToDownload(dM).Name
            WScript.Echo "'" & dFileURL & "',"
        Next
        WScript.Echo "];"
        
        WScript.Echo "var totalFiles = files.length, downloadedFiles = 0;"
        WScript.Echo "function downloadNext(index) {"
        WScript.Echo "    if (index >= files.length) {"
        WScript.Echo "        document.getElementById('progressText').innerHTML = 'Download Complete';"
        WScript.Echo "        return;"
        WScript.Echo "    }"
        WScript.Echo "    var fileURL = files[index];"
        WScript.Echo "    var iframe = document.createElement('iframe');"
        WScript.Echo "    iframe.style.display = 'none';"
        WScript.Echo "    iframe.src = fileURL;"
        WScript.Echo "    document.body.appendChild(iframe);"
        
        ' Append download status to the scroll box
        WScript.Echo "    var feed = document.getElementById('downloadFeed');"
        WScript.Echo "    var entry = document.createElement('div');"
        WScript.Echo "    entry.innerHTML = 'Downloading: ' + fileURL.split('/').pop();"
        WScript.Echo "    feed.appendChild(entry);"
        WScript.Echo "    feed.scrollTop = feed.scrollHeight;"
        
        ' Update progress bar
        WScript.Echo "    downloadedFiles++;"
        WScript.Echo "    var progress = Math.round((downloadedFiles / totalFiles) * 100);"
        WScript.Echo "    document.getElementById('progressBar').style.width = progress + '%';"
        WScript.Echo "    document.getElementById('progressText').innerHTML = progress + '%';"
        
        ' Move to the next file after a short delay
        WScript.Echo "    setTimeout(function() { downloadNext(index + 1); }, 800);"
        WScript.Echo "}"
        
        WScript.Echo "downloadNext(0);" ' Start the downloads
        WScript.Echo "</script>"
        
        WScript.Echo "</body></html>"
        WScript.Quit
    
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
        WScript.Echo ".back-button:hover { background-color: #cc2828; }"
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
        WScript.Echo ".progress-bar { width: 100%; background-color: #444; border-radius: 5px; margin-top: 10px; }"
        WScript.Echo ".progress { height: 20px; background-color: #4caf50; width: 0%; border-radius: 5px; }"
        WScript.Echo ".progress-text { color: #e0e0e0; text-align: center; }"
        WScript.Echo "</style></head><body>"
        ' WScript.Echo "<div class='menu-bar'><button onclick='showPleaseWait();window.location.href=""explorer.vbs?action=view&dir=" & dirPath & """'>Back to Directory</button></div>"
        WScript.Echo "<div style='margin:20px auto; padding:20px; background-color: #1e1e1e; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.5); width:500px; text-align: center;'>"
        WScript.Echo "<h2>Upload Files</h2>"
        WScript.Echo "<form id='uploadForm' method='post' enctype='multipart/form-data' style='margin:20px auto; padding:20px; background-color: #1e1e1e; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.5); width:300px;'>"
        WScript.Echo "<input type='file' id='fileInput' name='file' multiple style='margin:10px 0; display:block; width:100%;'><br>"
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
        ' Standard delete with Please Wait overlay.
        WScript.Echo "function showPleaseWait() {"
        WScript.Echo "  var overlay = document.createElement('div');"
        WScript.Echo "  overlay.id = 'pleaseWaitOverlay';"
        WScript.Echo "  overlay.style.position = 'fixed';"
        WScript.Echo "  overlay.style.top = '0';"
        WScript.Echo "  overlay.style.left = '0';"
        WScript.Echo "  overlay.style.width = '100%';"
        WScript.Echo "  overlay.style.height = '100%';"
        WScript.Echo "  overlay.style.backgroundColor = 'rgba(0,0,0,0.8)';"
        WScript.Echo "  overlay.style.display = 'flex';"
        WScript.Echo "  overlay.style.alignItems = 'center';"
        WScript.Echo "  overlay.style.justifyContent = 'center';"
        WScript.Echo "  overlay.style.zIndex = '1000';"
        WScript.Echo "  var popup = document.createElement('div');"
        WScript.Echo "  popup.style.backgroundColor = '#333';"
        WScript.Echo "  popup.style.color = 'white';"
        WScript.Echo "  popup.style.padding = '20px';"
        WScript.Echo "  popup.style.borderRadius = '5px';"
        WScript.Echo "  popup.innerHTML = 'Please wait...';"
        WScript.Echo "  overlay.appendChild(popup);"
        WScript.Echo "  document.body.appendChild(overlay);"
        WScript.Echo "}"
        
        ' Add an event listener for detecting history navigation
        WScript.Echo "window.addEventListener('pageshow', function(event) {"
        WScript.Echo "  if (event.persisted) {"
        WScript.Echo "    var overlay = document.getElementById('pleaseWaitOverlay');"
        WScript.Echo "    if (overlay) {"
        WScript.Echo "      overlay.parentNode.removeChild(overlay);"
        WScript.Echo "    }"
        WScript.Echo "  }"
        WScript.Echo "});"
        
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
        WScript.Echo "        var progressBar = document.createElement('div');"
        WScript.Echo "        progressBar.className = 'progress-bar';"
        WScript.Echo "        var progress = document.createElement('div');"
        WScript.Echo "        progress.className = 'progress';"
        WScript.Echo "        progressBar.appendChild(progress);"
        WScript.Echo "        var progressText = document.createElement('div');"
        WScript.Echo "        progressText.className = 'progress-text';"
        WScript.Echo "        progressBar.appendChild(progressText);"
        WScript.Echo "        entry.appendChild(progressBar);"
        WScript.Echo "        var startTime = Date.now();"
        WScript.Echo "        var lastLoaded = 0;"
        WScript.Echo "        var formData = new FormData();"
        WScript.Echo "        formData.append('file', item);"
        WScript.Echo "        var xhr = new XMLHttpRequest();"
        WScript.Echo "        xhr.open('POST', '/upload?filename=' + encodeURIComponent(item.name) + '&redirectto=cloudexplorer/explorer.vbs%3Faction%3Dhandleupload%26dir%3D' + encodeURIComponent(dir));"
        WScript.Echo "        xhr.upload.addEventListener('progress', function(e) {"
        WScript.Echo "            if (e.lengthComputable) {"
        WScript.Echo "                var percentComplete = (e.loaded / e.total) * 100;"
        WScript.Echo "                progress.style.width = percentComplete + '%';"
        WScript.Echo "                progressText.innerHTML = Math.round(percentComplete) + '%';"
        WScript.Echo "                var elapsedTime = (Date.now() - startTime) / 1000; // in seconds"
        WScript.Echo "                var bytesUploaded = e.loaded - lastLoaded;"
        WScript.Echo "                lastLoaded = e.loaded;"
        WScript.Echo "                var remainingBytes = e.total - e.loaded;"
        WScript.Echo "                var uploadedMB = (e.loaded / (1024 * 1024)).toFixed(2);"
        WScript.Echo "                var remainingMB = (remainingBytes / (1024 * 1024)).toFixed(2);"
        WScript.Echo "                progressText.innerHTML += ' | Uploaded: ' + uploadedMB + ' MB | Remaining: ' + remainingMB + ' MB';"
        WScript.Echo "            }"
        WScript.Echo "        });"
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
        WScript.Echo "let wakeLock = null;"
        WScript.Echo "async function requestWakeLock(){"
        WScript.Echo "try {"
        WScript.Echo "wakeLock = await navigator.wakeLock.request('screen');"
        WScript.Echo "console.log('Wake Lock active.');"
        WScript.Echo "wakeLock.addEventListener('release', () => { console.log('Wake Lock was released'); }"
        WScript.Echo ");"
        WScript.Echo "} catch (err) { console.error(`${err.name}, ${err.message}`);}}"
        WScript.Echo "document.addEventListener('visibilitychange', () => {    if (wakeLock !== null && document.visibilityState === 'visible') {        requestWakeLock();    }});"
        WScript.Echo "requestWakeLock();"
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
