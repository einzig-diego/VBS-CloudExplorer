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