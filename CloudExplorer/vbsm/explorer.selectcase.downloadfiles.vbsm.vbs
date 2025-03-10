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
    