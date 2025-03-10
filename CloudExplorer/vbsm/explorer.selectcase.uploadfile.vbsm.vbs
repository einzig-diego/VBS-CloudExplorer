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
        