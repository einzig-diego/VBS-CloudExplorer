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
     