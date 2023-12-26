' Author: gn027c
' Version: 4.1.0
' My bio: https://linktr.ee/gn027c

Dim response
Set shell = CreateObject("WScript.Shell")

Do
    text = InputBox("What content do you want to spam?", "[VALUE] - TEXT", "Enter the content you want to spam")
    If text = "" Then
        response = MsgBox("Incorrect, do you want to try again?", vbQuestion + vbYesNo, "[QUESTION] - SPAM AGAIN")
        If response = vbNo Then Exit Do
    Else
        response = vbYes
        Exit Do
    End If
Loop

If response <> vbNo Then
    Do
        times = InputBox("How many messages do you want to spam?", "[VALUE] - TIMES", "5")
        If times = "" Or Not IsNumeric(times) Then
            response = MsgBox("Incorrect, do you want to try again?", vbQuestion + vbYesNo, "[QUESTION] - SPAM AGAIN")
            If response = vbNo Then Exit Do
        Else
            response = vbYes
            Exit Do
        End If
    Loop
End If

If response <> vbNo Then
    Do
        speed = InputBox("How much speed do you want? (milliseconds unit).", "[VALUE] - SPEED", "500")
        If speed = "" Or Not IsNumeric(speed) Then
            response = MsgBox("Incorrect, do you want to try again?", vbQuestion + vbYesNo, "[QUESTION] - SPAM AGAIN")
            If response = vbNo Then Exit Do
        Else
            response = vbYes
            Exit Do
        End If
    Loop
End If

If response <> vbNo Then
    Do
        wait_s = InputBox("Delay in seconds to reach the spam area.", "[VALUE] - WAIT", "3")

        If wait_s = "" Or Not IsNumeric(wait_s) Then
            response = MsgBox("Incorrect, do you want to try again?", vbQuestion + vbYesNo, "[QUESTION] - SPAM AGAIN")
            If response = vbNo Then Exit Do
        Else
            response = vbYes
            Exit Do
        End If
    Loop
End If

if response <> vbNo then
    wait_mm = wait_s * 1000
    Do
    MsgBox "Wait for " & wait_s & " seconds to reach the spam area.", vbInformation, "[MESSAGE] - WAIT GO TO SPAM"
    WScript.Sleep wait_mm
        For i = 1 To times
            shell.SendKeys (text & "{enter}")
            WScript.Sleep speed
        Next
        WScript.Sleep 1000
        response = MsgBox("Do you want to spam again with the same content?", vbQuestion + vbYesNo, "[QUESTION] - SPAM AGAIN")
        If response = vbNo Then Exit Do
    Loop
End if

MsgBox "Spammer goes to sleep...", vbInformation, "[MESSAGE] - FINISH SPAM"