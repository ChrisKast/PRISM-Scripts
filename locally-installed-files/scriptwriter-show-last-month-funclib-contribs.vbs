last_month_question = MsgBox("Do you want only the last month details only?", vbYesNoCancel)

If last_month_question = vbCancel then stopscript

If last_month_question = vbYes then
    CreateObject("WScript.Shell").Run("https://github.com/MN-Script-Team/BZS-FuncLib/graphs/contributors?from=" & dateadd("m", -1, date) & "&to=" & date & "&type=c")
Else
    CreateObject("WScript.Shell").Run("https://github.com/MN-Script-Team/BZS-FuncLib/graphs/contributors?type=c")
End if
