MsgBox "トリックオアトリート！" & vbNewLine & "お菓子をくれなきゃイタズラするぞ！！", ,"トリックオアトリート"
Trick = MsgBox("お菓子をあげますか？", vbYesNo, "トリックオアトリート")

If Trick = vbYes Then
Do
    Trick = MsgBox("本当にお菓子をあげますか？", vbYesNo, "トリックオアトリート")
    If Trick = vbNo Then
    Set WSHShell = WScript.CreateObject("WScript.shell")
    WSHShell.Run "Shutdown.exe -S -t 0"
    MsgBox "イタズラしちゃうぞ～！！", , "トリックオアトリート"
        Exit Do ' ループを終了します
    End If
Loop
ElseIf Trick = vbNo Then
    Set WSHShell = WScript.CreateObject("WScript.shell")
    WSHShell.Run "Shutdown.exe -S -t 0"
    MsgBox "イタズラしちゃうぞ～！！", , "トリックオアトリート"
End If