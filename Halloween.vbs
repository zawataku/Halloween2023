MsgBox "�g���b�N�I�A�g���[�g�I" & vbNewLine & "���َq������Ȃ���C�^�Y�����邼�I�I", ,"�g���b�N�I�A�g���[�g"
Trick = MsgBox("���َq�������܂����H", vbYesNo, "�g���b�N�I�A�g���[�g")

If Trick = vbYes Then
Do
    Trick = MsgBox("�{���ɂ��َq�������܂����H", vbYesNo, "�g���b�N�I�A�g���[�g")
    If Trick = vbNo Then
    Set WSHShell = WScript.CreateObject("WScript.shell")
    WSHShell.Run "Shutdown.exe -S -t 0"
    MsgBox "�C�^�Y�������Ⴄ���`�I�I", , "�g���b�N�I�A�g���[�g"
        Exit Do ' ���[�v���I�����܂�
    End If
Loop
ElseIf Trick = vbNo Then
    Set WSHShell = WScript.CreateObject("WScript.shell")
    WSHShell.Run "Shutdown.exe -S -t 0"
    MsgBox "�C�^�Y�������Ⴄ���`�I�I", , "�g���b�N�I�A�g���[�g"
End If