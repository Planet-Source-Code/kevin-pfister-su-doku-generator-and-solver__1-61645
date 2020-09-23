Attribute VB_Name = "MOD_Functions"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_PASTE = &H302


Public Sub AddText(Text)
    Text = Replace(Text, Chr(13), ":")
    With FrmSudoku.RTBVerbose
        DoEvents
        .SelStart = Len(.Text)
        .SelColor = vbBlack
        If .SelStart <> 0 Then
            .SelText = vbCrLf
            If Mid(Text, 1, 3) = "ERR" Then
                Clipboard.Clear
                Clipboard.SetData FrmSudoku.PicErr.Picture
                .Locked = False
                SendMessage .hwnd, WM_PASTE, 0, 0
                Clipboard.Clear
                .Locked = True
                Text = Mid(Text, 4)
            End If
            If Mid(Text, 1, 3) = "FIN" Then
                Clipboard.Clear
                Clipboard.SetData FrmSudoku.PicFind.Picture
                .Locked = False
                SendMessage .hwnd, WM_PASTE, 0, 0
                Clipboard.Clear
                .Locked = True
                Text = Mid(Text, 4)
            End If
            .SelText = .SelText & Text
        Else
            .SelText = Text
        End If
        .SelStart = Len(.Text)
        DoEvents
    End With
End Sub
