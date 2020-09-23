VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   810
      Width           =   2715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const CB_SHOWDROPDOWN = &H14F
Dim PSTR As String
Private Sub Combo1_GotFocus()
    SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, True, 1
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 8 And (KeyCode < 35 Or KeyCode > 40) And KeyCode <> 46 And KeyCode <> 13 Then
        PSTR = Combo1.Text
        For i = 0 To Combo1.ListCount - 1
            If StrComp(PSTR, (Left(Combo1.List(i), Len(PSTR))), vbTextCompare) = 0 Then
                Combo1.ListIndex = i
            Exit For
            End If
        Next i
        Combo1.SelStart = Len(PSTR)
        Combo1.SelLength = Len(Combo1.Text) - Len(PSTR)
    End If
End Sub

Private Sub Form_Load()
    With Combo1
        .AddItem "ABCD"
        .AddItem "AEFG"
        .AddItem "ACFG"
        .AddItem "AFGH"
        .AddItem "AGHI"
        .AddItem "bkuy"
        .AddItem "KIJNS"
        .AddItem "JHD"
        .AddItem "HGFASD"
        .AddItem "TREKJ"
        .AddItem "ZXNBVC"
        .AddItem "QWEIU"
        
    End With
Combo1.Text = ""

End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Or (KeyAscii = 32 And Len(Combo1.Text) = 0) Then
        SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, True, 1
    ElseIf KeyAscii = 13 Then
        SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, 0, 1
    End If
    If KeyAscii = 32 And Len(Combo1.Text) = 0 Then KeyAscii = 0
    
End Sub
