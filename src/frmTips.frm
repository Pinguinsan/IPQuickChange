VERSION 5.00
Begin VB.Form frmTips 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Quick Change Tips"
   ClientHeight    =   2160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4365
   Begin VB.CommandButton PreviousTipButton 
      Caption         =   "Previous Tip"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton NextTipButton 
      Caption         =   "Next Tip"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox TipsCheckBox 
      Caption         =   "Show These Tips At Startup"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Tip 
      Caption         =   "Tip: Common PLC domains include 192.168.1.X and 192.168.0.X, with X representing either 2, 10, or 20"
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Tip 
      Caption         =   $"frmTips.frx":0000
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Tip 
      Caption         =   "Tip: Press ""enter"" at any time to execute the change"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Tip 
      Caption         =   "Tip: Press ""d"" at any point to switch between static and dynamic IP settings"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private CurrentTip As Integer
Private NumberOfTips As Integer


Private Sub Form_KeyPress(KeyAscii As Integer)
    If ((KeyAscii = 27) Or (KeyAscii = 12)) Then 'Esc & Enter key
        Call OKButton_Click
    End If
End Sub

Private Sub NextTipButton_KeyPress(KeyAscii As Integer)
    If ((KeyAscii = 27) Or (KeyAscii = 12)) Then 'Esc & Enter key
        Call OKButton_Click
    End If
End Sub

Private Sub OKButton_KeyPress(KeyAscii As Integer)
    If ((KeyAscii = 27) Or (KeyAscii = 12)) Then 'Esc & Enter key
        Call OKButton_Click
    End If
End Sub

Private Sub Tip_KeyPress(KeyAscii As Integer)
    If ((KeyAscii = 27) Or (KeyAscii = 12)) Then 'Esc & Enter key
        Call OKButton_Click
    End If
End Sub


Private Sub PreviousTipButton_KeyPress(KeyAscii As Integer)
    If ((KeyAscii = 27) Or (KeyAscii = 12)) Then 'Esc & Enter key
        Call OKButton_Click
    End If
End Sub

Private Sub TipsCheckBox_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 27) Then 'Esc
        Call OKButton_Click
    ElseIf (KeyAscii = 12) Then
        TipsCheckBox.Value = 1
    End If
End Sub

Private Sub Form_Load()
    
    frmTips.Top = (Screen.Height - frmTips.Height) / 2
    frmTips.Left = (Screen.Width - frmTips.Width) / 2
    TipsCheckBox.Value = 1
    NumberOfTips = Tip.Count
    CurrentTip = 2
    Call ActivateOneTip(CurrentTip)
    PreviousTipButton.TabIndex = 0
    NextTipButton.TabIndex = 1
    TipsCheckBox.TabIndex = 2
    OKButton.TabIndex = 3

End Sub

Private Sub ActivateOneTip(ByVal TipToEnable As Integer)
Dim i As Integer

    For i = 0 To (NumberOfTips - 1)
        If (TipToEnable = i) Then
            Tip(i).Enabled = True
            Tip(i).Visible = True
        Else
            Tip(i).Enabled = False
            Tip(i).Visible = False
        End If
    Next i
    
End Sub

Private Sub PreviousTipButton_Click()

    If (CurrentTip = 0) Then
        CurrentTip = NumberOfTips - 1
    Else
        CurrentTip = CurrentTip - 1
    End If
    Call ActivateOneTip(CurrentTip)
End Sub

Private Sub NextTipButton_Click()

    If (CurrentTip = (NumberOfTips - 1)) Then
        CurrentTip = 0
    Else
        CurrentTip = CurrentTip + 1
    End If
    Call ActivateOneTip(CurrentTip)

End Sub

Private Sub OKButton_Click()

    If (TipsCheckBox.Value = 1) Then
        Me.Hide
        Me.Enabled = False
        Me.Visible = False
        frmMain.TipsCheckBox.Value = 1
    Else
        Open filePath For Output As #1
            Print #1, "noshowtips"
        Close #1
        Me.Hide
        Me.Enabled = False
        Me.Visible = False
        frmMain.TipsCheckBox.Value = 0
    End If
    
End Sub


