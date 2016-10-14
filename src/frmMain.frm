VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Settings Quick Change"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox TipsCheckBox 
      Caption         =   "Show Tips At Startup"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Domain 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Domain 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Domain 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Subnet 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Subnet 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Subnet 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox IP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox IP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox IP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton ExecuteButton 
      Caption         =   "Execute"
      Height          =   975
      Left            =   2760
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CheckBox ManualSetCheckBox 
      Caption         =   "Manually Set Subnet/Domain"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CheckBox StaticCheckBox 
      Caption         =   "Set IPv4 Settings To Static"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Domain 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Subnet 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox IP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label StaticInformativeLabel 
      Caption         =   $"frmMain.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   27
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   3360
      TabIndex        =   15
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   3360
      TabIndex        =   14
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   3360
      TabIndex        =   13
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label DomainLabel 
      Alignment       =   2  'Center
      Caption         =   "Default Domain (Or Leave Blank)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1200
      TabIndex        =   9
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label SubnetLabel 
      Alignment       =   2  'Center
      Caption         =   "Subnet Mask (Or Leave Blank)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   135
   End
   Begin VB.Label IPLabel 
      Alignment       =   2  'Center
      Caption         =   "Please Input IP Address "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PreviousIP(4) As String
Private PreviousSubnet(4) As String
Private PreviousDomain(4) As String
Private DeviceList() As String

Private Declare Function executeStaticChange Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" (ByVal IPAddress As String, ByVal SubnetMask As String, ByVal DefaultDomain As String) As Integer
Private Declare Function executeDynamicChange Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" () As Integer
Private Declare Function generateDeviceList Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" (ByVal IPBase As String, ByVal IPTail As Integer) As Integer
Private Declare Function parseDeviceListToFile Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" (ByVal IPBase As String) As Integer
Private Declare Function cleanUpTmp Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" () As Integer
Private Declare Function delay Lib "../libraries/ChangeIPSettings/Release/ChangeIPSettings.dll" (ByVal howLong As Double) As Integer

Private Sub ExecuteButton_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 100) Then 'D Key
        If (StaticCheckBox.Value = 0) Then
            StaticCheckBox.Value = 1
        Else
            StaticCheckBox.Value = 0
        End If
    ElseIf (KeyAscii = 27) Then 'Esc key
        Call Form_QueryUnload(0, 0)
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 13) Then 'Return Key
        ExecuteButton.Value = True
    ElseIf (KeyAscii = 100) Then 'D Key
        If (StaticCheckBox.Value = 0) Then
            StaticCheckBox.Value = 1
        Else
            StaticCheckBox.Value = 0
        End If
    ElseIf (KeyAscii = 27) Then 'Esc key
        Call Form_QueryUnload(0, 0)
    End If
        
End Sub

Private Sub TipsCheckBox_Click()

    If (TipsCheckBox.Value = 1) Then
        ShowTipsAtStartup = True
    Else
        ShowTipsAtStartup = False
    End If

End Sub

Private Sub TipsCheckBox_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 13) Then 'Return Key
        If (TipsCheckBox.Value = 1) Then
            TipsCheckBox.Value = 0
        Else
            TipsCheckBox.Value = 1
        End If
    ElseIf (KeyAscii = 100) Then 'D Key
        If (StaticCheckBox.Value = 0) Then
            StaticCheckBox.Value = 1
        Else
            StaticCheckBox.Value = 0
        End If
    ElseIf (KeyAscii = 27) Then 'Esc key
        Call Form_QueryUnload(0, 0)
    End If

End Sub


Private Sub Form_Load()
Dim i As Integer

    Set Me.Icon = LoadResPicture("APPICON", vbResIcon)
    frmMain.Top = (Screen.Height - frmMain.Height) / 2
    frmMain.Left = (Screen.Width - frmMain.Width) / 2
    
    For i = 0 To 15
        If (i < 4) Then
            IP(i).TabIndex = i
        ElseIf (i < 8) Then
            Subnet(i - 4).TabIndex = i
            Subnet(i - 4).Enabled = False
        ElseIf (i < 12) Then
            Domain(i - 8).TabIndex = i
            Domain(i - 8).Enabled = False
        ElseIf (i = 12) Then
            StaticCheckBox.TabIndex = i
        ElseIf (i = 13) Then
            ManualSetCheckBox.TabIndex = i
        ElseIf (i = 14) Then
            TipsCheckBox.TabIndex = i
        Else
            ExecuteButton.TabIndex = i
        End If
    Next i
    
    IP(0).Text = "192"
    IP(1).Text = "168"
    IP(2).Text = "1"
    IP(3).Text = "250"
    
    Subnet(0).Text = "255"
    Subnet(1).Text = "255"
    Subnet(2).Text = "255"
    Subnet(3).Text = "0"
    
    Domain(0).Text = "192"
    Domain(1).Text = "168"
    Domain(2).Text = "1"
    Domain(3).Text = "1"
    
    Call StoreLastSettings
        
    StaticCheckBox.Value = 1
    
    If (ShowTipsOnStartup) Then
        TipsCheckBox.Value = 1
    End If
    
    For i = 0 To 3
        Subnet(i).Enabled = False
        Domain(i).Enabled = False
    Next i
    

End Sub



Private Sub IP_LostFocus(Index As Integer)
Dim Temp As Integer
    
    If (Not IsNumeric(IP(Index).Text)) Then
        IP(Index).Text = PreviousIP(Index)
    Else
        Temp = CInt(IP(Index).Text)
        If ((Temp < 0) Or (Temp > 255)) Then
            IP(Index).Text = PreviousIP(Index)
        Else
            PreviousIP(Index) = IP(Index).Text
        End If
        
        If (ManualSetCheckBox.Value = 0) Then
            If (Index = 3) Then
                Domain(Index) = "1"
            Else
                Domain(Index) = IP(Index)
            End If
        End If
    End If
    
End Sub

Private Sub Subnet_LostFocus(Index As Integer)
Dim Temp As Integer
    
    If (Not IsNumeric(Subnet(Index).Text)) Then
        Subnet(Index).Text = PreviousSubnet(Index)
    Else
        Temp = CInt(Subnet(Index).Text)
        If ((Temp < 0) Or (Temp > 255)) Then
            Subnet(Index).Text = PreviousSubnet(Index)
        Else
            PreviousSubnet(Index) = Subnet(Index).Text
        End If
    End If
End Sub

Private Sub Domain_LostFocus(Index As Integer)
Dim Temp As Integer
    
    If (Not IsNumeric(Domain(Index).Text)) Then
        Domain(Index).Text = PreviousDomain(Index)
    Else
        Temp = CInt(Domain(Index).Text)
        If ((Temp < 0) Or (Temp > 255)) Then
            Domain(Index).Text = PreviousDomain(Index)
        Else
            PreviousDomain(Index) = Domain(Index).Text
        End If
    End If
End Sub

Private Sub IP_GotFocus(Index As Integer)
    
    Call SelectAllText(IP(Index))
    
End Sub

Private Sub Subnet_GotFocus(Index As Integer)
    
    Call SelectAllText(Subnet(Index))
    
End Sub

Private Sub Domain_GotFocus(Index As Integer)
    
    Call SelectAllText(Domain(Index))
    
End Sub

Private Sub IP_KeyPress(Index As Integer, KeyAscii As Integer)

    If (KeyAscii = 1) Then
        Call SelectAllText(IP(Index))
    ElseIf (KeyAscii = 13) Then 'Enter Key
        ExecuteButton.Value = True
    ElseIf (KeyAscii = 100) Then 'D Key
        If (StaticCheckBox.Value = 0) Then
            StaticCheckBox.Value = 1
        Else
            StaticCheckBox.Value = 0
        End If
    ElseIf (KeyAscii = 27) Then 'Esc key
        Call Form_QueryUnload(0, 0)
    End If
    
    
End Sub

Private Sub Subnet_KeyPress(Index As Integer, KeyAscii As Integer)

    If (KeyAscii = 1) Then
        Call SelectAllText(Subnet(Index))
    ElseIf (KeyAscii = 13) Then 'Enter Key
        ExecuteButton.Value = True
    ElseIf (KeyAscii = 68) Then 'D Key
        If (StaticCheckBox.Value = 0) Then
            StaticCheckBox.Value = 1
        Else
            StaticCheckBox.Value = 0
        End If
    ElseIf (KeyAscii = 27) Then 'Esc key
        Call Form_QueryUnload(0, 0)
    End If
    
End Sub

Private Sub Domain_KeyPress(Index As Integer, KeyAscii As Integer)

    If (KeyAscii = 1) Then
        Call SelectAllText(Domain(Index))
    ElseIf (KeyAscii = 13) Then 'Enter Key
        ExecuteButton.Value = True
    ElseIf (KeyAscii = 100) Then 'D Key
        If (StaticCheckBox.Value = 0) Then
            StaticCheckBox.Value = 1
        Else
            StaticCheckBox.Value = 0
        End If
    ElseIf (KeyAscii = 27) Then 'Esc key
        Call Form_QueryUnload(0, 0)
    End If
    
End Sub

Private Sub StaticCheckBox_Click()
Dim i As Integer
    If (StaticCheckBox.Value = 0) Then
    
        Call StoreLastSettings
        
        ManualSetCheckBox.Value = 0
        
        For i = 0 To 3
            IP(i).Enabled = False
            IP(i).Text = "Auto"
            
            Subnet(i).Enabled = False
            Subnet(i).Text = "Auto"
            
            Domain(i).Enabled = False
            Domain(i).Text = "Auto"
        Next i
        
        ManualSetCheckBox.Enabled = False
        ExecuteButton.Caption = "Execute Dynamic"
        
    Else
    
        For i = 0 To 3
            IP(i).Enabled = True
            Subnet(i).Enabled = True
            Domain(i).Enabled = True
        Next i
            
        ManualSetCheckBox.Enabled = True
        ExecuteButton.Caption = "Execute Static"
        
        Call RestoreLastSettings
    
    End If
    
End Sub

Private Sub StaticCheckBox_KeyPress(KeyAscii As Integer)
    
     If (StaticCheckBox.Value = 0) Then
        If (KeyAscii = 13) Then
            StaticCheckBox.Value = 1
        End If
    Else
        If (KeyAscii = 13) Then
            StaticCheckBox.Value = 0
        End If
    End If
    If (KeyAscii = 100) Then 'D Key
        If (StaticCheckBox.Value = 0) Then
            StaticCheckBox.Value = 1
        Else
            StaticCheckBox.Value = 0
        End If
    ElseIf (KeyAscii = 27) Then 'Esc key
        Call Form_QueryUnload(0, 0)
    End If

End Sub

Private Sub ManualSetCheckBox_KeyPress(KeyAscii As Integer)
    
     If (ManualSetCheckBox.Value = 0) Then
        If (KeyAscii = 13) Then
            ManualSetCheckBox.Value = 1
        End If
    ElseIf (KeyAscii = 13) Then
        ManualSetCheckBox.Value = 0
    ElseIf (KeyAscii = 27) Then 'Esc key
        Call Form_QueryUnload(0, 0)
    End If

End Sub

Private Sub ManualSetCheckBox_Click()
Dim i As Integer

If (ManualSetCheckBox.Value = 0) Then
    
    For i = 0 To 3
        Subnet(i).Enabled = False
        Domain(i).Enabled = False
    Next i

    Subnet(0).Text = "255"
    Subnet(1).Text = "255"
    Subnet(2).Text = "255"
    Subnet(3).Text = "0"

    Domain(0).Text = IP(0).Text
    Domain(1).Text = IP(1).Text
    Domain(2).Text = IP(2).Text
    Domain(3).Text = "1"
    
Else
    For i = 0 To 3
        Subnet(i).Enabled = True
        Domain(i).Enabled = True
    Next i
End If

End Sub

Private Sub StoreLastSettings()

    For i = 0 To 3
        If (IsNumeric(IP(i).Text)) Then
            PreviousIP(i) = IP(i).Text
        End If
        If (IsNumeric(Subnet(i).Text)) Then
            PreviousSubnet(i) = Subnet(i).Text
        End If
        If (IsNumeric(Domain(i).Text)) Then
            PreviousDomain(i) = Domain(i).Text
        End If
    Next i

End Sub

Private Sub RestoreLastSettings()

    For i = 0 To 3
        IP(i).Text = PreviousIP(i)
        Subnet(i).Text = PreviousSubnet(i)
        Domain(i).Text = PreviousDomain(i)
    Next i

End Sub

Private Sub ExecuteButton_Click()

Dim IPTemp As String
Dim SubnetTemp As String
Dim DomainTemp As String
Dim CReturn As Integer
Dim MsgBoxMsg As String
Dim MsgBoxReturn As Integer
Dim IPBaseString As String
Dim IPTail As Integer
Dim IntCount As Integer
Dim ReturnVal As Integer
Dim IPCountString As String
Dim IPCount As Integer
Dim IPString() As String
Dim i As Integer

    If (StaticCheckBox.Value = 1) Then
        IPTemp = IP(0).Text & "." & IP(1).Text & "." & IP(2).Text & "." & IP(3).Text
        SubnetTemp = Subnet(0).Text & "." & Subnet(1).Text & "." & Subnet(2).Text & "." & Subnet(3).Text
        DomainTemp = Domain(0).Text & "." & Domain(1).Text & "." & Domain(2).Text & "." & Domain(3).Text
        CReturn = executeStaticChange(IPTemp, SubnetTemp, DomainTemp)
        If (CReturn = 0) Then
            CReturn = delay(2.5)
            MsgBoxMsg = "Successfully changed IP settings to static:" & vbNewLine & "IP Address: " & IPTemp & vbNewLine & "Subnet Mask: " & SubnetTemp & vbNewLine & "Default Domain: " & DomainTemp & vbNewLine & "Would you like to generate a list of devices on this network?"
            MsgBoxReturn = MsgBox(MsgBoxMsg, vbYesNoCancel, "IP Settings Quick Change Result")
            If (MsgBoxReturn = vbYes) Then
                IPBaseString = IP(0).Text & "." & IP(1).Text & "." & IP(2).Text & "."
                IPTail = CInt(IP(3).Text)
                ReturnVal = generateDeviceList(IPBaseString, IPTail)
                If (ReturnVal = 0) Then
                    ReturnVal = parseDeviceListToFile(IPBaseString)
                    If (ReturnVal = 0) Then
                        i = 0
                        Open "../tmp/parsedips.txt" For Input As #2
                            Line Input #2, IPCountString
                            IPCount = CInt(IPCountString)
                            ReDim IPString(IPCount) As String
                            For i = 0 To (IPCount - 1)
                                Line Input #2, IPString(i)
                            Next i
                        Close #2
                    
                        If (IPString(0) = "0") Then
                            MsgBoxReturn = MsgBox("No devices were found on this network", vbOKOnly, "Device List")
                        Else
                            MsgBoxMsg = ""
                            For i = 0 To UBound(IPString)
                                MsgBoxMsg = MsgBoxMsg & IPString(i) & vbNewLine
                            Next i
                            MsgBoxReturn = MsgBox(MsgBoxMsg, vbOKOnly, "Device List")
                        End If
                    ElseIf (ReturnVal = 1) Then
                        MsgBoxReturn = MsgBox("Something went wrong while parsing IP results, a device list cannot be generated.", vbOKOnly, "Device List")
                    ElseIf (ReturnVal = 2) Then
                        MsgBoxReturn = MsgBox("Something went wrong while trying to open the ping results file for reading, a device list cannot be generated.", vbOKOnly, "Device List")
                    Else
                        MsgBoxReturn = MsgBox("Something went wrong while trying to open the file to write the successful pings, a device list cannot be generated.", vbOKOnly, "Device List")
                    End If
                Else
                    MsgBoxReturn = MsgBox("Something went wrong while trying to ping network, a device list cannot be generated.", vbOKOnly, "Device List")
                End If
            End If
        Else
            MsgBoxMsg = "Something went wrong, settings have not been changed. Please see log file"
            MsgBoxReturn = MsgBox(MsgBoxMsg, vbOKOnly, "IP Settings Quick Change Result")
        End If
    Else
        CReturn = executeDynamicChange()
        If (CReturn = 0) Then
            MsgBoxMsg = "Successfully changed IP settings to dynamic"
            MsgBoxReturn = MsgBox(MsgBoxMsg, vbOKOnly, "IP Settings Quick Change Result")
        Else
            MsgBoxMsg = "Something went wrong, settings have not been changed. Please see log file"
            MsgBoxReturn = MsgBox(MsgBoxMsg, vbOKOnly, "IP Settings Quick Change Result")
        End If
    End If
End Sub

Private Sub SelectAllText(TextBoxToSelect As TextBox)

    TextBoxToSelect.SelStart = 0
    TextBoxToSelect.SelLength = Len(TextBoxToSelect.Text)

End Sub

Private Sub ParseStringToArray(ByVal StringToParse As String)

    DeviceList = Split(StringToParse, ",")

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim MsgBoxReturn As Integer
Dim CReturn As Integer
    
    MsgBoxReturn = MsgBox("Are you sure you want to quit?", vbYesNo, "Quit Confirmation")
    If (MsgBoxReturn = vbYes) Then
        If (TipsCheckBox.Value = 1) Then
            Open filePath For Output As #1
                Print #1, "showtips"
            Close #1
        Else
            Open filePath For Output As #1
                Print #1, "noshowtips"
            Close #1
        End If
        CReturn = cleanUpTmp
        End
    End If

End Sub


