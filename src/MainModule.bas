Attribute VB_Name = "MainModule"
Option Explicit

Public filePath As String
Public ShowTipsOnStartup As Boolean
Dim gISIDE As Boolean

Public Function IsIDE() As Boolean
    IsIDE = False
    Debug.Assert CheckIDE
    If gISIDE Then
        IsIDE = True
    End If
End Function

Private Function CheckIDE() As Boolean
    gISIDE = True
    CheckIDE = True
End Function

Sub Main()
Dim InputString As String
    If (IsIDE) Then
        ChDrive (App.Path)
        ChDir (App.Path)
    End If
    
    'Call ChangeDirToApp
    filePath = "../configuration/config.txt"
    'filePath = "C:/Program Files/IPQuickChange/configuration/config.txt"
    If Dir(filePath) <> "" Then
        Open filePath For Input As #1
            InputString = StrConv(InputB(LOF(1), 1), vbUnicode)
        Close #1
        If (InStr(InputString, "no") = 0) Then
            ShowTipsOnStartup = True
            frmMain.Show
            frmTips.Show
        Else
            frmMain.Show
            ShowTipsOnStartup = False
        End If
    Else
        frmMain.Show
        ShowTipsOnStartup = False
    End If

End Sub
