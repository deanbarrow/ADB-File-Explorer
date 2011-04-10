VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Input/Output"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLog 
      Caption         =   "Log"
      Height          =   6255
      Left            =   7320
      TabIndex        =   9
      Top             =   120
      Width           =   4815
      Begin RichTextLib.RichTextBox rtbLog 
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   10398
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmDemo.frx":0000
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   4935
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   7095
      Begin RichTextLib.RichTextBox rtbOutput 
         Height          =   4575
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8070
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmDemo.frx":0082
      End
   End
   Begin VB.Frame fraSettings 
      Caption         =   "User Commands"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtCommand 
         Height          =   285
         Left            =   120
         MaxLength       =   512
         TabIndex        =   1
         Text            =   "cmd"
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         Height          =   255
         Left            =   4320
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   5640
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtWrite 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "help"
         Top             =   720
         Width           =   4095
      End
      Begin VB.CommandButton cmdWrite 
         Caption         =   "Write to Pipe"
         Default         =   -1  'True
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim WithEvents StdIO As cStdIO
Attribute StdIO.VB_VarHelpID = -1
Dim bExitAfterCancel As Boolean

Private Sub cmdCancel_Click()
    If StdIO.Ready = False Then
        AddLog "[" & Time & "] Canceling program.."
        StdIO.Cancel
    Else
        AddLog "[" & Time & "] Try executing a program first ;)"
    End If
End Sub

Private Sub cmdExecute_Click()
    If StdIO.Ready = True Then
        StdIO.CommandLine = txtCommand.Text
        AddLog "[" & Time & "] Executing command:"
        AddLog "-> " & StdIO.CommandLine
        rtbOutput.Text = ""
        StdIO.ExecuteCommand
        'Or simply StdIO.ExecuteCommand txtCommand.Text
    Else
        AddLog "[" & Time & "] Cannot execute command, already in use!"
    End If
End Sub

Private Sub cmdWrite_Click()
    Dim lBytesWritten As Long
    If StdIO.Ready = False Then
        lBytesWritten = StdIO.WriteData(txtWrite.Text)
        If lBytesWritten = -1 Then
            AddLog "[" & Time & "] Failed to write bytes to pipe!"
        Else
            AddLog "[" & Time & "] Successfully wrote " & lBytesWritten & " bytes to pipe!"
        End If
    Else
        AddLog "[" & Time & "] Try executing a program first ;)"
    End If
End Sub

Private Sub Form_Load()
    Set StdIO = New cStdIO
    txtCommand.Text = Environ("ComSpec")
    AddLog "[" & Time & "] Successfully loaded:"
    AddLog "-> " & StdIO.Version
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    If StdIO.Ready = True Then
        End
    Else
        bExitAfterCancel = True
        cmdCancel_Click
    End If
End Sub

Private Sub StdIO_CancelFail()
    AddLog "[" & Time & "] Cancel failed to end program. No longer reading pipes."
    DoEvents
    If bExitAfterCancel Then End
End Sub

Private Sub StdIO_CancelSuccess()
    AddLog "[" & Time & "] Cancel success! No longer reading pipes."
    DoEvents
    If bExitAfterCancel Then End
End Sub

Private Sub StdIO_Complete()
    AddLog "[" & Time & "] Complete!"
End Sub

Private Sub StdIO_Error(ByVal Number As Integer, ByVal Description As String)
    AddLog "[" & Time & "] Error #" & Number & ": " & Description
End Sub

Private Sub StdIO_GotData(ByVal Data As String)
    AddOutput Data
End Sub

Private Sub AddLog(ByVal strData As String)
    rtbLog.Text = rtbLog.Text & strData & vbNewLine
    rtbLog.SelStart = Len(rtbLog.Text) - 2 'Cause of vbNewLine
End Sub

Private Sub AddOutput(ByVal strData As String)
    rtbOutput.Text = rtbOutput.Text & strData
    rtbOutput.SelStart = Len(rtbOutput.Text)
End Sub
