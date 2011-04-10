VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAppManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "App Manager"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkData 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdUninstall 
      Caption         =   "Uninstall"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox lstApp 
      Height          =   2985
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   7215
   End
   Begin VB.ListBox lstStatus 
      Height          =   1230
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   7215
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      Caption         =   "Bulk"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Single"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "To uninstall, select from list below then click uninstall. To remove installation data, tick the box.."
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "Coming soon, management of system apps, plus hopefully 'friendly' app names"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   7215
   End
   Begin VB.Label lblData 
      Caption         =   "Remove Data"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblPath 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "Install a single app or bulk install from a directory"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmAppManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InstallType As String

Private Sub cmdDir_Click()
    
    '-- Initialize Common Dialog control
    With cdMain
        .Flags = cdlOFNPathMustExist
        .Flags = .Flags Or cdlOFNHideReadOnly
        .Flags = .Flags Or cdlOFNNoChangeDir
        .Flags = .Flags Or cdlOFNExplorer
        .Flags = .Flags Or cdlOFNNoValidate
        .FileName = "*.*"
    End With
    
    Dim x As Integer
    '-- Cheap way to use the common dialog box as a directory-picker
    x = 3

    cdMain.CancelError = True        'do not terminate on error

    On Error Resume Next         'I will hande errors

    cdMain.Action = 1              'Present "open" dialog

    '-- If FileTitle is null, user did not override the default (*.*)
    If cdMain.FileTitle <> "" Then x = Len(cdMain.FileTitle)

    If Err = 0 Then
        ChDrive cdMain.FileName
        lblPath.Caption = Left(cdMain.FileName, Len(cdMain.FileName) - x)
        cmdInstall.Enabled = True
        InstallType = "Dir"
    Else
      '-- User pressed "Cancel"
    End If
    
End Sub

Private Sub cmdFile_Click()

        
    '-- Initialize Common Dialog control
    With cdMain
        .Flags = cdlOFNPathMustExist
        .Flags = .Flags Or cdlOFNHideReadOnly
        .Flags = .Flags Or cdlOFNNoChangeDir
        .Flags = .Flags Or cdlOFNExplorer
        .Flags = .Flags Or cdlOFNNoValidate
        .FileName = "*.apk"
    End With

    Dim x As Integer
    '-- Cheap way to use the common dialog box as a directory-picker
    x = 3

    cdMain.CancelError = True        'do not terminate on error

    On Error Resume Next         'I will hande errors

    cdMain.Action = 1              'Present "open" dialog

    '-- If FileTitle is null, user did not override the default (*.*)
    If cdMain.FileTitle <> "" Then x = Len(cdMain.FileTitle)

    If Err = 0 Then
        ChDrive cdMain.FileName
        lblPath.Caption = cdMain.FileName
        cmdInstall.Enabled = True
        InstallType = "File"
    Else
      '-- User pressed "Cancel"
    End If

End Sub

Private Sub cmdInstall_Click()

    'Install APK's
    MsgBox "Installing, please wait.."
        
    If InstallType = "File" Then
        lstStatus.AddItem "Installing " & lblPath.Caption
        frmMain.ADB "install " & Chr(34) & lblPath.Caption & Chr(34)
        If Right(frmMain.ReturnData, 2) = "s)" Then
            lstStatus.AddItem "Done"
            lstStatus.ListIndex = lstStatus.ListCount - 1
        End If
    
    ElseIf InstallType = "Dir" Then
    
        Dim File As String
        Dim Count As Integer
        Count = 0
        If Right$(lblPath.Caption, 1) <> "\" Then lblPath.Caption = lblPath.Caption & "\"
        Extention = "*.apk"
        File = Dir$(lblPath.Caption & Extention)
        Do While Len(File)
            lstStatus.AddItem "Installing " & lblPath.Caption & File
            frmMain.ADB "install " & Chr(34) & lblPath.Caption & "\" & File & Chr(34)
            If Right(frmMain.ReturnData, 2) = "s)" Then
                lstStatus.AddItem "Done"
                lstStatus.ListIndex = lstStatus.ListCount - 1
                Count = Count + 1
            End If
            File = Dir$
        Loop
        
        MsgBox Count & " applications installed"
    
    End If

End Sub

Private Sub Form_Load()
        
    InstallType = ""
    
    cmdInstall.Enabled = False
    
    ListApps
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmMain.Show
End Sub

Private Sub ListApps()

Dim path As String
Dim Command As String
Command = ADBPath & " shell busybox ls -1 -A -L -p --color=never /"


'if adb fails then retry
redo:

frmMain.RunCommand Command

'MsgBox "returndata: " & frmMain.ReturnData

'Start ADB if not running
    If Left(ReturnData, 20) = "* daemon not running" Then
        'AddLog ("ADB server not running")
        'AddLog "Starting ADB server.."
        Shell ("cmd /C " & ADBPath & " start-server")
        'AddLog "Done"
        GoTo redo
    End If

lstApp.Clear
'lstDeviceFile.Clear
'lstApp.AddItem ("..")
'If Not Path = "/" Then lstApp.AddItem ("/")

        Dim myarray() As String
        ReDim myarray(1000) As String
        myarray() = Split(path, "/")

        Dim i As Integer
        For i = 0 To UBound(myarray)
                If Not myarray(i) = "" Then
'                    lstApp.AddItem (" | " & myarray(i))
                Else
'                    If lstApp.ListCount = 0 Then lstApp.AddItem ("/")
                End If
        Next i

'Catch errors
If ReturnData = "" Then Exit Sub

If Left(ReturnData, Len(ReturnData) - 2) = "error: device not found" Then
    MsgBox "Device not found. Please make sure your phone is connected."
    lstApp.Clear
    'lstDeviceFile.Clear
    ADBRemount = False
    Exit Sub
End If

'Dim myarray() As String
myarray() = Split(ReturnData, vbCrLf)
'MsgBox ReturnData

'Dim i As Integer
For i = 0 To UBound(myarray)
    If (i < UBound(myarray)) Then
        myarray(i) = Left(myarray(i), Len(myarray(i)) - 1)
        If myarray(i) = "" Then GoTo redo
        If Right(myarray(i), 1) = "/" Then lstApp.AddItem ("" & Left(myarray(i), Len(myarray(i)) - 1))
        If Not Right(myarray(i), 1) = "/" Then lstApp.AddItem (myarray(i))
        DoEvents
    End If
Next i


End Sub

Private Function Strip(psLine As String, psRemoveStr As String) As String

    Dim iLoc As Integer
    iLoc = InStr(psLine, psRemoveStr)
    Do While iLoc > 0
        psLine = Left(psLine, iLoc - 1) & _
              Mid(psLine, iLoc + Len(psRemoveStr))
        iLoc = InStr(psLine, psRemoveStr)
    Loop
    Strip = psLine

End Function

