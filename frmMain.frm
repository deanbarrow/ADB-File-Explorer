VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADB File Explorer"
   ClientHeight    =   8160
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMove 
      Caption         =   "MV"
      Height          =   495
      Left            =   5040
      TabIndex        =   15
      ToolTipText     =   "Remove"
      Top             =   2520
      Width           =   615
   End
   Begin VB.ListBox lstDeviceFile 
      Height          =   3960
      Left            =   5760
      TabIndex        =   14
      Top             =   2160
      Width           =   4815
   End
   Begin VB.DriveListBox driDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.DirListBox dirDrive 
      Height          =   1665
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   4815
   End
   Begin VB.FileListBox filDrive 
      Height          =   3990
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   4815
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "R"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      ToolTipText     =   "Refresh"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdRm 
      Caption         =   "-"
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      ToolTipText     =   "Remove"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdMkdir 
      Caption         =   "+"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   "Make Directory"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdShell 
      Caption         =   "Shell"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   3720
      Width           =   615
   End
   Begin VB.ListBox lstLog 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   10455
   End
   Begin VB.CommandButton cmdRemount 
      Caption         =   "Remount"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdPull 
      Caption         =   "<"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "Pull"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdPush 
      Caption         =   ">"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      ToolTipText     =   "Push"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtDevice 
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox txtDrive 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.ListBox lstDeviceDir 
      Height          =   1620
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuInfo 
         Caption         =   "&Device Info"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuApp 
         Caption         =   "&App Manager"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Index           =   1
      Begin VB.Menu mnuWebsite 
         Caption         =   "&Website"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Register variables
Dim AppVersion As String
Dim ADBPath As String
Dim WithEvents StdIO As cStdIO
Attribute StdIO.VB_VarHelpID = -1
Dim bExitAfterCancel As Boolean
Public ReturnData As String
Dim DrivePath As String
Dim DevicePath As String
Dim ADBRemount As Boolean
Dim ActivePane As String
Dim ActiveDrive As String
Dim ActiveDevice As String


Public Sub AddLog(Log As String, Optional NoTime As Boolean)

    If Not NoTime = True Then Log = "[" & Time & "] " & Log
    lstLog.AddItem Log
    lstLog.ListIndex = lstLog.ListCount - 1

End Sub


Public Sub ADB(Command As String)

    Command = ADBPath & " " & Command

    RunCommand Command
    
    If Not ReturnData = "" Then ReturnData = Left(ReturnData, Len(ReturnData) - 2)
    'AddLog ReturnData
    
End Sub

Public Sub DOS(Command As String)

    Command = "cmd /C" & " " & Command

    RunCommand Command
    
    AddLog ReturnData
    
End Sub

Private Sub cmdMkdir_Click()
    
    'Make directory in active pane
    Dim Command As String
    Command = ""

    If ActivePane = "" Then
        MsgBox "Please select an item first!"
        Exit Sub
    ElseIf ActivePane = "Drive" Then
        Command = InputBox("Enter new dir name: ", "Add Directory", DrivePath & "\")
        If Command = "" Then
            'Cancelled
            Exit Sub
        End If
        If Left(Command, 1) = "/" Then
            MsgBox "Cannot create an Android path in Windows!"
            Exit Sub
        End If
        If Not Command = "" Then
            'Command = Chr(34) & Command & Chr(34)
            MkDir Command
            filDrive.Refresh
            dirDrive.Refresh
        End If
    ElseIf ActivePane = "Device" Then
        If Not DevicePath = "/" Then DevicePath = DevicePath & "/"
        Command = InputBox("Enter new dir name: ", "Add Directory", DevicePath)
        If Command = "" Then
            'Cancelled
            Exit Sub
        End If
        If Not Left(Command, 1) = "/" Then
            MsgBox "Cannot create a Windows path in Android!"
            Exit Sub
        End If
        If Not Command = "" Then
            ADB "shell busybox mkdir " & Chr(34) & Command & Chr(34)
            ListDevice DevicePath
        End If
    End If
    
End Sub

Private Sub cmdMove_Click()

 Dim Command As String
 
    If ActivePane = "" Then
        MsgBox "Please select an item first!"
        Exit Sub
    ElseIf ActivePane = "Drive" Then
        
        If Len(DrivePath) = 3 Then
            MsgBox "Cannot rename " & DrivePath & "!"
            Exit Sub
        End If
        
        If ActiveDrive = "Dir" Then
            Command = InputBox("Rename/move dir: ", "Rename/Move", DrivePath)
                If Command = "" Then
                    'Cancelled
                    Exit Sub
                End If
                If Left(Command, 1) = "/" Then
                    MsgBox "Cannot create an Android path in Windows!"
                    Exit Sub
                End If
                If Not Command = "" Then
                    'Command = Chr(34) & Command & Chr(34)
                    Name DrivePath As Command
                    dirDrive.path = Command
                    filDrive.Refresh
                    dirDrive.Refresh
                End If
        ElseIf ActiveDrive = "File" Then
            Command = InputBox("Rename/move dir: ", "Rename/Move", DrivePath & "\" & filDrive.FileName)
                If Command = "" Then
                    'Cancelled
                    Exit Sub
                End If
                If Left(Command, 1) = "/" Then
                    MsgBox "Cannot create an Android path in Windows!"
                    Exit Sub
                End If
                If Not Command = "" Then
                    'Command = Chr(34) & Command & Chr(34)
                    Name DrivePath & "\" & filDrive.FileName As Command
                    filDrive.Refresh
                    dirDrive.Refresh
                End If
        End If
    

    ElseIf ActivePane = "Device" Then
        
        If DevicePath = "/" Then
            MsgBox "Cannot move /!"
            Exit Sub
        End If
        
        If ActiveDevice = "Dir" Then
        
            If Not DevicePath = "/" Then DevicePath = DevicePath
            Command = InputBox("Rename/move dir: ", "Rename/Move", DevicePath)
            If Command = "" Then
                'Cancelled
                Exit Sub
            End If
            If Not Left(Command, 1) = "/" Then
                MsgBox "Cannot create a Windows path in Android!"
                Exit Sub
            End If
            If Not Command = "" Then
                ADB "shell busybox mv " & DevicePath & " " & Chr(34) & Command & Chr(34)
                ListDevice Command
            End If
        
        ElseIf ActiveDevice = "File" Then
        
            If Not DevicePath = "/" Then DevicePath = DevicePath
            Command = InputBox("Rename/move dir: ", "Rename/Move", DevicePath & "/" & lstDeviceFile.Text)
            If Command = "" Then
                'Cancelled
                Exit Sub
            End If
            If Not Left(Command, 1) = "/" Then
                MsgBox "Cannot create a Windows path in Android!"
                Exit Sub
            End If
            If Not Command = "" Then
                ADB "shell busybox mv " & Chr(34) & DevicePath & "/" & lstDeviceFile.Text & Chr(34) & " " & Chr(34) & Command & Chr(34)
                ListDevice DevicePath
            End If
        
        End If
        
        
    End If

End Sub

Private Sub cmdPull_Click()
    
    Dim PullPath As String
    
    If ActivePane = "" Or ActiveDrive = "" Or ActiveDevice = "" Then
        MsgBox "Please select an item first!"
        Exit Sub
    ElseIf ActiveDevice = "File" Then
    'Push file
        If DevicePath = "/" Then DevicePath = ""
        PullPath = DevicePath & "/" & lstDeviceFile.Text
        Dim RemoveSlash As String
        RemoveSlash = Right(Left(PullPath, 4), 2)
        
        If RemoveSlash = "//" Then
            PullPath = Left(PullPath, 3) & Right(PullPath, Len(PullPath) - 4)
        End If
    
    ElseIf ActiveDevice = "Dir" Then
    'Push directory
        PullPath = DevicePath
    
    End If
    
    
        ' Push file from drive to device
    Dim FromFile As String
    Dim ToFile As String
    FromFile = Chr(34) & PullPath & Chr(34)
    'If Not DevicePath = "/" Then ToFile = DevicePath & "/"
    
    Dim NewDirDrive As String
    
    'if pushing dir then add dir to tofile
    If ActiveDevice = "Dir" Then
                
        'Now we need to convert C:\file\dir to dir
        Dim myarray() As String
        ReDim myarray(1000) As String
        myarray() = Split(PullPath, "/")
        PullPath = ""

        Dim i As Integer
        For i = 0 To UBound(myarray)
            If myarray(i) = "" Then
                'i = UBound(myarray)
            Else
                NewDirDrive = myarray(i)
            End If
        Next i
        
        'If Len(FromFile) = 3 Then
            'If MsgBox("You are about to pull the entire device's contents." & vbCrLf & vbCrLf & "Continue?", vbYesNo, "Confirmation") = vbYes Then
            '    NewDirDrive = "root"
            'Else
            '    Exit Sub
            'End If
        'End If
                
    End If
    
    If NewDirDrive = "" Then
        ToFile = DrivePath
    Else
        If DevicePath = "/" Then DevicePath = ""
        ToFile = DrivePath & "\" & NewDirDrive
    End If
    
    ToFile = Chr(34) & ToFile & Chr(34)
    
    Dim Command As String
    Command = "pull " & FromFile & " " & ToFile
    
    If MsgBox("Will pull:" & vbCrLf & FromFile & " to " & ToFile & vbCrLf & vbCrLf & "Continue?", vbYesNo, "Confirmation") = vbYes Then
    ADB Command
    End If
    
    filDrive.Refresh
    dirDrive.Refresh
    
End Sub

Private Sub cmdPush_Click()
    
    Dim PushPath As String
    
    If ActivePane = "" Or ActiveDrive = "" Or ActiveDevice = "" Then
        MsgBox "Please select an item first!"
        Exit Sub
    ElseIf ActiveDrive = "File" Then
    'Push file
        PushPath = filDrive.path & "\" & filDrive.FileName
        Dim RemoveSlash As String
        RemoveSlash = Right(Left(PushPath, 4), 2)
        
        If RemoveSlash = "\\" Then
            PushPath = Left(PushPath, 3) & Right(PushPath, Len(PushPath) - 4)
        End If
    
    ElseIf ActiveDrive = "Dir" Then
    'Push directory
        PushPath = filDrive.path
    
    End If
    
    
    ' Push file from drive to device
    Dim FromFile As String
    Dim ToFile As String
    FromFile = Chr(34) & PushPath & Chr(34)
    'If Not DevicePath = "/" Then ToFile = DevicePath & "/"
    
    Dim NewDirDevice As String
    
    'if pushing dir then add dir to tofile
    If ActiveDrive = "Dir" Then
                
        'Now we need to convert C:\file\dir to dir
        Dim myarray() As String
        ReDim myarray(1000) As String
        myarray() = Split(PushPath, "\")
        PushPath = ""

        Dim i As Integer
        For i = 0 To UBound(myarray)
            If myarray(i) = "" Then
                i = UBound(myarray)
            Else
                NewDirDevice = myarray(i)
            End If
        Next i
        
        If Len(FromFile) = 5 Then
            MsgBox "Cannot push " & FromFile
            Exit Sub
        End If
                
    End If
    
    If NewDirDevice = "" Then
        ToFile = DevicePath
    Else
        If DevicePath = "/" Then DevicePath = ""
        ToFile = DevicePath & "/" & NewDirDevice
    End If
    
    ToFile = Chr(34) & ToFile & Chr(34)
    
    If ActiveDevice = "File" Then ToFile = ToFile & "/"
    
    Dim Command As String
    Command = "push " & FromFile & " " & ToFile
        
    If MsgBox("Will push:" & vbCrLf & FromFile & " to " & ToFile & vbCrLf & vbCrLf & "Continue?", vbYesNo, "Confirmation") = vbYes Then
    ADB Command
    End If
    
    ListDevice (DevicePath)
    
End Sub

Private Sub cmdRefresh_Click()
    
    filDrive.Refresh
    dirDrive.Refresh
    ListDevice DevicePath

End Sub

Private Sub cmdRemount_Click()

    If ADBRemount = False Then
    
        ' ADB Remount
        AddLog ("Mounting filesystem as read/write..")
                
        ADB "remount"
    
        ADBRemount = True
        
    Else
    
        AddLog "Filesystem already mounted as read/write"
        
    End If

End Sub


Private Sub cmdRm_Click()

    If ActivePane = "" Then
        MsgBox "Please select an item first!"
        Exit Sub
    ElseIf ActivePane = "Drive" Then
        
        If Len(DrivePath) = 3 Then
            MsgBox "Cannot delete " & DrivePath & "!"
            Exit Sub
        End If
        
        If ActiveDrive = "Dir" Then
            If MsgBox("Are you sure you want to delete:" & vbCrLf & DrivePath, vbYesNo, "Confirmation") = vbYes Then
                AddLog "Removing " & DrivePath
                Shell "cmd.exe /C rmdir /S /Q " & Chr(34) & DrivePath & Chr(34)
                Dim myarray() As String
                ReDim myarray(1000) As String
                myarray() = Split(DrivePath, "\")
                DrivePath = ""
                Dim i As Integer
                For i = 0 To UBound(myarray) - 1
                    If (i < UBound(myarray)) Then
                        DrivePath = DrivePath & myarray(i) & "\"
                    End If
                Next i
                dirDrive.path = DrivePath
                dirDrive.Refresh
                filDrive.Refresh
                AddLog "Done"
                Sleep 500
                dirDrive.Refresh
            End If
        ElseIf ActiveDrive = "File" Then
            If MsgBox("Are you sure you want to delete:" & vbCrLf & DrivePath & "\" & filDrive.FileName, vbYesNo, "Confirmation") = vbYes Then
                AddLog "Removing " & DrivePath & "\" & filDrive.FileName
                Kill DrivePath & "\" & filDrive.FileName
                dirDrive.Refresh
                filDrive.Refresh
                AddLog "Done"
            End If
        End If
    

    ElseIf ActivePane = "Device" Then
        
        If DevicePath = "/" Then
            MsgBox "Cannot delete /!"
            Exit Sub
        End If
        
        If ActiveDevice = "Dir" Then
        
            If MsgBox("Are you sure you want to delete:" & vbCrLf & DevicePath, vbYesNo, "Confirmation") = vbYes Then
                AddLog "Removing " & DevicePath
                ADB "shell busybox rm -rf " & Chr(34) & DevicePath & Chr(34)
                'Dim myarray() As String
                ReDim myarray(1000) As String
                myarray() = Split(DevicePath, "/")
                DevicePath = ""
                'Dim i As Integer
                For i = 0 To UBound(myarray) - 1
                    If (i < UBound(myarray)) Then
                        DevicePath = DevicePath & myarray(i) & "/"
                    End If
                Next i
                ListDevice DevicePath
                AddLog "Done"
            End If
        
        ElseIf ActiveDevice = "File" Then
        
            If MsgBox("Are you sure you want to delete:" & vbCrLf & DevicePath & "/" & lstDeviceFile.Text, vbYesNo, "Confirmation") = vbYes Then
                AddLog "Removing " & DevicePath & "/" & lstDeviceFile.Text
                ADB "shell busybox rm -rf " & Chr(34) & DevicePath & "/" & lstDeviceFile.Text & Chr(34)
                ListDevice DevicePath
                AddLog "Done"
            End If
        
        End If
        
        
    End If

    
End Sub

Private Sub cmdShell_Click()
    
    Dim ADBCommand As String
    ADBCommand = InputBox("Please enter a shell command to perform", "Shell")
    If Not ADBCommand = "" Then
        ADB ("shell " & ADBCommand)
        MsgBox ReturnData
    End If
    
End Sub

Private Sub dirDrive_Change()
    
    DrivePath = dirDrive.path
    filDrive.path = DrivePath
    txtDrive.Text = DrivePath
    
        
End Sub

Private Sub dirDrive_GotFocus()

    ActivePane = "Drive"
    ActiveDrive = "Dir"
    
End Sub

Private Sub dirDrive_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Hit enter to open path
    'If KeyCode = "13" Then MsgBox dirDrive.Path = dirDrive.*selection*
    
    
End Sub

Private Sub driDrive_Change()
    

    On Error Resume Next
    DrivePath = driDrive.Drive
    If Left(DrivePath, 2) = "c:" Then DrivePath = "C:\"
    dirDrive.path = DrivePath

End Sub

Private Sub filDrive_GotFocus()
    
    ActivePane = "Drive"
    ActiveDrive = "File"
    
End Sub

Private Sub Form_Load()

    'Used to testing
    'frmDemo.Show

    ' Set variables
    AppVersion = "v" & App.Major & "." & App.Minor
    ADBPath = "adb.exe -d"
    ADBRemount = False
    ActivePane = ""
    
    ' Default stores
    DrivePath = "C:\"
    DevicePath = "/"
    'txtDrive.Text = txtDrive.Text
    txtDevice.Text = txtDevice.Text
    dirDrive.path = DrivePath
    filDrive.path = DrivePath
    
    Set StdIO = New cStdIO
    'AddLog "[" & Time & "] Successfully loaded:"
    'AddLog "-> " & StdIO.Version
        
    AddLog "  ADB File Explorer " & AppVersion, True
    
    AddLog "If device shows no files, click Refresh ('R')", True
        
    ' Load default stores
    'ListDrive (DrivePath)
    ListDevice (DevicePath)
    
    
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Cancel = 1
    If StdIO.Ready = True Then
        End
    Else
        bExitAfterCancel = True
            If StdIO.Ready = False Then
                'AddLog "[" & Time & "] Canceling program.."
                StdIO.Cancel
            Else
                'AddLog "[" & Time & "] Try executing a program first ;)"
            End If
    End If
    
    AddLog "Killing ADB Server.."
    Shell ("cmd /C " & ADBPath & " start-server")
    AddLog "Done"
    
End Sub

Private Sub lstDeviceDir_GotFocus()

    ActivePane = "Device"
    ActiveDevice = "Dir"
    
End Sub

Private Sub lstDeviceDir_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then lstDeviceDir_DblClick

End Sub


Private Sub lstDeviceFile_GotFocus()

    ActivePane = "Device"
    ActiveDevice = "File"

End Sub



'Private Sub lstDrive_DblClick()
'
'    Dim Selection As String
'    Selection = lstDrive.Text
'
'    If Selection = ".." Then
'
'        If DrivePath = "C:\" Then Exit Sub
'
'        Dim myarray() As String
'        myarray() = Split(DrivePath, "\")
'        DrivePath = ""
'
'        Dim i As Integer
'        For i = 0 To UBound(myarray) - 1
'            If (i < UBound(myarray)) Then
'                DrivePath = DrivePath & myarray(i) & "\"
'            End If
'        Next i
'
'        DrivePath = Chr(34) & DrivePath & Chr(34)
'
'    Else
'        If Not Right(DrivePath, 1) = "\" Then DrivePath = DrivePath & "\"
'        DrivePath = Chr(34) & DrivePath & Selection & Chr(34)
'
'    End If
'
'    ListDrive DrivePath
'
'End Sub

'Private Sub lstDrive_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 13 Then lstDrive_DblClick
'
'End Sub

Private Sub mnuAbout_Click(Index As Integer)

    ' Load about screen
    MsgBox "ADB File Explorer " & AppVersion & vbCrLf & "Created by Dean Barrow"
    
End Sub

Private Sub mnuApp_Click()

    Me.Hide
    frmAppManager.Show

End Sub

Private Sub mnuBackup_Click()

    MsgBox "coming soon"

End Sub

Private Sub mnuExit_Click(Index As Integer)
    
    ' Exit
    End
    
End Sub

Private Sub mnuInfo_Click()

    MsgBox "coming soon"

End Sub

Private Sub mnuRestore_Click()

    MsgBox "coming soon"

End Sub

Private Sub mnuSettings_Click()

    MsgBox "coming soon"

End Sub

Private Sub mnuWebsite_Click()

    MsgBox "http://code.google.com/g/adb-file-explorer"
    
End Sub

Public Sub RunCommand(Command As String)

    ReturnData = ""
    
    If StdIO.Ready = True Then
        StdIO.CommandLine = Command
        'AddLog "[" & Time & "] Executing command:"
        'AddLog "-> " & StdIO.CommandLine
        StdIO.ExecuteCommand Command
        'Or simply StdIO.ExecuteCommand txtCommand.Text
    Else
        'AddLog "[" & Time & "] Cannot execute command, already in use!"
    End If

End Sub

'Private Sub ListDrive(Path As String)
'
'Dim Command As String
'Command = "cmd /c dir /D /OGN /B " & Path
'
''Update textbox, strip quotes after command
'Path = Strip(Path, Chr(34))
'If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
'If Path = "C:" Then Path = "C:\"
'txtDrive.Text = Path
'
'RunCommand Command
'
'Dim myarray() As String
'myarray() = Split(ReturnData, vbCrLf)
'
'lstDrive.Clear
'lstDrive.AddItem ("..")
'
'Dim i As Integer
'For i = 0 To UBound(myarray)
    'If (i < UBound(myarray)) Then
        'lstDrive.AddItem (myarray(i))
    'End If
'Next i
'
'End Sub

Private Sub lstDeviceDir_DblClick()
    
    Dim Selection As String
    Selection = lstDeviceDir.Text
    
    If Left(Selection, 3) = " | " Then
    
        Dim SkipSection As Boolean
    
        Selection = Right(Selection, Len(Selection) - 3)
        
        Dim myarray() As String
        ReDim myarray(1000) As String
        myarray() = Split(DevicePath, "/")
        DevicePath = ""

        Dim i As Integer
        For i = 0 To UBound(myarray) - 1
            If (i < UBound(myarray)) Then
                DevicePath = DevicePath & myarray(i) & "/"
                If myarray(i) = Selection Then
                    i = UBound(myarray) - 1
                    SkipSection = True
                End If
            End If
        Next i
        
        DevicePath = Chr(34) & DevicePath & Chr(34)
    End If
    
    If Selection = ".." Then
        
        If DevicePath = "/" Then Exit Sub
        
        'Dim myarray() As String
        ReDim myarray(1000) As String
        myarray() = Split(DevicePath, "/")
        DevicePath = ""

        'Dim i As Integer
        For i = 0 To UBound(myarray) - 1
            If (i < UBound(myarray)) Then
                DevicePath = DevicePath & myarray(i) & "/"
            End If
        Next i
    
        DevicePath = Chr(34) & DevicePath & Chr(34)
        
    ElseIf Selection = "/" Then
        
        DevicePath = Selection
        
    Else
        If SkipSection = False Then
            If Not Right(DevicePath, 1) = "/" Then DevicePath = DevicePath & "/"
            DevicePath = Chr(34) & DevicePath & Selection & Chr(34)
        End If
        
    End If
    
    ListDevice DevicePath
    
End Sub

Private Sub ListDevice(path As String)

Dim Command As String
Command = ADBPath & " shell busybox ls -1 -A -L -p --color=never " & path

'Update textbox, strip quotes after command
path = Strip(path, Chr(34))
If Right(path, 1) = "/" Then path = Left(path, Len(path) - 1)
If path = "" Then path = "/"
txtDevice.Text = path

'if adb fails then retry
redo:

RunCommand Command

'Start ADB if not running
    If Left(ReturnData, 20) = "* daemon not running" Then
        AddLog ("ADB server not running")
        AddLog "Starting ADB server.."
        Shell ("cmd /C " & ADBPath & " start-server")
        AddLog "Done"
        GoTo redo
    End If

lstDeviceDir.Clear
lstDeviceFile.Clear
'lstDeviceDir.AddItem ("..")
'If Not Path = "/" Then lstDeviceDir.AddItem ("/")

        Dim myarray() As String
        ReDim myarray(1000) As String
        myarray() = Split(path, "/")

        Dim i As Integer
        For i = 0 To UBound(myarray)
                If Not myarray(i) = "" Then
                    lstDeviceDir.AddItem (" | " & myarray(i))
                Else
                    If lstDeviceDir.ListCount = 0 Then lstDeviceDir.AddItem ("/")
                End If
        Next i

'Catch errors
If ReturnData = "" Then Exit Sub

If Left(ReturnData, Len(ReturnData) - 2) = "error: device not found" Then
    MsgBox "Device not found. Please make sure your phone is connected."
    lstDeviceDir.Clear
    lstDeviceFile.Clear
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
        If Right(myarray(i), 1) = "/" Then lstDeviceDir.AddItem ("" & Left(myarray(i), Len(myarray(i)) - 1))
        If Not Right(myarray(i), 1) = "/" Then lstDeviceFile.AddItem (myarray(i))
        DoEvents
    End If
Next i


End Sub

Private Sub StdIO_CancelFail()
    'AddLog "[" & Time & "] Cancel failed to end program. No longer reading pipes."
    DoEvents
    If bExitAfterCancel Then End
End Sub

Private Sub StdIO_CancelSuccess()
    'AddLog "[" & Time & "] Cancel success! No longer reading pipes."
    DoEvents
    If bExitAfterCancel Then End
End Sub

Private Sub StdIO_Complete()
    'AddLog "[" & Time & "] Complete!"
End Sub

Private Sub StdIO_Error(ByVal Number As Integer, ByVal Description As String)
    'AddLog "[" & Time & "] Error #" & Number & ": " & Description
End Sub

Private Sub StdIO_GotData(ByVal Data As String)
    ReturnData = Data
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

