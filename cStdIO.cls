VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStdIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The following are just global constant for the class itself
Const c_Version = "Standard Input/Output Class v0.1 BETA by Amine Haddad"

'The following are declarations of API calls
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function SetNamedPipeHandleState Lib "kernel32" (ByVal hNamedPipe As Long, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'The following are types used by API calls listed above
Private Type SECURITY_ATTRIBUTES
    nLength                 As Long
    lpSecurityDescriptor    As Long
    bInheritHandle          As Long
End Type

Private Type STARTUPINFO
    cb                      As Long
    lpReserved              As Long
    lpDesktop               As Long
    lpTitle                 As Long
    dwX                     As Long
    dwY                     As Long
    dwXSize                 As Long
    dwYSize                 As Long
    dwXCountChars           As Long
    dwYCountChars           As Long
    dwFillAttribute         As Long
    dwFlags                 As Long
    wShowWindow             As Integer
    cbReserved2             As Integer
    lpReserved2             As Long
    hStdInput               As Long
    hStdOutput              As Long
    hStdError               As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess                As Long
    hThread                 As Long
    dwProcessId             As Long
    dwThreadID              As Long
End Type

'The following are constants needed (but not all used) by API calls above
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const PIPE_WAIT = &H0
Private Const PIPE_NOWAIT = &H1
Private Const PIPE_READMODE_BYTE = &H0
Private Const PIPE_READMODE_MESSAGE = &H2
Private Const PIPE_TYPE_BYTE = &H0
Private Const PIPE_TYPE_MESSAGE = &H4

'The following are variables required throughout the program
Private mCommand    As String   'The command to process
Private mOutput     As String   'The final output of the whole program
Private mCancel     As Boolean  'Set this to True to cancel
Private mReady      As Boolean  'Are we ready to launch new command?

'The following are variables required throughbout the program's functions
Dim lRetVal         As Long                 'RETurn VALue of a certain function.
Dim hReadPipe       As Long                 'Read Pipe handle created by CreatePipe
Dim hWritePipe      As Long                 'Write Pite handle created by CreatePipe
Dim lBytesRead      As Long                 'Amount of byte read from the Read Pipe handle
Dim sBuffer         As String * 4096        'String buffer reading the Pipe
Dim hReadPipe2      As Long                 'Read Pipe handle created by CreatePipe
Dim hWritePipe2     As Long                 'Write Pite handle created by CreatePipe

'The following are events to be launched throughout the program
Public Event GotData(ByVal Data As String)
Public Event CancelSuccess()
Public Event CancelFail()
Public Event Complete()
Public Event Error(ByVal Number As Integer, ByVal Description As String)

'Definitions of error messages throughout the program (passed in the Error event):
'Error #400: Not ready to process another command.
'Error #401: Command Line empty.
'Error #402: Not processing a command to cancel.
'Error #403: Not ready to change settings.
'Error #404: CreatePipe failed.
'Error #405: SetNamedPipeHandleState failed.
'Error #406: CreateProcess failed.

'The following are properties that can be used to keep track of what we are doing
Public Property Let CommandLine(ByVal DoCommand As String)
    'This allows us to set the new command line to process.
    If mReady = True Then
        mCommand = DoCommand
    Else
        RaiseEvent Error(402, "Not ready to change settings.")
    End If
End Property

Public Property Get CommandLine() As String
    'This allows us to read the current command line setting.
    CommandLine = mCommand
End Property

Public Property Get Ready() As Boolean
    'This allows us to read the state of the program.
    'Will return True if it is ready to process another command.
    Ready = mReady
End Property

Public Property Get Version() As String
    'The version of this class.
    Version = c_Version
End Property

'The following are events initialized by the class
Private Sub Class_Initialize()
    'Once class started, we can't possibly already have a command running,
    'so we will set the ready variable to true so we can process another.
    mReady = True
End Sub

'The following are subs and functions used in the program.
Public Sub Cancel()
    'If called, and under condition a program is being processed, it will
    'interrupt and end the program.
    If mReady = False Then
        mCancel = True
    Else
        RaiseEvent Error(402, "Not processing a command to cancel.")
    End If
End Sub

Public Function ExecuteCommand(Optional CommandLine As String) As String
    'This is it. The function that will actually do the work. It is not hard,
    'read through the comments to understand.
    Dim tStartup        As STARTUPINFO          'Self explanatory..
    Dim tProc           As PROCESS_INFORMATION  'Self explanatory..
    Dim tSecAttr        As SECURITY_ATTRIBUTES  'Self explanatory..
    
    'Let's check if we are ready to process this command.
    If mReady = False Then
        'We are not. Warn the user and exit the function.
        RaiseEvent Error(400, "Not ready to process another command.")
        Exit Function
    End If
    
    'We are ready, let's tell it that we are not ready so we don't get another command
    'while processing the current one. Also set mCancel to false, we don't want to
    'cancel something before we start it do we ;)
    mReady = False
    mCancel = False
    
    'If the parameter we got is not empty, then let's overwrite the current mCommand value.
    If Len(CommandLine) > 0 Then
        mCommand = CommandLine
    End If
    
    'If we still have a empty command line (mCommand) then let's just tell the user and
    'exit the function.
    If Len(mCommand) = 0 Then
        mReady = True 'We put mReady before RaiseEvent because user might launch another command on the
        'error event. If we put it after, it would tell him not ready, but now it will tell him it is ready.
        RaiseEvent Error(401, "Command Line empty.")
        Exit Function
    End If
    
    'Let's set the Security Attributes that we will pass on
    tSecAttr.nLength = LenB(tSecAttr)
    tSecAttr.bInheritHandle = True
    tSecAttr.lpSecurityDescriptor = False
    
    'Now, we will create the output pipe. lRetVal will return 0 if it failed.
    lRetVal = CreatePipe(hReadPipe, hWritePipe, tSecAttr, 0&)
    
    'Let's check if it succeeded or failed.
    If lRetVal = 0 Then
        'If an error occur during the Pipe creation exit
        mReady = True
        RaiseEvent Error(404, "CreatePipe failed.")
        Exit Function
    End If
    
    'Do the input pipe
    lRetVal = CreatePipe(hReadPipe2, hWritePipe2, tSecAttr, 0&)
    If lRetVal = 0 Then
        'If an error occur during the Pipe creation exit
        mReady = True
        RaiseEvent Error(404, "CreatePipe failed.")
        Exit Function
    End If
    
    'The next step is to set it to non-blocking mode meaning that it will instantly
    'return when ReadFile is called (you will understand later).
    lRetVal = SetNamedPipeHandleState(hReadPipe, PIPE_READMODE_BYTE Or PIPE_NOWAIT, 0&, 0&)
    
    If lRetVal <> 0 Then
        'Well, we failed. Let's exit.
        '(NOTICE: You don't have to exit, but since this is to show how to make it
        '         non-blocking only then I will set it to exit when it fails.)
        mReady = True
        RaiseEvent Error(405, "SetNamedPipeHandleState failed.")
        Exit Function
    End If

    'Let's set the StartupInfo for the command line when it is launched
    tStartup.cb = LenB(tStartup)
    tStartup.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    tStartup.wShowWindow = SW_HIDE      'We want window to not show up so we use SW_HIDE.
    tStartup.hStdOutput = hWritePipe    'Set the StdOut and StdError output
    tStartup.hStdError = hWritePipe     'to the same Write Pipe handle.
    tStartup.hStdInput = hReadPipe2
    
    'Let's launch the program.
    lRetVal = CreateProcessA(0&, mCommand, tSecAttr, tSecAttr, _
        1&, NORMAL_PRIORITY_CLASS, 0&, 0&, tStartup, tProc)
        
    'Let's check if it succeeded.
    If lRetVal <> 1 Then
        'Unfortunatly, we failed. Maybe it can't find CMD.EXE?
        mReady = True
        RaiseEvent Error(406, "CreateProcess failed.")
        Exit Function
    End If
    
    'Now we will clear the mOutput variable
    mOutput = ""
    
    'Okay, from this point on we might need assistance from cKillProcess
    'So let's bring in KP ;)
    Dim KP As New cKillProcess
    
    'Now that all is set, let's start getting the output from the ReadPipe handle
    Do
        DoEvents                                'Let's not hog cpu
        If mCancel = True Then Exit Do          'If we need to cancel, exit do.
        Sleep 100                                'Just for smooth sailing.
        lRetVal = ReadFile(hReadPipe, sBuffer, 4096, lBytesRead, 0&)
        If lRetVal <> 0 Then
            'We got data!
            'Let's add it to mOutput (all data since begining)
            mOutput = mOutput & Left(sBuffer, lBytesRead)
            'And finally we will send data to the GotData event.
            RaiseEvent GotData(Left(sBuffer, lBytesRead))
            'Let's just not hog cpu again :P
            DoEvents
        End If
        'And loop until we don't see the process anymore :)
    Loop While KP.PIDInUse(tProc.dwProcessId)
    
    'Now we're done so close the opened handles
    Call CloseHandle(tProc.hProcess)
    Call CloseHandle(tProc.hThread)
    Call CloseHandle(hReadPipe)
    Call CloseHandle(hReadPipe2)
    Call CloseHandle(hWritePipe)
    Call CloseHandle(hWritePipe2)
    
    'Return the Outputs property with the entire DOS output
    ExecuteCommand = mOutput
    
    'Set it so we are ready to launch another command
    mReady = True
    
    'And finally, check if we ended with a cancel.
    'If we did, then end the process and call the event respectivly.
    'If we didn't, then call Complete.
    If mCancel Then
        If KP.KillProcess(tProc.dwProcessId) Then
            Set KP = Nothing
            RaiseEvent CancelSuccess
        Else
            Set KP = Nothing
            RaiseEvent CancelFail
        End If
    Else
        Set KP = Nothing
        RaiseEvent Complete
    End If
    
    'And we're done ;)
End Function

Public Function WriteData(ByVal strData As String) As Long
    'This function will return -1 if it failed to write to pipe,
    'otherwise, it will return the bytes written.
    Dim lBytesWritten As Long
    Dim arrByte() As Byte
    arrByte() = StrConv(strData & vbCrLf & Chr(0), vbFromUnicode)
    lRetVal = WriteFile(hWritePipe2, arrByte(0), UBound(arrByte), lBytesWritten, 0&)
    WriteData = IIf(lRetVal = 0, -1, lBytesWritten - 2)
    'Ok, in the line that just passed I did lBytesWritten - 3 because we added a vbCrLf (2 bytes) and a Chr(0) (1 byte)
    'and let's not forget its base 0 so 3 is really 0-1-2 so 2.
    'I didn't want those included because if the user sends 'hello' and it said
    'sent 8 bytes then he would just be wondering what happened. This will fix that problem.
End Function
