VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cKillProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The following are just global constant for the class itself
Const c_Version = "Kill Process Class v0.1 BETA by Amine Haddad"

'The following are just declarations (also known as API Calls)
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'The following are just constants and types declared for this class only
Private Const WM_CLOSE As Long = &H10
Private Const WM_DESTROY As Long = &H2
Private Const WM_ENDSESSION = &H16
Private Const PROCESS_TERMINATE As Long = &H1
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type

'The following are types to determine the settings of this class
Private Type t_Settings
    Init As Boolean
    Is9X As Boolean
End Type

Dim Settings As t_Settings

'The following is to allow the user to get the version of this class
Public Property Get Version() As String
    'The version of this class.
    Version = c_Version
End Property


'The following are the functions used in this class
Public Function KillProcess(ByVal lProcessID As Long) As Boolean
    Dim lHandle As Long
    
    If Not Settings.Init Then Call InitializeClass
    If ClosePID(lProcessID) = True Then
        If Settings.Is9X Then
            lHandle = OpenProcess(PROCESS_TERMINATE, False, lProcessID)
            If lHandle = 0 Then
                KillProcess = False
            Else
                KillProcess = CBool(TerminateProcess(lHandle, 0&))
                CloseHandle lHandle
            End If
        Else
            KillProcess = True
        End If
    Else
        KillProcess = False
    End If
End Function

Private Sub InitializeClass()
    'This function needs to be ran before running any other functions.
    'We NEED to know if we are in Windows 9x or not.
    Dim OsInfo As OSVERSIONINFO

    With OsInfo
        .OSVSize = Len(OsInfo)
        .szCSDVersion = Space(128)
        
        Call GetVersionEx(OsInfo)
        'After this line, we will know if the system is 9X or else.
        Settings.Is9X = (.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (.dwVerMajor > 4) Or (.dwVerMajor = 4 And .dwVerMinor > 0) Or (.PlatformID = VER_PLATFORM_WIN32_WINDOWS And .dwVerMajor = 4 And .dwVerMinor = 0)
    End With

    'We have successfully initialized this class.
    Settings.Init = True
End Sub

Private Function ClosePID(ByVal lProcessID As Long) As Boolean
    'This function here will go through all windows and kill the pid that it was given
    Dim hWndChild As Long
    Dim lThreadProcessID As Long
    
    hWndChild = GetWindow(GetDesktopWindow(), GW_CHILD)

    Do While (hWndChild <> 0)
        If (GetParent(hWndChild) = 0) Then
            Call GetWindowThreadProcessId(hWndChild, lThreadProcessID)
            If (lProcessID = lThreadProcessID) Then
                Call PostMessage(hWndChild, IIf(Settings.Is9X, WM_ENDSESSION, WM_CLOSE), IIf(Settings.Is9X, True, False), 0&)
                ClosePID = True
            End If
        End If
        
        hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
    Loop
End Function

Public Function PIDInUse(ByVal lProcessID As Long) As Boolean
    'This function will return true if the PID is in use.
    Dim hWndChild As Long
    Dim lThreadProcessID As Long
    
    hWndChild = GetWindow(GetDesktopWindow(), GW_CHILD)

    Do While (hWndChild <> 0)
        If (GetParent(hWndChild) = 0) Then
            Call GetWindowThreadProcessId(hWndChild, lThreadProcessID)
            If (lProcessID = lThreadProcessID) Then
                PIDInUse = True
                Exit Function
            End If
        End If
        
        hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
        DoEvents
    Loop
End Function
