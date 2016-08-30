Attribute VB_Name = "Module1"
Option Explicit

Public Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Public Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesRead As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_COMMAND = &H111
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function CreateIExprSrvObj Lib "msvbvm60" (ByVal u1_0 As Long, ByVal u2_4 As Long, ByVal u3_0 As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Any, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function CreateJobObject Lib "kernel32.dll" Alias "CreateJobObjectA" (ByRef lpJobAttributes As Any, ByVal lpName As String) As Long
Public Declare Function AssignProcessToJobObject Lib "kernel32.dll" (ByVal hJob As Long, ByVal hProcess As Long) As Long
Public Declare Function SetInformationJobObject Lib "kernel32.dll" (ByVal hJob As Long, ByVal JobObjectInformationClass As JOBOBJECTINFOCLASS, lpJobObjectInformation As Any, ByVal cbJobObjectInformationLength As Long) As Long
Public Enum JOBOBJECTINFOCLASS
    JobObjectBasicAccountingInformation = 1
    JobObjectBasicLimitInformation
    JobObjectBasicProcessIdList
    JobObjectBasicUIRestrictions
    JobObjectSecurityLimitInformation
    JobObjectEndOfJobTimeInformation
    JobObjectAssociateCompletionPortInformation
    JobObjectBasicAndIoAccountingInformation
    JobObjectExtendedLimitInformation
    JobObjectJobSetInformation
    MaxJobObjectInfoClass
End Enum
Public Const JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = &H2000 'extended limit
Public Type JOBOBJECT_BASIC_LIMIT_INFORMATION
    PerProcessUserTimeLimit As LARGE_INTEGER
    PerJobUserTimeLimit As LARGE_INTEGER
    LimitFlags As Long
    MinimumWorkingSetSize As Long
    MaximumWorkingSetSize As Long
    ActiveProcessLimit As Long
    Affinity As LARGE_INTEGER
    PriorityClass As Long
    SchedulingClass As Long
End Type
Public Type IO_COUNTERS
    ReadOperationCount As LARGE_INTEGER
    WriteOperationCount As LARGE_INTEGER
    OtherOperationCount As LARGE_INTEGER
    ReadTransferCount As LARGE_INTEGER
    WriteTransferCount As LARGE_INTEGER
    OtherTransferCount As LARGE_INTEGER
End Type
Public Type JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    BasicLimitInformation As JOBOBJECT_BASIC_LIMIT_INFORMATION
    IoInfo As IO_COUNTERS
    ProcessMemoryLimit As Long
    JobMemoryLimit As Long
    PeakProcessMemoryUsed As Long
    PeakJobMemoryUsed As Long
End Type
Public Declare Function AllocConsole Lib "kernel32" () As Long
Public Declare Function FreeConsole Lib "kernel32" () As Long
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Public Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Public Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub DebugBreak Lib "kernel32" ()
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYCAPTION = 4
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CYMENU = 15
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const MK_RBUTTON = &H2
Public Const MK_LBUTTON = &H1
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

