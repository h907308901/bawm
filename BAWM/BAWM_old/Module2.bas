Attribute VB_Name = "Module2"
Option Explicit

Dim hStdIn As Long
Dim hStdOut As Long
Dim hJob As Long
Dim hWnd1 As Long, hProcess1 As Long, pid1 As Long
Dim hWnd2 As Long, hProcess2 As Long, pid2 As Long
Dim hWnd3 As Long, hProcess3 As Long, pid3 As Long
Dim hWnd4 As Long, hProcess4 As Long, pid4 As Long
Dim Width As Long, Height As Long

Dim ExitFlag As Boolean

Public Const TOTAL_TIME = 218000 'milliseconds
Dim CurrentTime As Long 'milliseconds 'written by main thread
Dim TotalFrame As Long, CurrentFrame As Long 'written by main thread
Dim SleepValue As Long 'written by thread 2

Sub Main()
    Dim si As STARTUPINFO, pi As PROCESS_INFORMATION, ret As Long
    Dim eli As JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    Dim hwnd As Long, pid As Long, h As Long
    Dim hMenu As Long, hSubMenu As Long, id As Long
    Dim lpRect As RECT
    Dim Lines(1 To 61) As String
    Dim i As Long, j As Long, s As String
    Dim b1(1 To 512) As Byte, b2(1 To 512) As Byte, b3(1 To 512) As Byte, b4(1 To 512) As Byte
    Dim BeginTime As Long
    Dim XBase As Long, YBase As Long, Pos As Long
    Dim byt() As Byte, TempPath As String, str2 As String * 1024
    On Error GoTo Err
    AllocConsole
    hStdIn = GetStdHandle(STD_INPUT_HANDLE)
    hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    PrintLine "Extracting files..."
    ret = GetTempPath(1024, str2)
    TempPath = Left$(str2, ret)
    byt = LoadResData("DATA", "DATA")
    Open TempPath & "res.cab" For Binary As #1
    Put #1, , byt
    Close #1
    ret = CreateProcess(vbNullString, "expand " & TempPath & "res.cab -F:* " & TempPath, ByVal 0, ByVal 0, 0, 0, ByVal 0, vbNullString, si, pi)
    If ret = 0 Then
        PrintLine "Failed to extract"
        GoTo Final
    End If
    WaitForSingleObject pi.hProcess, -1
    CloseHandle pi.hProcess
    CloseHandle pi.hThread
    PrintLine "Creating job object..."
    hJob = CreateJobObject(ByVal 0, vbNullString)
    If hJob = 0 Then
        PrintLine "Failed to create job object, last err " & Err.LastDllError
        GoTo Final
    End If
    eli.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    ret = SetInformationJobObject(hJob, JobObjectExtendedLimitInformation, eli, Len(eli))
    If ret = 0 Then
        PrintLine "Failed to set job object, last err " & Err.LastDllError
        GoTo Final
    End If
    For i = 1 To 4
        PrintLine "Running Winmine " & CStr(i)
        Shell "winmine", vbMaximizedFocus
        Sleep 100
        Do
            hwnd = FindWindow(vbNullString, "扫雷")
            If ExitFlag Then Exit For
            If hwnd <> 0 Then
                If hwnd <> hWnd1 And hwnd <> hWnd2 And hwnd <> hWnd3 And hwnd <> hWnd4 Then
                    GetWindowThreadProcessId hwnd, pid
                    h = OpenProcess(&H1F0FFF, 0, pid)
                    ret = CreateThread(ByVal 0, 0, AddressOf ThreadFunc1, h, 0, 0)
                    If ret = 0 Then
                        PrintLine "Failed to create thread, last err " & Err.LastDllError
                        GoTo Final
                    End If
                    Select Case i
                    Case 1
                        hWnd1 = hwnd
                        hProcess1 = h
                    Case 2
                        hWnd2 = hwnd
                        hProcess2 = h
                    Case 3
                        hWnd3 = hwnd
                        hProcess3 = h
                    Case 4
                        hWnd4 = hwnd
                        hProcess4 = h
                    End Select
                    Exit Do
                End If
            End If
        Loop
    Next
    ret = AssignProcessToJobObject(hJob, hProcess1)
    If ret = 0 Then
        PrintLine "Failed to assign process to job, last err " & Err.LastDllError
        'GoTo Final
    End If
    ret = AssignProcessToJobObject(hJob, hProcess2)
    If ret = 0 Then
        PrintLine "Failed to assign process to job, last err " & Err.LastDllError
        'GoTo Final
    End If
    ret = AssignProcessToJobObject(hJob, hProcess3)
    If ret = 0 Then
        PrintLine "Failed to assign process to job, last err " & Err.LastDllError
        'GoTo Final
    End If
    ret = AssignProcessToJobObject(hJob, hProcess4)
    If ret = 0 Then
        PrintLine "Failed to assign process to job, last err " & Err.LastDllError
        'GoTo Final
    End If
    CloseHandle ret
    hMenu = GetMenu(hWnd1)
    hSubMenu = GetSubMenu(hMenu, 0)
    id = GetMenuItemID(hSubMenu, 4)
    PostMessage hWnd1, WM_COMMAND, id, 0
    Sleep 100
    hMenu = GetMenu(hWnd2)
    hSubMenu = GetSubMenu(hMenu, 0)
    id = GetMenuItemID(hSubMenu, 4)
    PostMessage hWnd2, WM_COMMAND, id, 0
    Sleep 100
    hMenu = GetMenu(hWnd3)
    hSubMenu = GetSubMenu(hMenu, 0)
    id = GetMenuItemID(hSubMenu, 4)
    PostMessage hWnd3, WM_COMMAND, id, 0
    Sleep 100
    hMenu = GetMenu(hWnd4)
    hSubMenu = GetSubMenu(hMenu, 0)
    id = GetMenuItemID(hSubMenu, 4)
    PostMessage hWnd4, WM_COMMAND, id, 0
    Sleep 100
    '游戏
    '0 开局
    '1 ----
    '2 初级
    '3 中级
    '4 高级
    '5 自定义
    GetWindowRect hWnd1, lpRect
    Width = lpRect.Right - lpRect.Left
    Height = lpRect.Bottom - lpRect.Top
    SetWindowText hWnd1, "Bad Apple on WinMine"
    SetWindowText hWnd2, "By h907308901"
    SetWindowText hWnd3, ""
    SetWindowText hWnd4, ""
    SetWindowPos hWnd1, -1, 0, 0, 0, 0, 1
    SetWindowPos hWnd2, -1, Width, 0, 0, 0, 1
    SetWindowPos hWnd3, -1, 0, Height, 0, 0, 1
    SetWindowPos hWnd4, -1, Width, Height, 0, 0, 1
    Open TempPath & "output.txt" For Input As #1
    i = 0
    Do Until EOF(1)
        Line Input #1, s
        If s = "" Then i = i + 1
    Loop
    Close #1
    TotalFrame = i
    PrintLine "Total frames: " & TotalFrame
    PrintLine "Running..."
    SleepValue = 50
    Sleep 500
    ret = CreateThread(ByVal 0, 0, AddressOf ThreadFunc2, 0, 0, 0)
    If ret = 0 Then
        PrintLine "Failed to create thread, last err " & Err.LastDllError
        GoTo Final
    End If
    CloseHandle ret
    Open TempPath & "output.txt" For Input As #1

    XBase = GetSystemMetrics(SM_CXFRAME)
    YBase = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYFRAME) + GetSystemMetrics(SM_CYMENU)
    Pos = (YBase + 16) * 65536 + (XBase + 16)
    PostMessage hWnd1, WM_LBUTTONDOWN, MK_LBUTTON, Pos
    PostMessage hWnd1, WM_LBUTTONUP, MK_LBUTTON, Pos
    PostMessage hWnd2, WM_LBUTTONDOWN, MK_LBUTTON, Pos
    PostMessage hWnd2, WM_LBUTTONUP, MK_LBUTTON, Pos
    PostMessage hWnd3, WM_LBUTTONDOWN, MK_LBUTTON, Pos
    PostMessage hWnd3, WM_LBUTTONUP, MK_LBUTTON, Pos
    PostMessage hWnd4, WM_LBUTTONDOWN, MK_LBUTTON, Pos
    PostMessage hWnd4, WM_LBUTTONUP, MK_LBUTTON, Pos
    Sleep 500
    For i = 1 To 510
        b1(i) = 0
    Next
    WriteProcessMemory hProcess1, &H1005361, b1(1), 512, 0&
    WriteProcessMemory hProcess2, &H1005361, b1(1), 512, 0&
    WriteProcessMemory hProcess3, &H1005361, b1(1), 512, 0&
    WriteProcessMemory hProcess4, &H1005361, b1(1), 512, 0&
    InvalidateRect hWnd1, ByVal 0, 0
    InvalidateRect hWnd2, ByVal 0, 0
    InvalidateRect hWnd3, ByVal 0, 0
    InvalidateRect hWnd4, ByVal 0, 0
    mciExecute "play " & TempPath & "BadApple.mp3"
    BeginTime = GetTickCount
    Do Until EOF(1) Or ExitFlag
        For i = 1 To 61
            Line Input #1, Lines(i)
        Next
        '1
        For i = 1 To 16
            For j = 1 To 30
                b1((i - 1) * 32 + j) = Mid$(Lines(Int(i * 1.5)), 30 + Int(100 / 62 * j), 1)
            Next
        Next
        '2
        For i = 1 To 16
            For j = 1 To 30
                b2((i - 1) * 32 + j) = Mid$(Lines(Int(i * 1.5)), 30 + Int(100 / 62 * (j + 32)), 1)
            Next
        Next
        '3
        For i = 1 To 16
            For j = 1 To 30
                b3((i - 1) * 32 + j) = Mid$(Lines(Int(i * 1.5) + 36), 30 + Int(100 / 62 * j), 1)
            Next
        Next
        '4
        For i = 1 To 16
            For j = 1 To 30
                b4((i - 1) * 32 + j) = Mid$(Lines(Int(i * 1.5) + 36), 30 + Int(100 / 62 * (j + 32)), 1)
            Next
        Next
        WriteProcessMemory hProcess1, &H1005361, b1(1), 512, 0&
        WriteProcessMemory hProcess2, &H1005361, b2(1), 512, 0&
        WriteProcessMemory hProcess3, &H1005361, b3(1), 512, 0&
        WriteProcessMemory hProcess4, &H1005361, b4(1), 512, 0&
        InvalidateRect hWnd1, ByVal 0, 0
        InvalidateRect hWnd2, ByVal 0, 0
        InvalidateRect hWnd3, ByVal 0, 0
        InvalidateRect hWnd4, ByVal 0, 0
        CurrentFrame = CurrentFrame + 1
        CurrentTime = GetTickCount - BeginTime
        Sleep SleepValue
    Loop
    GoTo Final
Err:
    PrintLine "Error #" & Err.Number & ": " & Err.Description
Final:
    PrintLine "Exiting..."
    Close
    TerminateProcess hProcess1, 0
    TerminateProcess hProcess2, 0
    TerminateProcess hProcess3, 0
    TerminateProcess hProcess4, 0
    CloseHandle hProcess1
    CloseHandle hProcess2
    CloseHandle hProcess3
    CloseHandle hProcess4
    CloseHandle hJob
    FreeConsole
End Sub

Function ThreadFunc1(Parameter As Long) As Long
    CreateIExprSrvObj 0, 4, 0
    WaitForSingleObject Parameter, -1
    ExitFlag = True
End Function

Function ThreadFunc2(Parameter As Long) As Long
    Dim a As Single, b As Single
    CreateIExprSrvObj 0, 4, 0
    Do
        Sleep 100
        If TOTAL_TIME - CurrentTime = 0 Then
            Exit Do
        End If
        Do While CurrentTime = 0
            Sleep 1000
        Loop
        a = 1000 * (TotalFrame - CurrentFrame) / (TOTAL_TIME - CurrentTime)
        b = 1000 * CurrentFrame / CurrentTime
        If Abs(a - b) > 10 Then
            SleepValue = SleepValue - (a - b) / 10
        Else
            SleepValue = SleepValue - (a - b)
        End If
        If SleepValue < 0 Then SleepValue = 0
        SetWindowText hWnd3, "FPS:" & CStr(b)
        SetWindowText hWnd4, "Delay:" & CStr(SleepValue)
    Loop
End Function

Sub PrintLine(s As String)
    WriteConsole hStdOut, ByVal s & vbCrLf, LenB(StrConv(s, vbFromUnicode)) + 2, 0, ByVal 0
End Sub
