VERSION 5.00
Begin VB.Form FHookXP 
   BackColor       =   &H80000016&
   Caption         =   "Clipboard Data Interceptor (Logger) by Alex Dragokas"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "FHookXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------
'Registry Hex Convertor
'
'Compiled by Dragokas
'   http://safezone.cc/members/dragokas.6966/
'-----------------------------------------------
'
'Parts of Code developed by:
'
'Karl E. Peterson - HookXP (Subclassing)
'   http://vb.mvps.org/samples/HookXP/
'
'Ross Donald - Clipboard Monitor (VB.NET code)
'   http://www.radsoftware.com.au/articles/clipboardmonitor.aspx
'
'Ronald Kas - Registry Handler
'   http://www.vbfrance.com/telecharger.aspx?ID=48372
'
'-----------------------------------------------

Option Explicit

' Subclassing interface
Implements IHookXP

' Message handlers
'Private m_ClipHook As CHookClipBoard
'Private WithEvents m_ClipboardEvent As CHookClipBoard

Const MAX_PATH                          As Long = 260&

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(255) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Declare Function GetClipboardOwner Lib "user32.dll" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, lpFilePart As Long) As Long
Private Declare Function QueryFullProcessImageName Lib "kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As Long, ByVal lpdwSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceW" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long
Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const EM_LIMITTEXT      As Long = &HC5&
Private Const ERROR_ACCESS_DENIED    As Long = 5&
Private Const ERROR_PARTIAL_COPY            As Long = 299&

Private WithEvents m_ClipHook As CHookClipBoard
Attribute m_ClipHook.VB_VarHelpID = -1
Private m_Clip As modClip
Private Reg As RegistryHandler

Private ffLog As Integer
Dim bIsWinVistaOrLater As Boolean

Private Sub Form_Resize()
    Text1.Width = Me.Width - 230
    Text1.Height = Me.Height - 550
End Sub

' *********************************************
'  Custom Events
' *********************************************
Private Sub m_ClipHook_Change(sMsg As String, Parameter As Long)
    TraceClipboardOwner sMsg, Parameter
End Sub

'Private Sub m_ClipHook_Change2()
'    TraceClipboardOwner "WM_CLIPBOARDUPDATE"
'End Sub
'
'Private Sub m_ClipHook_Change3()
'    TraceClipboardOwner "WM_DESTROYCLIPBOARD"
'End Sub

Sub TraceClipboardOwner(sMsg As String, Parameter As Long)
    Dim sKey$, arr, key
    Dim hClipOwner As Long
    Dim ProcessID As Long
    Dim hThreadId As Long
    Dim FilePath As String
    
    If sMsg = "WM_CHANGECBCHAIN" Then
        hClipOwner = Parameter
    Else
        hClipOwner = GetClipboardOwner()
    End If
    
    If hClipOwner <> 0 Then
        hThreadId = GetWindowThreadProcessId(ByVal hClipOwner, ProcessID)
        
        If ProcessID <> 0 Then
            FilePath = GetFilePathByPID(ProcessID)
        End If
        
        LogIt "PID = " & ProcessID & ". Path = " & FilePath
        
    End If
    
    ' Get Text Clipboard Data Class
    If m_Clip Is Nothing Then Set m_Clip = New modClip
    ' Get ClipBoard text Data
    sKey = m_Clip.ClipGet
    
    LogIt "Clipboard data: " & sKey & vbCrLf & _
        "Message: " & sMsg & vbCrLf & _
        "Time: " & Now() & vbCrLf & _
        "-------------------------------"
    
    ' If this is a key
    If Len(sKey) <> 0 And Left(sKey, 5) = "HKEY_" Then
        'Get Routines
'        arr = Reg.EnumKeys(sKey)
'        If IsArray(arr) Then
'            For Each key In arr
'                Debug.Print key
'            Next
'        End If
    End If
End Sub

Sub LogIt(sText As String)
    Me.Text1.Text = Me.Text1.Text & vbCrLf & sText
    Debug.Print sText
    Print #ffLog, sText
End Sub

' *********************************************
'  Native Events
' *********************************************

Private Sub Form_Load()
    ffLog = FreeFile()
    Open App.Path & "\Clipboard_Log.txt" For Output As #ffLog

    Dim osi As OSVERSIONINFOEX
    
    osi.dwOSVersionInfoSize = Len(osi)
    GetVersionEx osi
    bIsWinVistaOrLater = (osi.dwMajorVersion >= 6)

    SendMessage Me.Text1.hWnd, EM_LIMITTEXT, 0&, ByVal 0&

   ' Generic hook for odds and ends?
   Call HookSet(Me.hWnd, Me)
   ' Delegate monitoring of Clipboard to class.
   Set m_ClipHook = New CHookClipBoard
   m_ClipHook.hWnd = Me.hWnd
   ' Get Registry Handler Class
   Set Reg = New RegistryHandler
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Clear hook processing in this instance.
   ' No need to explictly destroy others.
   Call HookClear(Me.hWnd, Me)
   Close ffLog
End Sub

' *********************************************
'  Implemented Subclassing Interface
' *********************************************
Private Function IHookXP_Message(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
      'Debug.Print "hWnd: 0x"; Hex$(hWnd), "Msg: 0x"; Hex$(uiMsg), _
                  "wParam: 0x"; Hex$(wParam), "lParam: 0x"; Hex$(lParam), _
                  "RefData: "; dwRefData
   IHookXP_Message = HookDefault(hWnd, uiMsg, wParam, lParam)
End Function

' *********************************************
'  Private Methods
' *********************************************


Function GetFilePathByPID(PID As Long) As String
    On Error GoTo ErrorHandler:

    Const MAX_PATH_W                        As Long = 32767&
    Const PROCESS_VM_READ                   As Long = 16&
    Const PROCESS_QUERY_INFORMATION         As Long = 1024&
    Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
    
    Dim ProcPath    As String
    Dim hProc       As Long
    Dim cnt         As Long
    Dim pos         As Long
    Dim FullPath    As String
    Dim SizeOfPath  As Long
    Dim lpFilePart  As Long

    hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0&, PID)
    
    If hProc = 0 Then
        If Err.LastDllError = ERROR_ACCESS_DENIED Then
            hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0&, PID)
        End If
    End If
    
    If hProc <> 0 Then
    
        If bIsWinVistaOrLater Then
            cnt = MAX_PATH_W + 1
            ProcPath = Space$(cnt)
            Call QueryFullProcessImageName(hProc, 0&, StrPtr(ProcPath), VarPtr(cnt))
        End If
        
        If 0 <> Err.LastDllError Or Not bIsWinVistaOrLater Then     'Win 2008 Server (x64) can cause Error 128 if path contains space characters
        
            ProcPath = Space$(MAX_PATH)
            cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
        
            If cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
                ProcPath = Space$(MAX_PATH_W)
                cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
            End If
        End If
        
        If cnt <> 0 Then                          'clear path
            ProcPath = Left$(ProcPath, cnt)
            If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = Environ("SystemRoot") & Mid$(ProcPath, 12)
            If "\??\" = Left$(ProcPath, 4) Then ProcPath = Mid$(ProcPath, 5)
        End If
        
        If ERROR_PARTIAL_COPY = Err.LastDllError Or cnt = 0 Then     'because GetModuleFileNameEx cannot access to that information for 64-bit processes on WOW64
            ProcPath = Space$(MAX_PATH)
            cnt = GetProcessImageFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
            
            If cnt <> 0 Then
                ProcPath = Left$(ProcPath, cnt)
                
                ' Convert DosDevice format to Disk drive format
                If StrComp(Left$(ProcPath, 8), "\Device\", 1) = 0 Then
                    pos = InStr(9, ProcPath, "\")
                    If pos <> 0 Then
                        FullPath = ConvertDosDeviceToDriveName(Left$(ProcPath, pos - 1))
                        If Len(FullPath) <> 0 Then
                            ProcPath = FullPath & Mid$(ProcPath, pos + 1)
                        End If
                    End If
                End If
                
            End If
            
        End If
        
        If cnt <> 0 Then    'if process ran with 8.3 style, GetModuleFileNameEx will return 8.3 style on x64 and full pathname on x86
                            'so wee need to expand it ourself
        
            FullPath = Space$(MAX_PATH)
            SizeOfPath = GetFullPathName(StrPtr(ProcPath), MAX_PATH, StrPtr(FullPath), lpFilePart)
            If SizeOfPath <> 0& Then
                GetFilePathByPID = Left$(FullPath, SizeOfPath)
            Else
                GetFilePathByPID = ProcPath
            End If
            
        End If
        
        CloseHandle hProc
    End If
    
    Exit Function
ErrorHandler:
End Function


Public Function ConvertDosDeviceToDriveName(inDosDeviceName As String) As String
    On Error Resume Next

    Static DosDevices   As New Collection
    
    If DosDevices.Count Then
        ConvertDosDeviceToDriveName = DosDevices(inDosDeviceName)
        Exit Function
    End If
    
    Dim aDrive()        As String
    Dim sDrives         As String
    Dim cnt             As Long
    Dim i               As Long
    Dim DosDeviceName   As String
    
    cnt = GetLogicalDriveStrings(0&, StrPtr(sDrives))
    
    sDrives = Space(cnt)
    
    cnt = GetLogicalDriveStrings(Len(sDrives), StrPtr(sDrives))

    If 0 = Err.LastDllError Then
    
        aDrive = Split(Left$(sDrives, cnt - 1), vbNullChar)
    
        For i = 0 To UBound(aDrive)
            
            DosDeviceName = Space(MAX_PATH)
            
            cnt = QueryDosDevice(StrPtr(Left$(aDrive(i), 2)), StrPtr(DosDeviceName), Len(DosDeviceName))
            
            If cnt <> 0 Then
            
                DosDeviceName = Left$(DosDeviceName, InStr(DosDeviceName, vbNullChar) - 1)

                DosDevices.Add aDrive(i), DosDeviceName

            End If
            
        Next
    
    End If
    
    ConvertDosDeviceToDriveName = DosDevices(inDosDeviceName)
    
End Function

Function IsWOW64() As Boolean
    Dim hModule As Long, procAddr As Long, lIsWin64 As Long
    
    hModule = LoadLibrary(StrPtr("kernel32.dll"))
    If hModule Then
        procAddr = GetProcAddress(hModule, "IsWow64Process")
        If procAddr <> 0 Then
            IsWow64Process GetCurrentProcess(), lIsWin64
            IsWOW64 = CBool(lIsWin64)
        End If
        FreeLibrary hModule
    End If
End Function
