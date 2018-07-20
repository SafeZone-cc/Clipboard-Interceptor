VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHookClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2007-2009 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

'Code of SetClipboardViewer was translated from VB.NET code by Ross Donald
'http://www.radsoftware.com.au/articles/clipboardmonitor.aspx

' Win32 API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Subclassing interface
Implements IHookXP

' Events
Public Event Change(sMsg As String, Parameter As Long)   'WM_CHANGECBCHAIN
'Public Event Change2()  'WM_CLIPBOARDUPDATE
'Public Event Change3()  'WM_DESTROYCLIPBOARD

' Member variables
Private m_hWnd As Long
Private m_Enabled As Boolean
'Private m_SnapGap As Long
'Handle for next clipboard viewer...
Private mNextClipBoardViewerHWnd As Long

' Default values.
Private Const defEnabled As Boolean = True

'Constants for API Calls...
Private Const WM_NCDESTROY As Long = &H82&
Private Const WM_DRAWCLIPBOARD As Long = &H308&
Private Const WM_CHANGECBCHAIN As Long = &H30D&
Private Const WM_CLIPBOARDUPDATE As Long = &H31D&
Private Const WM_DESTROYCLIPBOARD As Long = &H307&

'Private Const defSnapGap As Long = 15

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Set defaults
   m_Enabled = defEnabled
   'Call SetHookClipboard
End Sub

Private Sub Class_Terminate()
   ' Tear down
   ChangeClipboardChain m_hWnd, mNextClipBoardViewerHWnd
   Call Unhook
End Sub

' *********************************************
'  Public Properties
' *********************************************
Public Property Let Enabled(ByVal NewValue As Boolean)
   m_Enabled = NewValue 'persist
End Property

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled  'report
End Property

Public Property Let hWnd(ByVal NewValue As Long)
   ' Clear any existing hook, first.
   Call Unhook
   ' Set hook into new window, if possible.
   If IsWindow(NewValue) Then
      If HookSet(NewValue, Me) Then
         m_hWnd = NewValue
      End If
   End If
   mNextClipBoardViewerHWnd = SetClipboardViewer(m_hWnd)
End Property

Public Property Get hWnd() As Long
   hWnd = m_hWnd  'report
End Property


' *********************************************
'  Private Methods
' *********************************************
'Private Sub SetHookClipboard()
'    'To register this form as a clipboard viewer...
''    mNextClipBoardViewerHWnd = SetClipboardViewer(m_hWnd)
'End Sub


Private Sub Unhook()
   ' Clear existing hook.
   If m_hWnd Then
      Call HookClear(m_hWnd, Me)
      m_hWnd = 0
   End If
End Sub

' *********************************************
'  Implemented Subclassing Interface
' *********************************************
Private Function IHookXP_Message(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
   'Dim pos As WINDOWPOS
   Dim EatIt As Boolean
   
   ' Special processing for messages we care about.
   Select Case uiMsg
      Case WM_DRAWCLIPBOARD 'The clipboard has changed...
          RaiseEvent Change("WM_DRAWCLIPBOARD", 0)
          SendMessage mNextClipBoardViewerHWnd, uiMsg, wParam, lParam
          EatIt = True

      Case WM_CLIPBOARDUPDATE
          RaiseEvent Change("WM_CLIPBOARDUPDATE", 0)
          SendMessage mNextClipBoardViewerHWnd, uiMsg, wParam, lParam
          EatIt = True

      Case WM_CHANGECBCHAIN 'Another clipboard viewer has removed itself...
         'If wParam = CType(mNextClipBoardViewerHWnd, IntPtr) Then
         
         'wParam ' is A handle to the window being removed from the clipboard viewer chain.
         RaiseEvent Change("WM_CHANGECBCHAIN", wParam)
         
         If wParam = mNextClipBoardViewerHWnd Then
            mNextClipBoardViewerHWnd = lParam
         Else
            SendMessage mNextClipBoardViewerHWnd, uiMsg, wParam, lParam
         End If
         EatIt = True
    
      Case WM_NCDESTROY
         Call Unhook  ' !!!
   End Select
   
   ' Pass back to default message handler.
   If EatIt = False Then
      IHookXP_Message = HookDefault(hWnd, uiMsg, wParam, lParam)
   End If
End Function



