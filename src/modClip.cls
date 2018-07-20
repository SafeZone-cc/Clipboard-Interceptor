VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "modClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vTypeGet As Long
Private vTypeSet As Long

'Enum Types
'
'vbCFBitmap
'vbCFDIB
'vbCFEMetafile
'vbCFFiles
'vbCFLink
'vbCFMetafile
'vbCFPalette
'vbCFRTF
'vbCFText

Private Sub Class_Initialize()
    vTypeGet = vbCFText
    vTypeSet = vbCFText
End Sub

'Type which we need to Get from ClipBoard
Public Property Let TypeGet(n As Long)
    vTypeGet = n
End Property

Public Property Get TypeGet() As Long
    TypeGet = vTypeGet
End Property

'Type of ClipBoard Data which we need to Set with value
Public Property Let TypeSet(n As Long)
    vTypeSet = n
End Property

Public Property Get TypeSet() As Long
    TypeSet = vTypeSet
End Property

'Put Data into ClipBoard
Public Property Let ClipPut(st As String)
    Clipboard.SetText st, vTypeSet
End Property

'Gain Data from Clipboard
Public Property Get ClipGet()
    If Clipboard.GetFormat(vTypeGet) Then
        ClipGet = Clipboard.GetText(vTypeGet)
    End If
End Property
