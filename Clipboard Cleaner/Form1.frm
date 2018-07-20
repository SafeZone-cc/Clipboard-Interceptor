VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Clipboard Cleaner :)"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Clipboard"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWndNewOwner As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long

Private Sub Command1_Click()
    If OpenClipboard(0&) Then
        Debug.Print EmptyClipboard
        'CloseClipboard
    End If
End Sub

