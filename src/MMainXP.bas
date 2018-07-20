Attribute VB_Name = "MMain"
' *************************************************************************
'  Copyright ©2009 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Sub Main()
   Dim frm As Form
   
   ' Silly XP Games.
   Call InitCommonControls
   
   ' Display a couple instances of main form.
   Set frm = New FHookXP
   frm.Show
'   Set frm = New FHookXP
'   frm.Show
End Sub

