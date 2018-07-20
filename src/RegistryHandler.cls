VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------------------'
'   This Registry handler is developed by Ronald Kas (r.kas@kaycys.com)                     '
'   from Kaycys (http://www.kaycys.com).                                                    '
'                                                                                           '
'   You may use this Registry Handler for all purposes except from making profit with it.   '
'   Check our site regulary for updates.                                                    '
'-------------------------------------------------------------------------------------------'

' Modified by Dragokas

Enum HKEYS
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Public Sub DeleteKey(RootKey As HKEYS, _
                     sKey As String)
    Dim x As Variant
    Dim i As Integer
    Dim OrgKey As String
    OrgKey = sKey

    Do
        x = EnumKeys(RootKey, sKey)

        If IsEmpty(x) = False Then
            If UBound(x) <= 0 Then
                DeleteRegistryKey RootKey, sKey

                For i = Len(sKey) To 1 Step -1

                    If Mid$(sKey, i, 1) = "\" Then
                        sKey = Left$(sKey, i - 1)
                        Exit For
                    End If

                Next

                If Len(sKey) < Len(OrgKey) Then
                    Exit Sub
                End If

            Else
                sKey = sKey & "\" & x(0)
            End If

        Else
            Exit Do
        End If

    Loop

End Sub

Public Sub DeleteValue(RootKey As HKEYS, _
                       sKey As String, _
                       sValueName As String)
    Dim lHandle As Long
    Call OpenRegistryKey(RootKey, sKey, lHandle)
    RegDeleteValue lHandle, sValueName
    CloseRegistryKey lHandle

End Sub

'Public Function EnumKeys(RootKey As HKEYS, _
'                         sKey As String) As Variant
Public Function EnumKeys(RootKey As HKEYS, sKey As String) As Variant
    MsgBox ""
    Dim hKeyHandle As Long
    Dim x As Variant
    Call RegOpenKeyEx(RootKey, sKey, 0&, lngKEY_ALL_ACCESS, hKeyHandle)
    Call EnumerateRegistryKeys(hKeyHandle, x)
    Call RegCloseKey(hKeyHandle)
    EnumKeys = x
End Function




'Public Function EnumValues(RootKey As HKEYS, _
'                           sKey As String) As Variant
Public Function EnumValues(sKey As String) As Variant
    
    Dim RootKey As HKEYS
    Dim vEnum As Variant
    Dim lHandle As Long
    Dim vRet As Variant
    Dim i As Integer
    
    Call OpenRegistryKey(RootKey, sKey, lHandle)
    EnumerateRegistryValuesByHandle lHandle, vEnum
    Call CloseRegistryKey(lHandle)
    ReDim vRet(UBound(vEnum, 2))

    For i = 0 To UBound(vEnum, 2)
        vRet(i) = vEnum(1, i)
    Next

    EnumValues = vRet

End Function

Private Function GetRootKey(sKey) As HKEYS
    'Dim sKey$
    

End Function

Public Function GetRegValue(RootKey As HKEYS, _
                            sKey As String, _
                            sValueName As String) As Variant
    Dim hKeyHandle As Long
    Dim vTemp As Variant
    Dim lngRet As Long

    lngRet = RegOpenKeyEx(RootKey, sKey, 0&, lngKEY_ALL_ACCESS, hKeyHandle)

    If lngRet = ERROR_SUCCESS Then
        ReadRegistryValue hKeyHandle, sValueName, vTemp
        Call RegCloseKey(hKeyHandle)
    Else
        vTemp = vbNullString
    End If

    GetRegValue = vTemp

End Function

Public Function WriteRegValue(ByVal vhRootKey As HKEYS, _
                              ByVal vstrKeyName As String, _
                              ByVal vstrValueName As String, _
                              ByVal vvntValue As Variant) As String
    Dim hKeyHandle As Long
  
    ' Call each of the neccessary functions in turn starting with OpenRegistryKey...
    Call OpenRegistryKey(vhRootKey, vstrKeyName, hKeyHandle)
    Call CreateRegistryKey(vhRootKey, vstrKeyName, hKeyHandle)
  
    ' then Write the value...
    Call WriteRegistryValue(hKeyHandle, vstrValueName, vvntValue)

    ' and close the opened key.
    Call CloseRegistryKey(hKeyHandle)

End Function

'------------------------------------------------------------------

Private Function CloseRegistryKey(ByVal vhKeyHandle As Long) As Boolean
    Dim lngReturn As Long
    
    lngReturn = RegCloseKey(vhKeyHandle)

    If lngReturn <> lngERROR_SUCCESS Then
        CloseRegistryKey = False
    Else
        CloseRegistryKey = True
    End If

End Function

Private Function CreateRegistryKey(ByVal vhKeyHandle As Long, _
                                   ByVal vstrKeyName As String, _
                                   ByRef rhNewKeyHandle As Long) As Long
    Dim lngReturn As Long, lngDepth As Long
    Dim typSecurityAttributes As SECURITY_ATTRIBUTES
  
    ' and then create the key.
    typSecurityAttributes.nLength = 50
    typSecurityAttributes.lpSecurityDescriptor = 0
    typSecurityAttributes.bInheritHandle = True
    lngReturn = RegCreateKeyEx(vhKeyHandle, vstrKeyName, 0, lngREG_SZ, lngREG_OPTION_NON_VOLATILE, lngKEY_ALL_ACCESS, typSecurityAttributes, rhNewKeyHandle, lngDepth)

    If lngReturn <> lngERROR_SUCCESS Then
    
    End If

    Exit Function

End Function

Private Function DeleteRegistryKey(ByVal vhKeyHandle As Long, _
                                   ByVal vstrKeyName As String) As String
    Dim lngReturn As Long
  
    ' and then delete the key.
    lngReturn = RegDeleteKey(vhKeyHandle, vstrKeyName)
End Function

Private Function EnumerateRegistryKeys(ByVal vhKeyHandle As Long, _
                                       ByRef rvntKeys As Variant) As String

    Dim strValue As String, strClass As String
    Dim lngDataLen As Long, lngValueLen As Long, lngReturn As Long, lngIndex As Long
    Dim lngClass As Long
    Dim strNodes() As String
    Dim typFileTime As FILE_TIME

    lngIndex = 0

    ' then loop through the nodes under the 'base node'...
    Do
        lngValueLen = 2000
        strValue = String$(lngValueLen, 0)
        lngDataLen = 2000

        ' and read the names of all the nodes under it...
        lngReturn = RegEnumKeyEx(vhKeyHandle, lngIndex, strValue, lngValueLen, 0&, strClass, lngClass, typFileTime)
        strValue = Left$(strValue, lngValueLen)

        '   checking for problems.
        If lngReturn <> lngERROR_SUCCESS And lngReturn <> lngNO_MORE_NODES Then
            ReDim Preserve strNodes(0)
            strNodes(0) = vbNullString
            Exit Do
        End If

        ' Add each node into an array...
        If lngReturn = lngERROR_SUCCESS Then
            ReDim Preserve strNodes(lngIndex)
            strNodes(lngIndex) = strValue
            lngIndex = lngIndex + 1
        End If

        ' and loop until the enumeration return fails.
    Loop While lngReturn <> lngNO_MORE_NODES

    rvntKeys = strNodes()
    Erase strNodes

End Function

Private Function EnumerateRegistryValuesByHandle(ByVal vhKeyHandle As Long, _
                                                 ByRef rvntValues As Variant) As String
    Dim strValue As String
    Dim lngData As Long, lngDataLen As Long, lngValueLen As Long, lngReturn As Long, lngIndex As Long
    Dim lngValueType As Long
    Dim strNodes() As String
    
    ' then loop through the nodes under the 'base node'...
    Do
        lngValueLen = 2000
        strValue = String$(lngValueLen, 0)
        lngDataLen = 2000
    
        ' and read the names of all the nodes under it...
        lngReturn = RegEnumValue(vhKeyHandle, lngIndex, ByVal strValue, lngValueLen, 0&, lngValueType, ByVal lngData, lngDataLen)
        strValue = Left$(strValue, lngValueLen)
      
        ' checking for problems.
        If lngReturn <> lngERROR_SUCCESS And lngReturn <> lngNO_MORE_NODES Then
        End If
      
        ' Add each node into an array...
        ReDim Preserve strNodes(0 To 1, 0 To lngIndex)
        strNodes(0, lngIndex) = CStr(lngValueType)
        strNodes(1, lngIndex) = strValue
        lngIndex = lngIndex + 1
      
        ' and loop until the enumeration return fails.
    Loop While lngReturn <> lngNO_MORE_NODES

    rvntValues = strNodes()
    Erase strNodes
End Function

Private Function OpenRegistryKey(ByVal vhRootKey As Long, _
                                 ByVal vstrKeyName As String, _
                                 ByRef rhKeyHandle As Long) As String
    Dim lngReturn As Long, hKeyHandle As Long
    
    ' then open the passed registry node (key) in the passed root key...
    lngReturn = RegOpenKeyEx(vhRootKey, vstrKeyName, 0&, lngKEY_ALL_ACCESS, hKeyHandle)

    If lngReturn <> lngERROR_SUCCESS Then
    End If
    
    rhKeyHandle = hKeyHandle
End Function

Private Function ReadRegistryValue(ByVal vhKeyHandle As Long, _
                                   ByVal vstrValueName As String, _
                                   ByRef rvntValue As Variant) As String
    Dim strValueName As String, strData As String
    Dim lngReturn As Long, lngIndex As Long, lngValuesCount As Long, lngValueType As Long, lngValueLen As Long
    Dim lngValueMax As Long, lngData As Long, lngDataLen As Long
    Dim blnData As Boolean
    Dim vntValues As Variant
    Dim typFileTime As FILE_TIME
  
    ' Check that all required variables have been passed...
    If vhKeyHandle <= 0 Then
    End If

    If vstrValueName = "" Then
    End If
  
    ' and enumerate the keys to see what type of value is stored in the one to return. First get the number of values
    ' and the maximum name length of those stored in the passed key...
    lngReturn = RegQueryInfoKey(vhKeyHandle, "", 0&, 0&, 0&, 0&, 0&, lngValuesCount, lngValueMax, 0&, 0&, typFileTime)

    If lngReturn <> lngERROR_SUCCESS Then
    End If

    lngValueLen = Len(vstrValueName) + 1
  
    ' then loop through the values until the requested value name is found.
    Call EnumerateRegistryValuesByHandle(vhKeyHandle, vntValues)

    For lngIndex = 0 To UBound(vntValues, 2)
        lngReturn = lngERROR_FAILURE
        strValueName = vntValues(1, lngIndex)
    
        ' Check that the currently enumerated key is the one requested...
        If LCase$(vstrValueName) = LCase$(strValueName) Then
            lngValueType = vntValues(0, lngIndex)
            lngValueLen = Len(strValueName)
      
            ' and, depending on the value type, read and return the stored value...
            Select Case lngValueType

                Case lngREG_BINARY
    
                    ' it's a binary value...
                    lngDataLen = 1
                    lngReturn = RegEnumValue(vhKeyHandle, lngIndex, strValueName, lngValueLen, 0&, lngValueType, blnData, lngDataLen)
                    rvntValue = blnData
                    Exit For

                Case lngREG_DWORD
    
                    ' it's a DWord...
                    lngDataLen = 4
                    lngReturn = RegEnumValue(vhKeyHandle, lngIndex, strValueName, lngValueLen, 0&, lngValueType, lngData, lngDataLen)
                    rvntValue = lngData
                    Exit For

                Case lngREG_SZ
    
                    ' it's a string value.
                    lngDataLen = 2048
                    strData = String$(lngDataLen, 0)
                    lngReturn = RegQueryValueEx(vhKeyHandle, strValueName, 0&, lngValueType, strData, lngDataLen)
                    rvntValue = Left$(strData, lngDataLen - 1)
                    Exit For
            End Select

        End If

    Next

    If lngReturn <> lngERROR_SUCCESS And lngReturn <> lngERROR_MORE_DATA Then
    End If

End Function

Private Function ReadValue(ByVal vhRootKey As Long, _
                           ByVal vstrKeyName As String, _
                           ByVal vstrValueName As String, _
                           ByRef rvntValue As Variant, _
                           Optional ByVal vvntDefault As Variant) As String
    Dim strReturn As String, strLanguageOffset As String, strMessage As String, strError As String
    Dim hKeyHandle As Long, lngReturn As Long, lngValueType As Long, hNewKeyHandle As Long
 
    ' Call each of the neccessary functions in turn starting with OpenRegistryKey...
    Do
        Call OpenRegistryKey(vhRootKey, vstrKeyName, hKeyHandle)

        If Not IsMissing(vvntDefault) Then
            Call CreateRegistryKey(vhRootKey, vstrKeyName, hKeyHandle)
            GoSub ValueWrite
        End If
  
        ' then read the value...
        strReturn = ReadRegistryValue(hKeyHandle, vstrValueName, rvntValue)

        If strReturn <> "" Then
            If Not IsMissing(vvntDefault) And rvntValue = "" Or rvntValue = 0 Then
                GoSub ValueWrite
            Else
                ReadValue = strReturn
                Exit Function
            End If

        Else
            Exit Do
        End If
  
        ' and close the opened key.
        Call CloseRegistryKey(hKeyHandle)
    Loop
  
    ' and close the opened key.
    Call CloseRegistryKey(hKeyHandle)
    Exit Function

ValueWrite:
    strReturn = WriteRegistryValue(hKeyHandle, vstrValueName, vvntDefault)

    If strReturn <> "" Then
        ReadValue = strReturn
        Exit Function
    End If

    Return
End Function

Private Function WriteRegistryValue(ByVal vhKeyHandle As Long, _
                                    ByVal vstrValueName As String, _
                                    ByVal vvntValue As Variant) As String
    Dim strValue As String
    Dim lngReturn As Long, lngValue As Long, lngLength As Long
    Dim blnValue As Boolean
  
    ' Check that all passed parameters are filled...
  
    ' and then write the value to the Value.
    Select Case VarType(vvntValue)

        Case vbString
            strValue = vvntValue & Chr$(0)
            lngLength = Len(strValue)
            lngReturn = RegSetValueExString(vhKeyHandle, vstrValueName, 0&, lngREG_SZ, strValue, lngLength)

        Case vbBoolean
            blnValue = CBool(vvntValue)
            lngReturn = RegSetValueExBoolean(vhKeyHandle, vstrValueName, 0&, lngREG_BINARY, blnValue, 1&)

        Case vbByte, vbInteger, vbLong
            lngValue = CLng(vvntValue)
            lngReturn = RegSetValueExLong(vhKeyHandle, vstrValueName, 0&, lngREG_DWORD, lngValue, 4&)

        Case Else
    End Select

    If lngReturn <> lngERROR_SUCCESS Then
    End If

End Function




