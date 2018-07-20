Attribute VB_Name = "modRegistry"
Option Explicit
'-------------------------------------------------------------------------------------------'
'   This Registry handler is developed by Ronald Kas (r.kas@kaycys.com)                     '
'   from Kaycys (http://www.kaycys.com).                                                    '
'                                                                                           '
'   You may use this Registry Handler for all purposes except from making profit with it.   '
'   Check our site regulary for updates.                                                    '
'-------------------------------------------------------------------------------------------'

' Declare Windows API functions...
Public Declare Function RegCloseKey _
               Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegCreateKeyEx _
               Lib "advapi32.dll" _
               Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                        ByVal lpSubKey As String, _
                                        ByVal Reserved As Long, _
                                        ByVal lpClass As String, _
                                        ByVal dwOptions As Long, _
                                        ByVal samDesired As Long, _
                                        ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                        ByRef phkResult As Long, _
                                        ByRef lpdwDisposition As Long) As Long

Public Declare Function RegDeleteKey _
               Lib "advapi32.dll" _
               Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                      ByVal lpSubKey As String) As Long

Public Declare Function RegEnumKeyEx _
               Lib "advapi32.dll" _
               Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                                      ByVal dwIndex As Long, _
                                      ByVal lpName As String, _
                                      ByRef lpcbName As Long, _
                                      ByVal lpReserved As Long, _
                                      ByVal lpClass As String, _
                                      ByRef lpcbClass As Long, _
                                      lpftLastWriteTime As FILE_TIME) As Long

Public Declare Function RegEnumValue _
               Lib "advapi32.dll" _
               Alias "RegEnumValueA" (ByVal hKey As Long, _
                                      ByVal dwIndex As Long, _
                                      ByVal lpValueName As String, _
                                      ByRef lpcbValueName As Long, _
                                      ByVal lpReserved As Long, _
                                      ByRef lpType As Long, _
                                      ByRef lpData As Any, _
                                      ByRef lpcbData As Long) As Long

Public Declare Function RegOpenKeyEx _
               Lib "advapi32.dll" _
               Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                      ByVal lpSubKey As String, _
                                      ByVal ulOptions As Long, _
                                      ByVal samDesired As Long, _
                                      ByRef phkResult As Long) As Long

Public Declare Function RegQueryValueEx _
               Lib "advapi32" _
               Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                         ByVal lpValueName As String, _
                                         ByVal lpReserved As Long, _
                                         ByRef lpType As Long, _
                                         ByVal lpData As String, _
                                         ByRef lpcbData As Long) As Long

Public Declare Function RegQueryInfoKey _
               Lib "advapi32.dll" _
               Alias "RegQueryInfoKeyA" (ByVal hKey As Long, _
                                         ByVal lpClass As String, _
                                         ByRef lpcbClass As Long, _
                                         ByVal lpReserved As Long, _
                                         ByRef lpcSubKeys As Long, _
                                         ByRef lpcbMaxSubKeyLen As Long, _
                                         ByRef lpcbMaxClassLen As Long, _
                                         ByRef lpcValues As Long, _
                                         ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As FILE_TIME) As Long

Public Declare Function RegSetValueExString _
               Lib "advapi32.dll" _
               Alias "RegSetValueExA" (ByVal hKey As Long, _
                                       ByVal lpValueName As String, _
                                       ByVal Reserved As Long, _
                                       ByVal dwType As Long, _
                                       ByVal lpValue As String, _
                                       ByVal cbData As Long) As Long

Public Declare Function RegSetValueExBoolean _
               Lib "advapi32" _
               Alias "RegSetValueExA" (ByVal hKey As Long, _
                                       ByVal lpValueName As String, _
                                       ByVal Reserved As Long, _
                                       ByVal dwType As Long, _
                                       ByRef lpData As Boolean, _
                                       ByVal cbData As Long) As Long

Public Declare Function RegSetValueExLong _
               Lib "advapi32.dll" _
               Alias "RegSetValueExA" (ByVal hKey As Long, _
                                       ByVal lpValueName As String, _
                                       ByVal Reserved As Long, _
                                       ByVal dwType As Long, _
                                       ByRef lpValue As Long, _
                                       ByVal cbData As Long) As Long

Public Declare Function RegDeleteValue _
               Lib "advapi32.dll" _
               Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String) As Long
                        
' Declare Windows API constants...
Public Const lngHKEY_CLASSES_ROOT = &H80000000

Public Const lngHKEY_CURRENT_USER = &H80000001

Public Const lngHKEY_LOCAL_MACHINE = &H80000002

Public Const lngHKEY_USERS = &H80000003

Public Const lngERROR_SUCCESS = 0&

Public Const lngERROR_FAILURE = 13&

Public Const lngUNREADABLE_NODE = 234&

Public Const lngNO_MORE_NODES = 259&

Public Const lngERROR_MORE_DATA = 234&

Public Const lngREG_OPTION_NON_VOLATILE = 0

Public Const lngSYNCHRONIZE = &H100000

Public Const lngSTANDARD_RIGHTS_READ = &H20000

Public Const lngKEY_QUERY_VALUE = &H1

Public Const lngKEY_ENUMERATE_SUB_KEYS = &H8

Public Const lngKEY_NOTIFY = &H10

Public Const lngKEY_SET_VALUE = &H2

Public Const lngKEY_CREATE_SUB_KEY = &H4

Public Const lngKEY_CREATE_LINK = &H20

Public Const lngSTANDARD_RIGHTS_ALL = &H1F0000

Public Const lngKEY_READ = ((lngSTANDARD_RIGHTS_READ Or lngKEY_QUERY_VALUE Or lngKEY_ENUMERATE_SUB_KEYS Or lngKEY_NOTIFY) And (Not lngSYNCHRONIZE))

Public Const lngKEY_ALL_ACCESS = ((lngSTANDARD_RIGHTS_ALL Or lngKEY_QUERY_VALUE Or lngKEY_SET_VALUE Or lngKEY_CREATE_SUB_KEY Or lngKEY_ENUMERATE_SUB_KEYS Or lngKEY_NOTIFY Or lngKEY_CREATE_LINK) And (Not lngSYNCHRONIZE))

Public Const lngREG_SZ = 1

Public Const lngREG_BINARY = 3

Public Const lngREG_DWORD = 4

Public Const ERROR_SUCCESS = 0&

' Declare Windows API types...
Public Type FILE_TIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type



