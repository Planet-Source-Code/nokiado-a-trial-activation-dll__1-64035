Attribute VB_Name = "modRegCode"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Registry API
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'HKEY Constants
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

'KEY Access Constants
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_EVENT = &H1     '  Event contains key event record
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'ERROR Constants
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_MORE_DATA = 234 'dderror

'REG Data Constants
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Public Const REG_OPTION_VOLATILE = 1

'TYPES Declaration
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Public Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type

Public Termino As Boolean
Public PreviaInstancia As Boolean



Public Function GenCode(ByVal strUN As String) As String
Dim P1 As Long, P2 As Long, P3 As Long
Dim S1 As String, S2 As String, S3 As String
Dim j As Integer

For j = 1 To Len(strUN)
    P1 = P1 + Asc(Mid(strUN, j, 1)) * 65
Next
strUN = LCase(strUN)
For j = 1 To Len(strUN)
    P2 = P2 + Asc(Mid(strUN, j, 1)) * 50
Next
strUN = UCase(strUN)
For j = 1 To Len(strUN)
    P3 = P3 + Asc(Mid(strUN, j, 1)) * 75
Next

S1 = CStr(Hex(P1))
S2 = CStr(Hex(P2))
S3 = CStr(Hex(P3))

If Len(S1) > 4 Then S1 = Left(S1, 4)
If Len(S2) > 4 Then S2 = Left(S2, 4)
If Len(S3) > 4 Then S3 = Left(S3, 4)

GenCode = S1 & "-" & S2 & "-" & S3
End Function

Public Function MakeRegEntries(ByVal strRC As String)
'This function will make registration entries
'TODO
Dim ret As Long, hKey As Long, dispo As Long
Dim strValueName As String
Dim SA As SECURITY_ATTRIBUTES

ret = RegCreateKeyEx(HKEY_LOCAL_MACHINE, _
"Software\Dalpcorp\Registro", 0, vbNullString, _
REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, hKey, dispo)

If ret = ERROR_SUCCESS Then
    strValueName = "RegCode"
    ret = RegSetValueEx(hKey, ByVal strValueName, 0, REG_SZ, _
    ByVal strRC & vbNullChar, Len(strRC))
    ret = RegCloseKey(hKey)
End If
End Function

Public Function getComputerID() As String
Dim fso, d
Set fso = CreateObject("Scripting.FileSystemObject")
Set d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName("C:\")))
getComputerID = d.SerialNumber
Set fso = Nothing
Set d = Nothing
End Function

Private Function setInstallDate()
'This function sets installation date
Dim ret As Long, hKey As Long, dispo As Long
Dim strValueName As String, strIDate As String
Dim SA As SECURITY_ATTRIBUTES

strIDate = Now 'Today

ret = RegCreateKeyEx(HKEY_LOCAL_MACHINE, _
"Software\SystemDate", 0, vbNullString, _
REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, hKey, dispo)

If ret = ERROR_SUCCESS Then
    strValueName = "InstallDate"
    ret = RegSetValueEx(hKey, ByVal strValueName, 0, REG_SZ, _
    ByVal strIDate & vbNullChar, Len(strIDate))
    ret = RegCloseKey(hKey)
End If
End Function

Private Function getInstallDate() As String
'This function returns install date
Dim ret As Long, hKey As Long
Dim strValueName As String
strValueName = "InstallDate"
ret = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\SystemDate", 0, KEY_READ, hKey)
If ret = ERROR_SUCCESS Then
    Dim lngData As Long
    lngData = 255
    strValueName = strValueName & vbNullChar
    getInstallDate = Space(lngData)
    ret = RegQueryValueEx(hKey, ByVal strValueName, 0, REG_SZ, ByVal getInstallDate, lngData)
    ret = RegCloseKey(hKey)
    getInstallDate = Trim(Left(getInstallDate, lngData - 1))
End If
End Function

Public Function IsRegistered() As Boolean
'Retunns true if registered
Dim ret As Long, hKey As Long
Dim strValueName As String, strVal As String

'Try to read registraion entry
strValueName = "RegCode": strVal = ""
ret = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Dalpcorp\Registro", 0, KEY_READ, hKey)
If ret = ERROR_SUCCESS Then
    Dim lngData As Long
    lngData = 255
    strValueName = strValueName & vbNullChar
    strVal = Space(lngData)
    ret = RegQueryValueEx(hKey, ByVal strValueName, 0, REG_SZ, ByVal strVal, lngData)
    ret = RegCloseKey(hKey)
    strVal = Left(strVal, lngData - 1)
End If

If strVal = "" Then 'No registration entry
    IsRegistered = False
    Exit Function
End If

'Reg entry found, now check validity
strVal = UCase(strVal)
If GenCode(getComputerID()) <> strVal Then
    IsRegistered = False 'Wrong Code
    Exit Function
Else
    IsRegistered = True 'Registration OK
End If
End Function

Private Function setLastUse()
'This function sets usage date
Dim ret As Long, hKey As Long, dispo As Long
Dim strValueName As String, strIDate As String
Dim SA As SECURITY_ATTRIBUTES

strIDate = Now 'Today

ret = RegCreateKeyEx(HKEY_LOCAL_MACHINE, _
"Software\SystemDate", 0, vbNullString, _
REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, hKey, dispo)

If ret = ERROR_SUCCESS Then
    strValueName = "UseDate"
    ret = RegSetValueEx(hKey, ByVal strValueName, 0, REG_SZ, _
    ByVal strIDate & vbNullChar, Len(strIDate))
    ret = RegCloseKey(hKey)
End If
End Function

Private Function getLastUse() As String
'This function returns use date
Dim ret As Long, hKey As Long
Dim strValueName As String
strValueName = "UseDate"
ret = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\SystemDate", 0, KEY_READ, hKey)
If ret = ERROR_SUCCESS Then
    Dim lngData As Long
    lngData = 255
    strValueName = strValueName & vbNullChar
    getLastUse = Space(lngData)
    ret = RegQueryValueEx(hKey, ByVal strValueName, 0, REG_SZ, ByVal getLastUse, lngData)
    ret = RegCloseKey(hKey)
    getLastUse = Trim(Left(getLastUse, lngData - 1))
End If
End Function

Public Function getTrialDays() As Integer
'Returns no of days left in trial
On Error GoTo ErrHandler
Dim strIDate As String, strUDate As String

strIDate = getInstallDate 'Installation Date
strUDate = getLastUse 'Last used

'Not found new installation
If strIDate = "" And strUDate = "" Then
    getTrialDays = 30
    Call setInstallDate
    Call setLastUse
    Exit Function
End If

'Cheating
If (strIDate <> "" And strUDate = "") Then
    getTrialDays = 0
    Exit Function
End If

If (strIDate = "" And strUDate <> "") Then
    getTrialDays = 0
    Exit Function
End If
strIDate = FormatDateTime(strIDate, vbShortDate)
strUDate = FormatDateTime(strUDate, vbShortDate)

'Clock Reset
If Int(CDate(strUDate)) > Int(Now) Then
ErrHandler: 'On all errors set trial expired
    getTrialDays = 0
    Exit Function
End If

Call setLastUse
getTrialDays = 30 - (Int(Now) - Int(CDate(strIDate)))
End Function
