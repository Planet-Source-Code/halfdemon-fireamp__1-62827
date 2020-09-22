Attribute VB_Name = "modRegistry"
' not yet implemented, probably defuct
Private Declare Function RegCreateKey& Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, lphKey As Long)
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpszSubKey As String, ByVal fdwType As Long, ByVal lpszValue As String, ByVal dwLength As Long)
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 256&
Private Const REG_SZ = 1

Public Sub registerType(fType As String, fDescription As String, fKey As String)
On Error GoTo e
    Dim sKeyName  As String  'Key
    Dim sKeyValue As String  'Key Value
    Dim ret       As Long    'error status
    Dim lphKey    As Long    'key handle
   
    'makes a root entry
    sKeyName = fKey
    sKeyValue = fDescription
    ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    ret = RegSetValue&(lphKey&, Empty, REG_SZ, sKeyValue, 0&)

    'makes an extention association
    sKeyName = fType
    sKeyValue = fKey
    ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    ret = RegSetValue&(lphKey, Empty, REG_SZ, sKeyValue, 0&)

    'command line for sKeyName
    sKeyName = "FireAMP"
     sKeyValue = App.Path & "\" & App.EXEName & ".exe %1"
     ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    ret = RegSetValue&(lphKey, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
    Exit Sub
e:

End Sub

