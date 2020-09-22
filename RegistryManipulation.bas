Attribute VB_Name = "RegistryManipulation"
Option Explicit
'ABREVIATIONS
'HKEY_CURRENT_USER          HKCU
'HKEY_LOCAL_MACHINE         HKLM
'HKEY_CLASSES_ROOT          HKCR
'HKEY_USERS                 HKEY_USERS
'HKEY_CURRENT_CONFIG        HKEY_CURRENT_CONFIG

'REGREAD      objectName.RegRead(strName)
'type           description                                         in the form of
'REG_SZ         A string                                            A string
'REG_DWORD      A number                                            An integer
'REG_BINARY     A binary value                                      A VBArray of integers
'REG_EXPAND_SZ  An expandable string(e.g., "%windir%\\calc.exe")    A string
'REG_MULTI_SZ   An array of strings                                 A VBArray of strings
'examples
'Dim WshShell, bKey
'Set WshShell = WScript.CreateObject("WScript.Shell")
'
'WshShell.RegWrite "HKCU\Software\ACME\FortuneTeller\", 1, "REG_BINARY"
'WshShell.RegWrite "HKCU\Software\ACME\FortuneTeller\MindReader", "Goocher!", "REG_SZ"
'
'bKey = WshShell.RegRead("HKCU\Software\ACME\FortuneTeller\")
'WScript.Echo WshShell.RegRead("HKCU\Software\ACME\FortuneTeller\MindReader")
'
'WshShell.RegDelete "HKCU\Software\ACME\FortuneTeller\MindReader"
'WshShell.RegDelete "HKCU\Software\ACME\FortuneTeller\"
'WshShell.RegDelete "HKCU\Software\ACME\"

'REGWRITE     objectName.RegWrite(strName, anyValue [,strType])
'String       REG_SZ
'String       REG_EXPAND_SZ
'Integer      REG_DWORD
'Integer      REG_BINARY

'REGDELETE    object.RegDelete(strName)


Public WShell As Object 'object used to manipulate registry
Public J As Integer     'counter

Public Sub CreateObj()
  Set WShell = CreateObject("WScript.Shell")
End Sub

Public Sub DeleteObj()
  Set WShell = Nothing
End Sub

Public Function WSRead(mValue As String) As String
  WSRead = WShell.RegRead(mValue)
End Function

Public Sub WSDelete(mValue As String)
  WShell.RegDelete (mValue)
End Sub

Public Sub WSWrite(mKey As String, mValue As String)
  WShell.RegWrite mKey, mValue
End Sub
