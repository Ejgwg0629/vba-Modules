const KEY_QUERY_VALUE = &H0001
const KEY_SET_VALUE = &H0002
const KEY_CREATE_SUB_KEY = &H0004
const DELETE = &H00010000

const HKEY_LOCAL_MACHINE = &H80000002
const HKEY_CURRENT_USER = &H80000001

strComputer = "."
Set StdOut = WScript.StdOut

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Policies\Microsoft\office\16.0\outlook\security"
strValueName = "level"

oReg.GetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
'WScript.Echo "Current History Buffer Size: " & dwValue

setValue = 1

oReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,setValue
oReg.GetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
'WScript.Echo "Current History Buffer Size: " & dwValue

