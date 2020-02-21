
'==================================================================
' control subsystem
'==================================================================
run32()
elevate()


'==================================================================
' Public variable definition
'==================================================================
Set MessageBox = CreateObject("SfcMini.DynaCall")
MessageBox.Declare "user32", "MessageBoxTimeoutA"

const REG_SZ = 1
const REG_EXPAND_SZ = 2
const REG_BINARY = 3
const REG_DWORD = 4
const REG_MULTI_SZ = 7

const HKEY_LOCAL_MACHINE = &H80000002
const HKEY_CURRENT_USER = &H80000001

const KEY_QUERY_VALUE = &H0001
const KEY_SET_VALUE = &H0002
const KEY_CREATE_SUB_KEY = &H0004
const DELETE = &H00010000

' my varialbe
strComputer = "."
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
strValueName = "AutoConfigURL"
strValue = "http://webproxy.ac.toyotasystems.com/proxy-ts.pac"

Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

' Go !!!
main()


'==================================================================
' do WHAT you like
'==================================================================
Sub main()
    if hasValue(HKEY_CURRENT_USER, strKeyPath, strValueName) then
        deleteValue HKEY_CURRENT_USER, strKeyPath, strValueName
        MessageBox 0, "disabled proxy.", "RayTool", vbInformation, 0, 900
    else
        rc = oReg.SetStringValue(HKEY_CURRENT_USER, strKeyPath, strValueName, strValue)
        if (rc <> 0) or (err.number <> 0) then
            wscript.echo "failed to enable proxy."
            wscript.quit
        else
            MessageBox 0, "enabled proxy: " & strValue, "RayTool", vbInformation, 0, 900
        end if
    end if
End Sub


'==================================================================
' common functions
'==================================================================
Sub createKey(def_key, sub_key)
    rc = oReg.CreateKey(def_key, sub_key)
    if (rc <> 0) or (err.number <> 0) then
        wscript.echo "failed to create key " & sub_key
        wscript.quit
    end if
End Sub

Function hasValue(def_key, sub_key, value_name)
    hasValue = false
    oReg.EnumValues def_key, sub_key, value_names, value_types
    for i = 0 to ubound(value_names)
        if value_names(i) = value_name then
            hasValue = true
            exit for
        end if
    next
End Function

Sub createDWORDValue(def_key, sub_key, value_name, value)
    oReg.SetDWORDValue def_key, sub_key, value_name, value
    if err.number <> 0 then
        wscript.echo "failed to create value " & value_name
        wscript.quit
    end if
End Sub

Sub deleteValue(def_key, sub_key, value_name)
    rc = oReg.DeleteValue(def_key, sub_key, value_name, value)
    if (rc <> 0) or (err.number <> 0) then
        wscript.echo "failed to delete value " & value_name
        wscript.quit
    end if
End Sub

Sub enumSubKey(def_key, sub_key)
    oReg.EnumValues def_key, sub_key, value_names, value_types
    For i = 0 To UBound(value_names)
        result = result & "Value Name: '" & value_names(i) & "', "
        Select Case value_types(i)
            Case REG_SZ
                result = result & "Data Type: String"
            Case REG_EXPAND_SZ
                result = result & "Data Type: Expanded String"
            Case REG_BINARY
                result = result & "Data Type: Binary"
            Case REG_DWORD
                result = result & "Data Type: DWORD"
            Case REG_MULTI_SZ
                result = result & "Data Type: Multi String"
        End Select
        result = result & vbCrLf
    Next
    wscript.echo result
End Sub


'==================================================================
' test subsystem
'==================================================================
Sub testEnumSubKey()
    enumSubkey HKEY_CURRENT_USER, strKeyPath
    wscript.quit
End Sub


'==================================================================
' runnable support subsystem
'==================================================================
Sub run32()
    'Author: Demon
    'Date: 2015/7/9
    'Website: http://demon.tw

    Dim strComputer, objWMIService, colItems, objItem, strSystemType
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
    
    For Each objItem in colItems
        strSystemType = objItem.SystemType
    Next
    
    If InStr(strSystemType, "x64") > 0 Then
        Dim fso, WshShell, strFullName
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set WshShell = CreateObject("WScript.Shell")
        strFullName = WScript.FullName
        If InStr(1, strFullName, "system32", 1) > 0 Then
            strFullName = Replace(strFullName, "system32", "SysWOW64", 1, 1, 1)
            WshShell.Run strFullName & " " &_
                """" & WScript.ScriptFullName & """", 10, False
            WScript.Quit
        End If
    End If
End Sub

Sub elevate()
    If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName _
        , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
    End If
End Sub

