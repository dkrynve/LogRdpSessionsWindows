'Dim WshSysEnv, sClientName
'Set WshSysEnv = WshShell.Environment("PROCESS")
'sClientName = WshSysEnv("CLIENTNAME")
'WScript.Echo objShell.Environment("PROCESS").Item("COMPUTERNAME") 
'msgbox(sClientName)
Set wshShell = CreateObject( "WScript.Shell" )
WScript.Echo wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
'wshShell = Nothing