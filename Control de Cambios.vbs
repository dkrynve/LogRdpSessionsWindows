DIM textmsg, usuario, sistema, sujeto, usutxt, comtxt, cuerpo_mail



Set WshNetwork = WScript.CreateObject("WScript.Network")

usuario=WshNetwork.UserName
sistema=WshNetwork.ComputerName
usutxt=" Usuario  "
comtxt=" Equipo  "
textmsg=" -- "
'----------------------------------------------------------------------
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("c:\Windows\Control de Cambios\Texto_mail.txt", ForReading)
strResponses = objTextFile.ReadAll
cuerpo_mail=strResponses
'Wscript.Echo strResponses
objTextFile.Close



'---------------------------------------
sujeto="Monitoreo de Logueo " & usutxt & usuario &textmsg & comtxt & sistema

Set objMessage = CreateObject("CDO.Message") 
'objMessage.Subject = "Monitoreo de Logueo -> Logeo de Usuario" '& WshNetwork.UserName "Desde" & WshNetwork.UserName
objMessage.Subject = sujeto
objMessage.From = "argsaolog@turner.com" 
objMessage.To = "argentinasao@turner.com ; argsaolog@turner.com" 
objMessage.TextBody = strResponses

'==This section provides the configuration information for the remote SMTP server.
'==Normally you will only change the server name or IP.
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

'Name or IP of Remote SMTP Server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "ARGSMTP.TURNER.COM"

'Server port (typically 25)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 

objMessage.Configuration.Fields.Update

'==End remote SMTP server configuration section==
objMessage.AddAttachment "C:\Windows\Control de Cambios\info.txt"

objMessage.Send

objMessage.Configuration.Fields.Update

