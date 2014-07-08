'
' Sends the currently logged in user's LDAP display name
' and machine name to a listener on Google App Engine (GAE).
'
' This data will be used by the listener to broadcast
' a help request message on any listening client.
'
' Jordan Pratt
' 2014-JUL-07
'

Const LISTENER_URL = "http://valid-hall-624.appspot.com"
'Const LISTENER_URL = "http://localhost:9080" 

On Error Resume Next
dda = dingDataEntity()
If Err.Number <> 0 Then
   MsgBox "Unable to ding! Tell Jordan Pratt (jpratt@mtroyal.ca) that the Dinger is broken!" & Err.Number
Else
   Call ringBell(dda) 
End If


Sub ringBell(dingDataEntity)
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", LISTENER_URL, false
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp.send dingDataEntity
End Sub


Function dingDataEntity()
    dingDataEntity = "machine=" + machineName() + "&user=" + fullUserName()
End Function


Function machineName()
    Set shell = WScript.CreateObject( "WScript.Shell" )
    machineName = shell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
End Function


Function fullUserName()
    Set adSysInfo = CreateObject("ADSystemInfo")
    userName = adSysInfo.UserName
    Set user = GetObject("LDAP://" & userName)
    fullUserName = user.Get("displayName")
End Function