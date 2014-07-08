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

'Const LISTENER_URL = "http://valid-hall-624.appspot.com"
Const HAND_UP = "Up"
Const HAND_DOWN = "Down"

Const LISTENER_URL = "http://localhost:9080" 

' On Error Resume Next
dda = dingDataEntity()
If Err.Number <> 0 Then
   MsgBox "Unable to ding! Tell Jordan Pratt (jpratt@mtroyal.ca) that the Dinger is broken!" & Err.Number
Else
    Call ringBell(dda) 
    Call toggleHandState()
End If


Function getOriginState()
    If (handIsUp()) Then getOriginState = HAND_UP Else getOriginState = HAND_DOWN
End Function

    
Function getDestState()
    If (handIsUp()) Then getDestState = HAND_DOWN Else getDestState = HAND_UP
End Function


Sub toggleHandState()
    Dim oldFileName, newFileName
    Call setNames(oldFileName, newFileName)
    Call renameFile(oldFileName, newFileName)
End Sub

Sub renameFile(oldFileName, newFileName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(oldFileName) Then
        Set stateFile = fso.GetFile(oldFileName)
        stateFile.Move(newFileName)
    Else
        fso.createTextFile(newFileName)
    End If
End Sub

Sub setNames(oldFileName, newFileName)
    oldFileName = "hand" + getOriginState() + ".state"
    newFileName = "hand" + getDestState() + ".state"
End Sub

Function handIsUp()
    Set fso = CreateObject("Scripting.FileSystemObject")
    handIsUp = fso.FileExists("handUp.state")
End Function

Sub ringBell(dingDataEntity)
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", LISTENER_URL, false
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp.send dingDataEntity
End Sub


Function dingDataEntity()
    dingDataEntity = "state=" + getDestState() + "&machine=" + machineName() + "&fullName=" + fullUserName() + "&userName=" + userName()
End Function


Function machineName()
    Set shell = WScript.CreateObject( "WScript.Shell" )
    machineName = shell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
End Function


Function fullUserName()
    fullUserName = user().Get("displayName")
End Function
        
Function userName()
    userName = user().Get("sAMAccountName")
End Function
        
Function user()
    Set adSysInfo = CreateObject("ADSystemInfo")
    uName = adSysInfo.UserName
    Set user = GetObject("LDAP://" & uName)
End Function