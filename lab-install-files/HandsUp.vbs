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

' Const LISTENER_URL = "http://valid-hall-624.appspot.com"
Const LISTENER_URL = "http://localhost:9080" 

Const HAND_UP = "Up"
Const HAND_DOWN = "Down"
Const DESKTOP = &H10&
Const UP_ICON = "hand-active.ico"
Const DOWN_ICON = "hand-inactive.ico"

On Error Resume Next

Set SHELL = WScript.CreateObject( "WScript.Shell" )

dda = dingDataEntity()
If Err.Number <> 0 Then
   MsgBox "Unable to ding! Tell Jordan Pratt (jpratt@mtroyal.ca) that the Dinger is broken!"
Else
    Call ringBell(dda) 
    Call toggleHandState()
    Call toggleShortcutIcon()
End If


' ===================================
' = PROCEDURES
' ===================================


Function dingDataEntity()
    dingDataEntity = "state=" + getDestState() + "&machine=" + machineName() + "&fullName=" + fullUserName() + "&userName=" + userName()
End Function


Sub ringBell(dingDataEntity)
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", LISTENER_URL, false
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp.send dingDataEntity
End Sub


Function machineName()
    machineName = SHELL.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
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


Sub toggleHandState()
    If (handIsUp()) Then putHandDown() Else PutHandUp()
End Sub


Sub putHandUp()
    SHELL.RegWrite "HKCU\HandsUp\", "T", "REG_SZ"
End Sub


Sub putHandDown()
    SHELL.RegWrite "HKCU\HandsUp\", "F", "REG_SZ"
End Sub


Function getDestState()
    If (handIsUp()) Then getDestState = HAND_DOWN Else getDestState = HAND_UP
End Function


Function handIsUp()
    On Error Resume Next
    regValue = SHELL.RegRead("HKCU\HandsUp\")
    If Err.Number <> 0 Then
        Call putHandDown()
        Err.Clear
    End If
    If (SHELL.RegRead("HKCU\HandsUp\") = "T") Then handIsUp = True Else handIsUp = False
End Function


Sub toggleShortcutIcon()
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.NameSpace(DESKTOP)

    Set objFolderItem = objFolder.ParseName("Ding!.lnk")
    Set objShortcut = objFolderItem.GetLink

    objShortcut.SetIconLocation getFullTargetIconPath(),0
    objShortcut.Save
End Sub


Function getFullTargetIconPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    currDir = fso.GetAbsolutePathName(".") + "\"
    If (handIsUp()) Then getFullTargetIconPath = (currDir & UP_ICON) Else getFullTargetIconPath = (currDir & DOWN_ICON)
End Function

            
