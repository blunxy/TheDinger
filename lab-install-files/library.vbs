Const LISTENER_URL = "http://valid-hall-624.appspot.com"
' Const LISTENER_URL = "http://localhost:8080" 

Const HAND_UP = "Up"
Const HAND_DOWN = "Down"
Const DESKTOP = &H10&
Const UP_ICON = "hand-active.ico"
Const DOWN_ICON = "hand-inactive.ico"

Const REG_KEY = "HKCU\HandsUp\"

HAND_STATUS_REG_VAL = REG_KEY + "handstatus"
GUID_REG_VAL = REG_KEY + "guid"

Set SHELL = WScript.CreateObject( "WScript.Shell" )


Function dingDataEntity()
    dingDataEntity = "state=" + getOriginalState() + "&machine=" + machineName() + "&fullName=" + fullUserName() + "&userName=" + userName() + "&guid=" + getGuid()
End Function


Function generateGuid()
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    guid = TypeLib.Guid
    guid = Replace(guid, "{", "")
    guid = Replace(guid, "}", "")
    guid = Replace(guid, "-", "")
    generateGuid = Left(guid, 10)
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
    On Error Resume Next
    fullUserName = user().Get("displayName")
    If Err.Number <> 0 Then
        fullUserName = "No Displayname"
    End If
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
    SHELL.RegWrite HAND_STATUS_REG_VAL, "T", "REG_SZ"
    SHELL.RegWrite GUID_REG_VAL, generateGuid(), "REG_SZ"
End Sub


Sub putHandDown()
    SHELL.RegWrite HAND_STATUS_REG_VAL, "F", "REG_SZ"
End Sub

    
Sub initRegistry()
    SHELL.RegWrite HAND_STATUS_REG_VAL, "F", "REG_SZ"
    SHELL.RegWrite GUID_REG_VAL, "", "REG_SZ"
End Sub

Function getOriginalState()
    If (handIsUp()) Then getOriginalState = HAND_UP Else getOriginalState = HAND_DOWN
End Function


Function handIsUp()
    On Error Resume Next
    regValue = SHELL.RegRead(REG_KEY)
    If Err.Number <> 0 Then
        Call initRegistry()
        Err.Clear
    End If
    If (SHELL.RegRead(HAND_STATUS_REG_VAL) = "T") Then handIsUp = True Else handIsUp = False
End Function

            
Function getGuid()
    getGuid = SHELL.RegRead(GUID_REG_VAL)
End Function

            
Sub toggleShortcutIcon()
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.NameSpace(DESKTOP)

    Set objFolderItem = objFolder.ParseName("HandsUp!.lnk")
    Set objShortcut = objFolderItem.GetLink

    objShortcut.SetIconLocation getFullTargetIconPath(),0
    objShortcut.Save
End Sub


Function getFullTargetIconPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    currDir = fso.GetAbsolutePathName(".") + "\"
    If (handIsUp()) Then getFullTargetIconPath = (currDir & UP_ICON) Else getFullTargetIconPath = (currDir & DOWN_ICON)
End Function