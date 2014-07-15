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

Call Import("library.vbs")

'On Error Resume Next


If Err.Number <> 0 Then
   MsgBox "Unable to ding! Tell Jordan Pratt (jpratt@mtroyal.ca) that the Dinger is broken!"
Else
    Call makeSureHandDown()
End If


' ===================================
' = PROCEDURES
' ===================================

Sub makeSureHandDown()
    If (handIsUp()) Then 
        Call putHandDown()
        ringBell(dingDataEntity()) 
    End If
End Sub


'===================================================
'                 Import Code
'===================================================

Sub Import(strFile)
 
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set wshShell = CreateObject("Wscript.Shell")

    strFile = WshShell.ExpandEnvironmentStrings(strFile)
    strFile = objFs.GetAbsolutePathName(strFile)

    Set objFile = objFs.OpenTextFile(strFile)

    strCode = objFile.ReadAll

    objFile.Close

    ExecuteGlobal strCode
 
End Sub