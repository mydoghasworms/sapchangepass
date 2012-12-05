Set ctlLogon = CreateObject("SAP.LogonControl.1")
Set funcControl = CreateObject("SAP.Functions")
Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")

'''Obtain current and new password from user
currpass = InputBox("Enter current password")
newpass = InputBox("Enter new password")

''Initialize variables
Set outFile = objFileSystemObject.CreateTextFile("passchangelog.txt", True)

''' For each system, call subroutine to log on to system and change password
''' Parameters are: hostname, system id, system no., client, language, user
''' ADD A LINE PER SYSTEM FOR WHICH YOU WANT TO CHANGE YOUR PASSWORD
ChangePass "host1.company.com", "ABC", "00", "100", "EN", "MYUSER"
ChangePass "host1.company.com", "ABC", "00", "200", "EN", "MYUSER"
ChangePass "host2.company.com", "DEF", "01", "300", "EN", "MYUSER"
ChangePass "host3.company.com", "XYZ", "31", "400", "EN", "MYUSER"

''' Cleanup
outFile.Close
Set outFile = Nothing
Set ctlLogon = Nothing
Set funcControl = Nothing
Set objFileSystemObject = Nothing

'''***** Log on to system and change password *****
Sub ChangePass(appserver, sysid, sysno, client, lang, user)

''' Establish new connection
  Set objConnection = ctlLogon.NewConnection

''' Set logon details
  objConnection.ApplicationServer = appserver
  objConnection.System = sysid
  objConnection.SystemNumber = sysno
  objConnection.client = client
  objConnection.Language = lang
  objConnection.user = user
  objConnection.Password = currpass

''' Log on to system
  booReturn = objConnection.Logon(0, True)
  outFile.Write sysid & " " & client & ": "

''' Check if logon successful
  If booReturn <> True Then
    objConnection.LastError
    outFile.Write "Can't log on"
    Exit Sub
  Else
    outFile.Write "Login OK"
  End If

''' Prepare to call change password function
  funcControl.connection = objConnection
  Set CHPASS_FN = funcControl.Add("SUSR_USER_CHANGE_PASSWORD_RFC")
  Set expPassword = CHPASS_FN.Exports("PASSWORD")
  Set expNewPass = CHPASS_FN.Exports("NEW_PASSWORD")
  Set expFillRet = CHPASS_FN.Exports("USE_BAPI_RETURN")
  Set impReturn = CHPASS_FN.Imports("RETURN")
  expPassword.Value = currpass
  expNewPass.Value = newpass
  expFillRet.Value = "1"

''' Call change password function
  If CHPASS_FN.Call = True Then
    outFile.Write (", Called Function")
    Message = impReturn("MESSAGE")
    outFile.WriteLine " : " & Message
  Else
    outFile.Write (", Call to function failed")
  End If

  outFile.WriteLine vbNewLine
End Sub