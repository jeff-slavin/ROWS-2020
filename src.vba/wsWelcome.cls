Option Explicit

Private Sub cmdLogin_Click()
'Attempts to login the user
'If successful then shows the "Logout" section and takes them to the dashboard worksheet
'If successful but the password is marked temporary, then shows the "TempPassword" section

'First perform pre-checks on the username/password

     Dim msgResponse As Messages
     Dim sUsername As String
     Dim sPassword As String
     Dim sTempPassword As String
     Dim sNewPassword As String
     Dim sRetypeNewPassword As String
     
     sUsername = Worksheets("Welcome").Range(gc_wsWelcomeUsername).Value
     sPassword = Worksheets("Welcome").Range(gc_wsWelcomePassword).Value

     'Check username was entered
     If Len(sUsername) < 1 Then
          Worksheets("Welcome").Range(gc_wsWelcomeLoginError).Value = "Please enter a username."
          Exit Sub
     End If
     
     'Check password was entered
     If Len(sPassword) < 1 Then
          Worksheets("Welcome").Range(gc_wsWelcomeLoginError).Value = "Please enter a password."
          Exit Sub
     End If
     
     'A username and password have been entered, now attempt to login
     
     'First check our global instance, if this is nothing then reset it
     'If it's not nothing, then let's log it out to be safe
     
     If Not g_cUser Is Nothing Then
          'Global instance is alive, so let's reset it to log it out
          Call g_cUser.Reset
     Else
          'Global instance is dead, so let's set it up (will automatically bet set to logged out and blank)
          Set g_cUser = New clsUser
     End If
     
     'Attempt Login
     msgResponse = g_cUser.Login(sUsername, sPassword)
     
     If msgResponse = Messages.msgTrue Then
          'Successful login
          
          'Show Logout section of Welcome worksheet
          'Take user to Dashboard worksheet (keeping Welcome worksheet visible/selectable)
          Call wsWelcome_ShowSection("Logout")
          Call Worksheets_ShowWelcomeAndOne("Dashboard")
          
          'Display Logged In Information In the Logout Section of the Welcome worksheet
          Worksheets("Welcome").Range(gc_wsWelcomeLoggedInUsername).Value = g_cUser.Username
          Worksheets("Welcome").Range(gc_wsWelcomeLoggedInRole).Value = g_cUser.Role
     ElseIf msgResponse = msgTemporaryPassword Then
          'User needs to reset their password (as current password is a temp password)
          MsgBox "Your password is marked as temporary. You are required to change your password to login.", vbOKOnly, "Temporary Password Reset"
          
          Call wsWelcome_ShowSection("TempPassword")
          Call Worksheets_ShowWelcomeAndOne("Welcome")
          
     Else
          'Display error depending on the message returned
          
          Select Case msgResponse
          
               Case Messages.msgFailedDBConnection
                    Worksheets("Welcome").Range(gc_wsWelcomeLoginError).Value = "Database connection error."
               Case Messages.msgFailedQuery
                    Worksheets("Welcome").Range(gc_wsWelcomeLoginError).Value = "Database query error."
               Case Messages.msgFalse
                    Worksheets("Welcome").Range(gc_wsWelcomeLoginError).Value = "Invalid username/password combination."
          End Select
          
     End If
     

End Sub

Private Sub cmdLogout_Click()

     'Check our global clsUser instance
     If Not g_cUser Is Nothing Then
          Call g_cUser.Reset
     Else
          Set g_cUser = New clsUser
     End If
     
     'Hide all worksheets except for the 'Welcome' worksheet and show the Login section
     Call wsWelcome_ShowSection("Login")
     Call Worksheets_ShowWelcomeAndOne("Welcome")
     
End Sub

Private Sub cmdUpdatePassword_Click()
'Forces the user to update their password
'Checks if they entered the correct temp password again
'Checks their inputs
'Updates the password in the system
'Logs the user in if successful

     Dim sTempPassword As String
     Dim sNewPassword As String
     Dim sRetypeNewPassword As String
     Dim msgResponse As Messages
     
     sTempPassword = Worksheets("Welcome").Range(gc_wsWelcomeTempPassword).Value
     sNewPassword = Worksheets("Welcome").Range(gc_wsWelcomeNewPassword).Value
     sRetypeNewPassword = Worksheets("Welcome").Range(gc_wsWelcomeRetypePassword).Value
     
     'First check if the user entered everything we need
     
     'Did they enter in their temp password
     If Len(sTempPassword) < 1 Then
          Worksheets("Welcome").Range(gc_wsWelcomeTempPasswordError).Value = "Please enter in your temporary password."
          Exit Sub
     End If
     
     'Did they enter in a new password
     If Len(sNewPassword) < 1 Then
          Worksheets("Welcome").Range(gc_wsWelcomeTempPasswordError).Value = "Please enter in a new password."
          Exit Sub
     End If
     
     'Did they retype the new password
     If Len(sRetypeNewPassword) < 1 Then
          Worksheets("Welcome").Range(gc_wsWelcomeTempPasswordError).Value = "Please retype the new password."
          Exit Sub
     End If
     
     'Does the new password and retyped new password match
     If sNewPassword <> sRetypeNewPassword Then
          Worksheets("Welcome").Range(gc_wsWelcomeTempPasswordError).Value = "New password and retyped new password do not match."
          Exit Sub
     End If
     
     'Is the new password at least 5 characters?
     If Len(sNewPassword) < 5 Then
          Worksheets("Welcome").Range(gc_wsWelcomeTempPasswordError).Value = "Your new password must be at least 5 characters in length."
          Exit Sub
     End If
     
     'At this point, the user has entered in their temporary password (we need to check if this is valid)
     'as well as a matching new & retyped new password
     
     'Before we udpate the password, let's make sure they entered their valid temp password
     'If we clear our the original password they entered when they try to login when we send them here, this will not work
     If sTempPassword <> Worksheets("Welcome").Range(gc_wsWelcomePassword).Value Then
          Worksheets("Welcome").Range(gc_wsWelcomeTempPasswordError).Value = "The existing temporary password you entered is invalid."
          Exit Sub
     Else
          'User has entered their correct temporary password and a valid new password (at least 5 characters)
          'Update the database with this new password
          'Login the user
          msgResponse = DBUser_ChangePassword(sUsername, sNewPassword)
          
          'Based on the response, let's take action
          If msgResponse = Messages.msgTrue Then
               'Successful! So let's login the user
          Else
               'Failed! Let's show the appropriate error message
          End If
          
          
          'TODO TODO TODO
          
     End If

End Sub