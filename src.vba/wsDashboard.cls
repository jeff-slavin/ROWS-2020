Option Explicit

Private Sub Worksheet_Activate()
'This is called whenever the Dashboard worksheet is selected
'Let's enable/disable the buttons based on the current permissions

     'Variable Declarations
     Dim sPermissions As String

     'First check if there is a valid user logged in
     'If not, let's call the logout function to kick the user back to the Welcome/Login worksheet
     If Not g_cUser Is Nothing Then
          If g_cUser.Username = "" Then
               Call g_cUser.Logout
               Exit Sub
          End If
     Else
          'No valid user global instance
          'Set this instance and call logout
          Set g_cUser = New clsUser
          Call g_cUser.Logout
          Exit Sub
     End If
     
     'If we got this far, we have a valid user saved in the global clsUser instance
     'Now let's set the buttons based on permissions

     'First save the permission string from g_cUser as a local variable to avoid continuously calling a function
     sPermissions = g_cUser.Permissions

'*****************************************************
'User Management Buttons
'*****************************************************
     
     'Create User
     If InStr(sPermissions, "User_Create") > 0 Then
          Me.cmdUserCreate.Enabled = True
     Else
          Me.cmdUserCreate.Enabled = False
     End If
     
     'Deactivate User
     If InStr(sPermissions, "User_Deactivate") > 0 Then
          Me.cmdUserDeactivate.Enabled = True
     Else
          Me.cmdUserDeactivate.Enabled = False
     End If
     
     'Edit/Change User Role
     If InStr(sPermissions, "User_EditRole") > 0 Then
          Me.cmdUserChangeRole.Enabled = True
     Else
          Me.cmdUserChangeRole.Enabled = False
     End If
     
     'Edit/Change User Permissions
     If InStr(sPermissions, "User_EditPermissions") > 0 Then
          Me.cmdUserEditPermissions.Enabled = True
     Else
          Me.cmdUserEditPermissions.Enabled = False
     End If
     
     'Reset User Password
     If InStr(sPermissions, "User_ResetPassword") > 0 Then
          Me.cmdUserResetPassword.Enabled = True
     Else
          Me.cmdUserResetPassword.Enabled = False
     End If
     
     

End Sub