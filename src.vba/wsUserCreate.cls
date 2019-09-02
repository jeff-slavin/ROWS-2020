Option Explicit

Private Sub cmdUserCreate_Cancel_Click()

     'Send back to the dashboard
     
     Call Worksheets_ShowWelcomeAndOne("Dashboard")

End Sub

Private Sub Worksheet_Activate()
'This fires every time this worksheet is selected to view
'Clear out any old data

     'Variable Declarations
     Dim msgResponse As Messages
     Dim asLesserRoles() As String
     Dim sLesserRoles As String
     Dim i As Integer
     
     If IsArray(asLesserRoles) Then Erase asLesserRoles
     
     'Make sure we have a valid user
     If Not g_cUser Is Nothing Then
     
          'do we have a username, and does this user have this permission?
          If g_cUser.Username = "" Or InStr(g_cUser.Permissions, "User_Create") = 0 Then
               Call g_cUser.Logout
               Exit Sub
          End If
     
     Else
          'Global instance not set, so create it and logout
          Set g_cUser = New clsUser
          Call g_cUser.Logout
          Exit Sub
          
     End If
     
     'Clear out the form
     
     
     'Clear username & temp password cells
     Me.Range(gc_wsUserCreateUsername).Value = ""
     Me.Range(gc_wsUserCreateTempPassword).Value = ""
     
     'Set the roles drop-down
     'Can only create roles with a rank less than the current logged in user
     'TODO: Building a clsUser function to return roles that are ranked less
     'TODO: Then also need to check that the current user even has any of these permissions
     'We only want users to be able to create roles that are less than their role
          'and only add permissions that they have (even if it is a role default permission)
          'if they don't have it, they can't grant it
     msgResponse = DBUser_GetLesserRoles(g_cUser.Username, asLesserRoles)
     
     'Check for errors
     If msgResponse <> Messages.msgTrue Then Exit Sub  'error occurred
     
     'We have a valid response
     sLesserRoles = ""
     For i = LBound(asLesserRoles) To UBound(asLesserRoles)
          sLesserRoles = sLesserRoles & asLesserRoles(i) & ","
     Next i
     
     With Range(gc_wsUserCreateRole).Validation
          .Delete
          .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sLesserRoles
          .IgnoreBlank = True
          .InCellDropdown = True
          .InputTitle = ""
          .ErrorTitle = ""
          .InputMessage = ""
          .ErrorMessage = ""
          .ShowInput = True
          .ShowError = True
     End With
     
     'Clear memeory
     If IsArray(asLesserRoles) Then Erase asLesserRoles

End Sub