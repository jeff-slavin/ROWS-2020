Option Explicit

Public Function wsCreateUser_FillOutPermissionsTable(ByVal sRole As String, ByRef sError As String) As Messages
'Get a list of all possible permissions that are active
'List all of these permissions out
'For "Enabled" column, do the following
'    Pull in the default role permissions
'    If the current user also has these permissions, then set them to enabled in the permissions table
'    Permissions that do not meet these two criteria, grey out? and do not let the user set these
          'lock the cell?


     'Variable Declarations
     

     'Set initial response states
     wsCreateUser_FillOutPermissionsTable = Messages.msgFalse
     sError = ""
     
     

End Function

Public Sub wsUserCreate_ClearSheet()
'Clear the worksheet cells that require inputs:
'    Username
'    Temp Password
'    Role
'    Permissions table

     'Variable Declarations
     Dim ws As Worksheet
     
     Set ws = Worksheets("UserCreate")
     
     'Clear out the global error cell
     ws.Range(gc_rErrorCell).Value = ""
     
     'Clear username and password
     ws.Range(gc_wsUserCreateUsername).Value = ""
     ws.Range(gc_wsUserCreateTempPassword).Value = ""
     
     'Clear Role drop-down
     ws.Range(gc_wsUserCreateRole).Validation.Delete
     ws.Range(gc_wsUserCreateRole).Value = ""
     
     'Clear Permissions Table
     Call wsUserCreate_ClearPermissionTable
     
     'Clear memory
     Set ws = Nothing

End Sub

Public Sub wsUserCreate_ClearPermissionTable()
'Clear out the permissions table range

     Worksheets("UserCreate").Range(gc_wsUserCreatePermissionTableRange).ClearContents

End Sub

Public Function wsUserCreate_SetRoleDropdown() As Messages

     'Variable Declarations
     Dim asLesserRoles() As String
     Dim sLesserRoles As String
     Dim i As Integer
     
     wsUserCreate_SetRoleDropdown = DBUser_GetLesserRoles(g_cUser.Username, asLesserRoles)
     
     'Check for errors
     If wsUserCreate_SetRoleDropdown <> Messages.msgTrue Then GoTo wsUserCreate_SetRoleDropdown_Error  'error occurred
     
     'We have a valid response
     sLesserRoles = ""
     For i = LBound(asLesserRoles) To UBound(asLesserRoles)
          sLesserRoles = sLesserRoles & asLesserRoles(i) & ","
     Next i

     'Setup the drop-down
     With Worksheets("UserCreate").Range(gc_wsUserCreateRole).Validation
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
     
     'Make sure the dropdown is blank to start
     Worksheets("UserCreate").Range(gc_wsUserCreateRole).Value = ""
     
wsUserCreate_SetRoleDropdown_Error:
     'Clear memory
     If IsArray(asLesserRoles) Then Erase asLesserRoles

End Function

Public Function wsCreateUser_Activate(ByRef sError As String) As Messages
     
     'Set initial return state
     wsCreateUser_Activate = msgFalse
     sError = ""
     
     'Make sure we have a valid user
     If Not g_cUser Is Nothing Then
     
          'do we have a username, and does this user have this permission?
          If g_cUser.Username = "" Or InStr(g_cUser.Permissions, "User_Create") = 0 Then
               Call g_cUser.Logout
               Exit Function
          End If
     
     Else
     
          'Global instance not set, so create it and logout
          Set g_cUser = New clsUser
          Call g_cUser.Logout
          Exit Function
          
     End If
     
     'Clear out the form
     Call wsUserCreate_ClearSheet
     
     'Fill in the lesser user roles drop-down
     wsCreateUser_Activate = wsUserCreate_SetRoleDropdown
     
     'If error, then exit the function and do not continue
     If wsCreateUser_Activate <> Messages.msgTrue Then
          sError = "Error : Unable to fill out the role dropdown in the UserCreate form :: Messages.ErrorCode = " & Str(wsCreateUser_Activate)
          Exit Function
     End If
     
End Function