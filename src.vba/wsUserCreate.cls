Option Explicit

Private Sub cmdUserCreate_Cancel_Click()

     'Send back to the dashboard
     
     Call Worksheets_ShowWelcomeAndOne("Dashboard")

End Sub

Private Sub Worksheet_Activate()
'This fires every time this worksheet is selected to view

     Dim msgResponse As Messages
     Dim sError As String
     
     msgResponse = wsCreateUser_Activate(sError)
     
     'Check for error
     If msgResponse <> Messages.msgTrue Then
          Call Worksheets_ShowWelcomeAndOneWithError("Dashboard", sError)
     End If
          

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

     Dim msgResponse As Messages
     Dim sError As String

     MsgBox Target.Address

     'Did the user select a new role (not blank)
     If Target.Address = gc_wsUserCreateRole Or Target.Address = gc_wsUserCreateRoleRange Then
          
          'Clear out permissions no matter what, a change happened
          Call wsUserCreate_ClearPermissionTable
          
          'Is the Target Cell not empty? If so, let's do the permissions table
          'Reason for calling Cells(1,1) is so that this avoids the error when Target is a range
          'In this case our role cell is 2 cells merged, so we get a range here
          If Target.Cells(1, 1).Value <> "" Then
               'Fill out the permission table based on the selected role in this dropdown cell
               msgResponse = wsCreateUser_FillOutPermissionsTable(Target.Value, sError)
          End If
          
     End If
     

End Sub