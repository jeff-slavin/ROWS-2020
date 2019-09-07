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

     'Did the user select a new role (not blank)
     If Target.Address = gc_wsUserCreateRole Then
          If Target.Value <> "" Then
               'Fill out the permission table based on the selected role in this dropdown cell
               'TODO
          End If
     End If
     

End Sub