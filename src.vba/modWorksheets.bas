Option Explicit

'*************************************************************************************
'Worksheets :            WELCOME
'*************************************************************************************

Public Sub Worksheets_ShowAll()
'Testing Function - should not be called when in production

     Dim ws As Worksheet
     
     For Each ws In ActiveWorkbook.Worksheets
          ws.Visible = xlSheetVisible
     Next ws
     
     Set ws = Nothing

End Sub

Public Sub Worksheets_ShowWelcomeAndOneWithError(ByVal sSheetName As String, ByVal sError As String)

     Dim ws As Worksheet
     
     For Each ws In ActiveWorkbook.Worksheets
          
          If ws.Name = sSheetName Or ws.Name = "Welcome" Then
               'If Welcome or the requested worksheet, show it
               ws.Visible = xlSheetVisible
               
               'Show the error
               If ws.Name = sSheetName Then ws.Range(gc_rErrorCell).Value = sError
               
               'Only activate the "Welcome" worksheet if it was the sheet asked for
               If ws.Name <> "Welcome" Or sSheetName = "Welcome" Then ws.Activate
          Else
               'Hide all other worksheets
               ws.Visible = xlSheetVeryHidden
          End If
     
     Next ws
     
     Set ws = Nothing

End Sub

Public Sub Worksheets_ShowWelcomeAndOne(ByVal sSheetName As String)
'Keeps the Welcome worksheet visible
'Hides all other worksheets except for the one that has the name matching the parameter sSheetName
'If you call this with "Welcome" then it will only show the "Welcome" worksheet and activate it
'If you call this with anything other than "Welcome", it will show that other sheet and activate it (while keeping "Welcome" visible)

     Dim ws As Worksheet
     
     For Each ws In ActiveWorkbook.Worksheets
          
          If ws.Name = sSheetName Or ws.Name = "Welcome" Then
               'If Welcome or the requested worksheet, show it
               ws.Visible = xlSheetVisible
               
               'clear error cell
               If ws.Name = sSheetName Then ws.Range(gc_rErrorCell).Value = ""
               
               'Only activate the "Welcome" worksheet if it was the sheet asked for
               If ws.Name <> "Welcome" Or sSheetName = "Welcome" Then ws.Activate
          Else
               'Hide all other worksheets
               ws.Visible = xlSheetVeryHidden
          End If
     
     Next ws
     
     Set ws = Nothing

End Sub

Public Sub wsWelcome_ShowSection(ByVal sSectionName As String)
'Shows the section called out by parameter sSectionName on the 'Welcome' worksheet
'Uses global variables to tell what ranges to show/hide

     Dim ws As Worksheet
     
     Set ws = Worksheets("Welcome")
     
     ws.Activate

     Select Case sSectionName
     
          Case "Login"
               ws.Range(gc_wsLoginRange).EntireRow.Hidden = False
               ws.Shapes("cmdLogin").Top = gc_wscmdLoginTop
               ws.Shapes("cmdLogin").Visible = True
               ws.Range("D7").Select
               ws.Range(gc_wsWelcomeUsername).Value = ""
               ws.Range(gc_wsWelcomePassword).Value = ""
               ws.Range(gc_wsWelcomeLoginError).Value = ""
               
               ws.Range(gc_wsLogoutRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogout").Visible = False
               
               ws.Range(gc_wsTempPasswordRange).EntireRow.Hidden = True
               ws.Shapes("cmdUpdatePassword").Visible = False
               
          Case "Logout"
               ws.Range(gc_wsLoginRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogin").Visible = False
               
               ws.Range(gc_wsLogoutRange).EntireRow.Hidden = False
               ws.Shapes("cmdLogout").Top = gc_wscmdLogoutTop
               ws.Shapes("cmdLogout").Visible = True
               ws.Range("D17").Select
               
               If Not g_cUser Is Nothing Then
                    ws.Range(gc_wsWelcomeLoggedInUsername).Value = g_cUser.Username
                    ws.Range(gc_wsWelcomeLoggedInRole).Value = g_cUser.Role
               Else
                    Set g_cUser = New clsUser
                    Call wsWelcome_ShowSection("Login")
                    Call Worksheets_ShowWelcomeAndOne("Welcome")
                    Set ws = Nothing
                    Exit Sub
               End If
               
               ws.Range(gc_wsTempPasswordRange).EntireRow.Hidden = True
               ws.Shapes("cmdUpdatePassword").Visible = False
               
          Case "TempPassword"
               ws.Range(gc_wsLoginRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogin").Visible = False
               
               ws.Range(gc_wsLogoutRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogout").Visible = False
               
               ws.Range(gc_wsWelcomeTempPassword).Value = ""
               ws.Range(gc_wsWelcomeNewPassword).Value = ""
               ws.Range(gc_wsWelcomeRetypePassword).Value = ""
               ws.Range(gc_wsWelcomeTempPasswordError).Value = ""
               
               ws.Range(gc_wsTempPasswordRange).EntireRow.Hidden = False
               ws.Shapes("cmdUpdatePassword").Top = gc_wscmdTempPasswordTop
               ws.Shapes("cmdUpdatePassword").Visible = True
               ws.Range("E26").Select
     
     End Select
     
     Set ws = Nothing

End Sub