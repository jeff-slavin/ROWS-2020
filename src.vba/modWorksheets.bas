Option Explicit

'*************************************************************************************
'Worksheets :            WELCOME
'*************************************************************************************

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

     Select Case sSectionName
     
          Case "Login"
               ws.Range(gc_wsLoginRange).EntireRow.Hidden = False
               ws.Shapes("cmdLogin").Top = gc_wscmdLoginTop
               ws.Shapes("cmdLogin").Visible = True
               ws.Range("D7").Select
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
               ws.Range(gc_wsTempPasswordRange).EntireRow.Hidden = True
               ws.Shapes("cmdUpdatePassword").Visible = False
               
          Case "TempPassword"
               ws.Range(gc_wsLoginRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogin").Visible = False
               ws.Range(gc_wsLogoutRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogout").Visible = False
               ws.Range(gc_wsTempPasswordRange).EntireRow.Hidden = False
               ws.Shapes("cmdUpdatePassword").Top = gc_wscmdTempPasswordTop
               ws.Shapes("cmdUpdatePassword").Visible = True
               ws.Range("E26").Select
     
     End Select
     
     Set ws = Nothing

End Sub