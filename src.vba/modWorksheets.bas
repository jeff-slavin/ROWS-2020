Option Explicit

'*************************************************************************************
'Worksheets :            WELCOME
'*************************************************************************************

Public Sub wsWelcome_ShowSection(ByVal sSectionName As String)
'Shows the section called out by parameter sSectionName on the 'Welcome' worksheet
'Uses global variables to tell what ranges to show/hide

     Dim ws As Worksheet
     
     Set ws = Worksheets("Welcome")

     Select Case sSectionName
     
          Case "Login"
               ws.Range(gc_wsLoginRange).EntireRow.Hidden = False
               ws.Shapes("cmdLogin").Visible = True
               ws.Range(gc_wsLogoutRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogout").Visible = False
               ws.Range(gc_wsTempPasswordRange).EntireRow.Hidden = True
               
          Case "Logout"
               ws.Range(gc_wsLoginRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogin").Visible = False
               ws.Range(gc_wsLogoutRange).EntireRow.Hidden = False
               ws.Shapes("cmdLogout").Visible = True
               ws.Range(gc_wsTempPasswordRange).EntireRow.Hidden = True
               
          Case "TempPassword"
               ws.Range(gc_wsLoginRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogin").Visible = False
               ws.Range(gc_wsLogoutRange).EntireRow.Hidden = True
               ws.Shapes("cmdLogout").Visible = False
               ws.Range(gc_wsTempPasswordRange).EntireRow.Hidden = False
     
     End Select
     
     Set ws = Nothing

End Sub