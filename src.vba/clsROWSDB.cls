Option Explicit

'***************************************************************************************************
'Class Name:   clsROWSDB
'---------------------------------------------------------------------------------------------------
'Purpose:      This class wraps the ADODB Connection and Recordset objects.
'              This handles all database connections to the ROWS Access database.
'              The private member variables are never exposed outside of this class.
'              Handles setting up and closing down connections to the database.
'              Handles running queries to the database
'              Gives wrapped access to the ADOB Recordset object via passthrough functions
'***************************************************************************************************

'Private Member Variables
Private m_adoConnection As ADODB.Connection
Private m_adoRecordset As ADODB.Recordset

Private Sub Class_Initialize()
'Constructor
'Initializes the private member variables (but does not setup any DB connection)

     Set m_adoConnection = New ADODB.Connection
     Set m_adoRecordset = New ADODB.Recordset

End Sub

Private Sub Class_Terminate()
'Destructor
'Closes any connection and frees up memory used by the private member variables

     Call Me.Done
          
End Sub

Public Function Connect() As Messages
'Sets up a connection with the database
'Returns True/False based on if the connection was successful

     Dim sFilePath As String
     
     'Set initial return state as False
     Connect = Messages.msgFailedDBConnection
     
     'Close any previous connections
     Call Me.Done
    
     'Build path to the database
     sFilePath = ActiveWorkbook.Path & gc_sROWSDBFilename

     'Check if file exists, exit function if it does not
     If Dir(sFilePath) = "" Then Exit Function

     'Setup new connection
     Set m_adoConnection = New ADODB.Connection
     m_adoConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sFilePath

     'Check if connection was successful
     If m_adoConnection.State <> adStateOpen Then
          'Not successful to close any connections and exit function with False
          Call Me.Done
          Exit Function
     End If
     
     'Connection Successful
     Connect = Messages.msgTrue
     Exit Function

End Function

Public Function Query(ByVal sSQL As String, Optional ByVal bReadOnly As Boolean = True) As Messages
'runs the query, saves the results in m_adoRecordset member variable
'returns True/False based on if the query was successful

     'Set initial return state as false
     Query = Messages.msgFailedDBConnection
     
     'Close any previous recordsets & connections
     Call Me.Done
     
     'Open the connection
     If Me.Connect = Messages.msgFailedDBConnection Then Exit Function
     
     Query = Messages.msgFailedQuery
     
     Set m_adoRecordset = New ADODB.Recordset

     'Try / Catch around the running of the SQL Query
     On Error GoTo Query_Error
     
          'If bReadOnly parameter is passed as true, then open the database for adLockReadOnly
          If bReadOnly = True Then
               m_adoRecordset.Open sSQL, m_adoConnection, adOpenKeyset, adLockReadOnly
          Else
               m_adoRecordset.Open sSQL, m_adoConnection, adOpenKeyset, adLockOptimistic
          End If
     
     On Error GoTo 0
     'End of Try / Catch

     Query = Messages.msgTrue
     
     Exit Function
     
Query_Error:
     Call Me.Done
     Exit Function

End Function

Public Sub Done()
' closes the connection & recordset
' frees up any allocated memory for the connection and recordset
' also called by Class_Terminate

     If Not m_adoRecordset Is Nothing Then
          If m_adoRecordset.State <> adStateClosed Then m_adoRecordset.Close
          Set m_adoRecordset = Nothing
     End If
     
     If Not m_adoConnection Is Nothing Then
          If m_adoConnection.State <> adStateClosed Then m_adoConnection.Close
          Set m_adoConnection = Nothing
     End If
     
End Sub

Public Function RecordCount() As Integer
'passthrough function

     If Not m_adoRecordset Is Nothing Then RecordCount = m_adoRecordset.RecordCount

End Function

Public Function BOF() As Boolean
'passthrough function

     If Not m_adoRecordset Is Nothing Then BOF = m_adoRecordset.BOF

End Function

Public Function EOF() As Boolean
'passthrough function

     If Not m_adoRecordset Is Nothing Then EOF = m_adoRecordset.EOF
     
End Function

Public Sub MoveFirst()
'passthrough procedure

     If Not m_adoRecordset Is Nothing Then m_adoRecordset.MoveFirst
     
End Sub

Public Sub MoveLast()
'passthrough procedure

     If Not m_adoRecordset Is Nothing Then m_adoRecordset.MoveLast
     
End Sub

Public Sub MoveNext()
'passthrough procedure

     If Not m_adoRecordset Is Nothing Then m_adoRecordset.MoveNext
     
End Sub

Public Sub MovePrevious()
'passthrough procedure

     If Not m_adoRecordset Is Nothing Then m_adoRecordset.MovePrevious
     
End Sub

Public Function FieldsCount() As Integer
'passthrough procedure

     If Not m_adoRecordset Is Nothing Then FieldsCount = m_adoRecordset.Fields.Count
     
End Function

Public Function Fields(ByVal sFieldName As String, ByRef vValue As Variant) As Boolean
'passthrough procedure
'returns true/false based on if the requested sFieldName was found

     Fields = False

     'Try / Catch Start
     On Error GoTo Fields_Error

          If Not m_adoRecordset Is Nothing Then Fields = m_adoRecordset.Fields(sFieldName).Value
     
     On Error GoTo 0
     'End Try/Catch
     
     Fields = True
     Exit Function
     
Fields_Error:
     Exit Function
     
End Function
