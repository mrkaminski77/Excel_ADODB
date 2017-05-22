Attribute VB_Name = "ADODB_Module"
Option Explicit
'
'   ADODB ROUTINES
'
'
' When creating a new connection the string is all that is required
Private Const ADO_Connection As String = "Driver={SQL Server};Server=SCSBENSQLDEV01; Database=ATOReporting_UAT; Trusted_Connection=yes;"


Public Function SQL_GetScalar(commandText As String) As String
'   This function exeutes command text and expects a single scalar result which is returned as a String
    Dim ADODB_Recordset As Object
    Set ADODB_Recordset = CreateObject("ADODB.RecordSet")
    ADODB_Recordset.Open commandText, ADO_Connection
    SQL_GetScalar = Replace(ADODB_Recordset.GetString(), vbCr, vbNullString)
    ADODB_Recordset.ActiveConnection.Close
    Set ADODB_Recordset = Nothing
End Function


Public Sub SQL_Command(commandText As String)
'   This sub executes a command. Used for insert procedures.
    Dim ADODB_Command As Object
    Set ADODB_Command = CreateObject("ADODB.Command")
    ADODB_Command.ActiveConnection = ADO_Connection
    ADODB_Command.commandText = commandText
    ADODB_Command.Execute
    ADODB_Command.ActiveConnection.Close
    Set ADODB_Command = Nothing
    Exit Sub
End Sub

Public Function SQL_GetArray(commandText As String) As Variant()
'   Use this function to execute procedures you expect a table of results for. Returns an array.
    Dim ADODB_Recordset As Object
    Set ADODB_Recordset = CreateObject("ADODB.RecordSet")
    ADODB_Recordset.Open commandText, ADO_Connection
    SQL_GetArray = Application.WorksheetFunction.Transpose(ADODB_Recordset.GetRows())
    ADODB_Recordset.ActiveConnection.Close
    Set ADODB_Recordset = Nothing
End Function

Public Function CreateCommandText(procName As String, Optional paramString As String="") As String
'   convert the procedure name and paramString in to commandtext
    CreateCommandText = "EXEC " & procName & " " & paramString & ";"
End Function


Public Function GetSQLUserName() As String
    GetSQLUserName = SQL_GetScalar("SELECT CURRENT_USER;")
End Function


Public Function SQL_Type(param As Variant) As String
'   Convert a variable into sql type
        Select Case VarType(param)
            Case 2, 3, 4, 5, 14
                SQL_Type = CStr(param)
            Case 7
                SQL_Type = SQL_Date(CDate(param))
            Case 8
                SQL_Type = SQL_StringLiteral(CStr(param))
        End Select
End Function

Private Function SQL_Date(d As Date) As String
'   Convert date into ODBC format
    SQL_Date = "{d'" & Format(d, "YYYY-MM-DD") & "'}"
End Function

Private Function SQL_DateTime(d As Date) As String
'   Convert DateTime in to ODBC format
    SQL_DateTime = "{dt'" & Format(d, "YYYY-MM-DD hh:mm:ss") & "'}"
End Function

Private Function SQL_StringLiteral(text As String) As String
'   Convert string in to SQL string
    SQL_StringLiteral = "'" & Replace(text, "'", "''") & "'"
End Function

