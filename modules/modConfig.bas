Attribute VB_Name = "modConfig"
' #######################################
' # Programmer : Jose Mari Caballa Rey  #
' # Examination Number : 1234567890     #
' # Examination Date : 27 November 2015 #
' #######################################

Public myRes As ADODB.Recordset
Public myCon As ADODB.Connection
Public SQL As String
Public dbName As String

'open connection string
Public Function Open_Connection()
    'set connection
    Set myCon = New ADODB.Connection
    'database name
    dbName = "innerjoin.mdb"
    'open connection
    myCon.ConnectionString = "provider=microsoft.jet.oledb.4.0; " & _
    "data source=" & App.Path & "\" & dbName & ";"
    myCon.Open
End Function

'close connection string
Public Function Close_Conneciton()
    'close connection
    myCon.Close
    'set nothing
    Set myCon = Nothing
End Function
