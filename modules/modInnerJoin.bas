Attribute VB_Name = "modInnerJoin"
' #######################################
' # Programmer : Jose Mari Caballa Rey  #
' # Examination Number : 1234567890     #
' # Examinaiton Date : 27 November 2015 #
' #######################################

'next userid function
Public Function Next_UserID()
    'set recordset
    Set myRes = New ADODB.Recordset
    'open query
    SQL = "SELECT"
    myRes.Open SQL, myCon, 3, 2
    
    
End Function

'load record
Public Function Load_Record()
    'set recordset
    Set myRes = New ADODB.Recordset
    'open query
    'SQL = "SELECT * FROM Table1 INNER JOIN Table2 ON Table1.userid = Table2.userid ORDER BY Table1.userid = Table2.userid ASC"
    SQL = "SELECT * FROM Table1 INNER JOIN Table2 ON Table1.userid = Table2.userid"
    myRes.Open SQL, myCon, 3, 2
End Function

'close record
Public Function Close_Record()
    'close query
    myRes.Close
    'set nothing
    Set myRes = Nothing
End Function


