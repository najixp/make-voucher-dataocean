Attribute VB_Name = "Module1"
Option Explicit

Public dbxSQLacc As ADODB.Connection
Public dbxSQL As ADODB.Connection
Public CN As Connection
Public ServerDateRC As Recordset
Public ServerDate As Date
Public Function fOpenRSSQL(ByRef pRS As ADODB.Recordset, ByVal pStrSQL As String)
Set pRS = New ADODB.Recordset
pRS.Source = pStrSQL
pRS.ActiveConnection = dbxSQL
pRS.Open pStrSQL, dbxSQL, adOpenStatic, adLockOptimistic
End Function
Public Function fOpenRSSQLacc(ByRef pRS As ADODB.Recordset, ByVal pStrSQL As String)
Set pRS = New ADODB.Recordset
pRS.Source = pStrSQL
pRS.ActiveConnection = dbxSQLacc
pRS.Open pStrSQL, dbxSQLacc, adOpenStatic, adLockOptimistic

End Function


Public Sub MKConnStrSQL()
    Set dbxSQL = New ADODB.Connection
    Set dbxSQLacc = New ADODB.Connection
    'dbxSQL.ConnectionString = "Provider=SQLNCLI10;Server=asnserver1;Initial Catalog=OnLine_Stock_2019;uid=sa;password=;Trusted_Connection=yes;"
    'dbxSQL.ConnectionString = "Driver={SQL Server};Server=asnserver1;Database=OnLine_Stock_2019;Uid=sa;Pwd=;Trusted_Connection=no;"
    dbxSQL.ConnectionString = "Provider=sqloledb;Data Source=ASN_01;Initial Catalog=OnLine_Stock_2019;User Id=sa;Password=;Trusted_Connection=no;"
    dbxSQL.Open
    dbxSQLacc.ConnectionString = "Provider=sqloledb;Data Source=ASN_01;Initial Catalog=OnLine_Account_2019;User Id=sa;Password=;Trusted_Connection=no;"
    dbxSQLacc.Open

End Sub
Public Function ConnectToServer() As String


    Dim ServiceName As String
Dim aa
            
    Set CN = New Connection
    
    
    Err.Clear
    'ww = "PROVIDER=MSDASQL;driver={SQL Server};server=MAKKAHSERVER\BADAWOOD;uid=sa;pwd=123;database=badawood1;"
    'aa = "PROVIDER=MSDASQL;driver={SQL Server};server=" & ServiceName & ";uid=" & FuncLogin & ";pwd=" & FuncPass & ";database=" & DBNameInServer & ";"
    aa = "PROVIDER=MSDASQL;driver={SQL Server};server=ASN_01;uid=sa;pwd=;database=OnLine_Account_2019;"
    'aa = "PROVIDER=sqlncli;Driver={SQL Native Cient};server=" & ServiceName & ";uid=" & FuncLogin & ";pwd=" & FuncPass & ";database=" & DBNameInServer & ";MARS Connection=True;DataTypeCompatibility=80;"
    'aa = "PROVIDER=MSDASQL.1;driver={SQL Server};server=" & ServiceName & ";uid=" & FuncLogin & ";pwd=" & FuncPass & ";database=" & DBNameInServer & ";MARS Connection=True;DataTypeCompatibility=80;"
    
    On Error Resume Next
    CN.Open aa

    If Err Then
        MsgBox Err.Description
    End If
    
    If Err Then
        Err.Clear
        Set CN = New Connection
        CN.Open "PROVIDER=MSDASQL;driver={SQL Server};server=ASN_01;uid=sa;pwd=;database=OnLine_Account_2019;"
        If Err > 0 Or CN.State = adStateClosed Then
            MsgBox Err.Description
            End
        End If
    End If
    
    If CN.State <> adStateOpen Then
        MsgBox "I?C? ?? C?CE?C?", vbCritical
        End
    End If
    
    ConnectToServer = "asnserver1"

End Function



