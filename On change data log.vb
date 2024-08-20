Private Sub ND6_Change()
'CW-1 DETAILS

    Dim objConnection As ADODB.Connection
    Dim objRecordset1 As ADODB.Recordset
    Dim StrSQL1 As String
    Set objConnection = New ADODB.Connection
    Set objRecordset1 = New ADODB.Recordset
    Dim Date1 As String
    
    objConnection.ConnectionString = "PROVIDER = SQLOLEDB;database=ITC_Checkweigher;server=localhost;uid=sa;pwd=Admin@1234;Connect Timeout=5;"

    objConnection.Open
    If Err.Description <> "" Then
        MsgBox "DATABASE CONNECTION ERROR" 'Err.Description
    End If

    Date1 = Format(Now(), "yyyy-mm-dd hh:mm:ss")

   '======================================================================
   
    StrSQL1 = "Insert into CW_DATA_STD Values ('" & Date1 & "', 1, '" & SD1.Value & "', " & ND1.Value & ", " & ND2.Value & ", " & ND3.Value & ", " & ND4.Value & ", '" & SD2.Value & "', '" & SD3.Value & "', " & ND5.Value & ")"
    
    
    objRecordset1.Open StrSQL1, objConnection, adOpenKeyset

    '======================================================================
   
    'objRecordset1.Close

    objConnection.Close

    Set objConnection = Nothing
    Set objRecordset1 = Nothing

End Sub

Private Sub ND12_Change()
'CW-2 DETAILS

    Dim objConnection As ADODB.Connection
    Dim objRecordset1 As ADODB.Recordset
    Dim StrSQL1 As String
    Set objConnection = New ADODB.Connection
    Set objRecordset1 = New ADODB.Recordset
    Dim Date1 As String
    
    objConnection.ConnectionString = "PROVIDER = SQLOLEDB;database=ITC_Checkweigher;server=localhost;uid=sa;pwd=Admin@1234;Connect Timeout=5;"

    objConnection.Open
    If Err.Description <> "" Then
        MsgBox "DATABASE CONNECTION ERROR" 'Err.Description
    End If

    Date1 = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    '======================================================================
   
    StrSQL1 = "Insert into CW_DATA_STD Values ('" & Date1 & "', 2, '" & SD4.Value & "', " & ND7.Value & ", " & ND8.Value & ", " & ND9.Value & ", " & ND10.Value & ", '" & SD5.Value & "', '" & SD6.Value & "', " & ND11.Value & ")"
    
    objRecordset1.Open StrSQL1, objConnection, adOpenKeyset

    '======================================================================
   
    'objRecordset1.Close

    objConnection.Close

    Set objConnection = Nothing
    Set objRecordset1 = Nothing

End Sub

Private Sub ND18_Change()
'CW-3 DETAILS

    Dim objConnection As ADODB.Connection
    Dim objRecordset1 As ADODB.Recordset
    Dim StrSQL1 As String
    Set objConnection = New ADODB.Connection
    Set objRecordset1 = New ADODB.Recordset
    Dim Date1 As String
    
    objConnection.ConnectionString = "PROVIDER = SQLOLEDB;database=ITC_Checkweigher;server=localhost;uid=sa;pwd=Admin@1234;Connect Timeout=5;"

    objConnection.Open
    If Err.Description <> "" Then
        MsgBox "DATABASE CONNECTION ERROR" 'Err.Description
    End If

    Date1 = Format(Now(), "yyyy-mm-dd hh:mm:ss")

    '======================================================================
   
    StrSQL1 = "Insert into CW_DATA_STD Values ('" & Date1 & "', 3, '" & SD7.Value & "', " & ND13.Value & ", " & ND14.Value & ", " & ND15.Value & ", " & ND16.Value & ", '" & SD8.Value & "', '" & SD9.Value & "', " & ND17.Value & ")"
       
    objRecordset1.Open StrSQL1, objConnection, adOpenKeyset

    '======================================================================
   
    'objRecordset1.Close

    objConnection.Close

    Set objConnection = Nothing
    Set objRecordset1 = Nothing

End Sub

Private Sub ND24_Change()
'CW-4 DETAILS

    Dim objConnection As ADODB.Connection
    Dim objRecordset1 As ADODB.Recordset
    Dim StrSQL1 As String
    Set objConnection = New ADODB.Connection
    Set objRecordset1 = New ADODB.Recordset
    Dim Date1 As String
    
    objConnection.ConnectionString = "PROVIDER = SQLOLEDB;database=ITC_Checkweigher;server=localhost;uid=sa;pwd=Admin@1234;Connect Timeout=5;"

    objConnection.Open
    If Err.Description <> "" Then
        MsgBox "DATABASE CONNECTION ERROR" 'Err.Description
    End If

    Date1 = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    '======================================================================
   
    StrSQL1 = "Insert into CW_DATA_STD Values ('" & Date1 & "', 4, '" & SD10.Value & "', " & ND19.Value & ", " & ND20.Value & ", " & ND21.Value & ", " & ND22.Value & ", '" & SD11.Value & "', '" & SD12.Value & "', " & ND23.Value & ")"
    
    objRecordset1.Open StrSQL1, objConnection, adOpenKeyset

    '======================================================================
   
    'objRecordset1.Close

    objConnection.Close

    Set objConnection = Nothing
    Set objRecordset1 = Nothing

End Sub

Private Sub ND30_Change()
'CW-5 DETAILS

    Dim objConnection As ADODB.Connection
    Dim objRecordset1 As ADODB.Recordset
    Dim StrSQL1 As String
    Set objConnection = New ADODB.Connection
    Set objRecordset1 = New ADODB.Recordset
    Dim Date1 As String
    
    objConnection.ConnectionString = "PROVIDER = SQLOLEDB;database=ITC_Checkweigher;server=localhost;uid=sa;pwd=Admin@1234;Connect Timeout=5;"

    objConnection.Open
    If Err.Description <> "" Then
        MsgBox "DATABASE CONNECTION ERROR" 'Err.Description
    End If

    Date1 = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    '======================================================================
   
    StrSQL1 = "Insert into CW_DATA_STD Values ('" & Date1 & "', 5, '" & SD13.Value & "', " & ND25.Value & ", " & ND26.Value & ", " & ND27.Value & ", " & ND28.Value & ", '" & SD14.Value & "', '" & SD15.Value & "', " & ND29.Value & ")"
    
    objRecordset1.Open StrSQL1, objConnection, adOpenKeyset

    '======================================================================
   
    'objRecordset1.Close

    objConnection.Close

    Set objConnection = Nothing
    Set objRecordset1 = Nothing

End Sub

Private Sub ND36_Change()
'CW-6 DETAILS

    Dim objConnection As ADODB.Connection
    Dim objRecordset1 As ADODB.Recordset
    Dim StrSQL1 As String
    Set objConnection = New ADODB.Connection
    Set objRecordset1 = New ADODB.Recordset
    Dim Date1 As String
    
    objConnection.ConnectionString = "PROVIDER = SQLOLEDB;database=ITC_Checkweigher;server=localhost;uid=sa;pwd=Admin@1234;Connect Timeout=5;"

    objConnection.Open
    If Err.Description <> "" Then
        MsgBox "DATABASE CONNECTION ERROR" 'Err.Description
    End If

    Date1 = Format(Now(), "yyyy-mm-dd hh:mm:ss")

    '======================================================================
   
    StrSQL1 = "Insert into CW_DATA_STD Values ('" & Date1 & "', 6, '" & SD16.Value & "', " & ND31.Value & ", " & ND32.Value & ", " & ND33.Value & ", " & ND34.Value & ", '" & SD17.Value & "', '" & SD18.Value & "', " & ND35.Value & ")"
    
    objRecordset1.Open StrSQL1, objConnection, adOpenKeyset

    '======================================================================
   
    'objRecordset1.Close

    objConnection.Close

    Set objConnection = Nothing
    Set objRecordset1 = Nothing

End Sub

Private Sub ND42_Change()
'CW-7 DETAILS

    Dim objConnection As ADODB.Connection
    Dim objRecordset1 As ADODB.Recordset
    Dim StrSQL1 As String
    Set objConnection = New ADODB.Connection
    Set objRecordset1 = New ADODB.Recordset
    Dim Date1 As String
    
    objConnection.ConnectionString = "PROVIDER = SQLOLEDB;database=ITC_Checkweigher;server=localhost;uid=sa;pwd=Admin@1234;Connect Timeout=5;"

    objConnection.Open
    If Err.Description <> "" Then
        MsgBox "DATABASE CONNECTION ERROR" 'Err.Description
    End If

    Date1 = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    '======================================================================
   
    StrSQL1 = "Insert into CW_DATA_STD Values ('" & Date1 & "', 7, '" & SD19.Value & "', " & ND37.Value & ", " & ND38.Value & ", " & ND39.Value & ", " & ND40.Value & ", '" & SD20.Value & "', '" & SD21.Value & "', " & ND41.Value & ")"
    
    objRecordset1.Open StrSQL1, objConnection, adOpenKeyset

    '======================================================================
   
    'objRecordset1.Close

    objConnection.Close

    Set objConnection = Nothing
    Set objRecordset1 = Nothing

End Sub

Private Sub ND48_Change()
'CW-8 DETAILS

    Dim objConnection As ADODB.Connection
    Dim objRecordset1 As ADODB.Recordset
    Dim StrSQL1 As String
    Set objConnection = New ADODB.Connection
    Set objRecordset1 = New ADODB.Recordset
    Dim Date1 As String
    
    objConnection.ConnectionString = "PROVIDER = SQLOLEDB;database=ITC_Checkweigher;server=localhost;uid=sa;pwd=Admin@1234;Connect Timeout=5;"

    objConnection.Open
    If Err.Description <> "" Then
        MsgBox "DATABASE CONNECTION ERROR" 'Err.Description
    End If

    Date1 = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    '======================================================================
   
    StrSQL1 = "Insert into CW_DATA_STD Values ('" & Date1 & "', 8, '" & SD22.Value & "', " & ND43.Value & ", " & ND44.Value & ", " & ND45.Value & ", " & ND46.Value & ", '" & SD23.Value & "', '" & SD24.Value & "', " & ND47.Value & ")"
   
    objRecordset1.Open StrSQL1, objConnection, adOpenKeyset

    '======================================================================
   
    'objRecordset1.Close

    objConnection.Close

    Set objConnection = Nothing
    Set objRecordset1 = Nothing

End Sub
