Attribute VB_Name = "sqlModule"
'@Lang VBA

Public Function sqlSELECT(accessFile, colArray, dbTable, whereArray, compareArray, oppArray) As Variant

Dim result() As Variant
Dim setNum As Integer
setNum = 0
Sql = "SELECT "

colLen = UBound(colArray)
whereLen = UBound(whereArray)

For i = 0 To colLen
    If i = colLen Then
        Sql = Sql & colArray(i)
    Else
        Sql = Sql & colArray(i) & ", "
    End If
Next i

Sql = Sql & " FROM " & dbTable & ""

If whereLen > 0 Or whereArray(0) <> "" Then
    Sql = Sql & " WHERE "
    For i = 0 To whereLen
            Sql = Sql & whereArray(i) & " " & oppArray(i) & " "
            Select Case VarType(compareArray(i))
                Case vbString
                    If InStr(compareArray(i), "'") > 0 Then
                    Sql = Sql & """" & compareArray(i) & """"
                    ElseIf InStr(compareArray(i), Chr(34)) > 0 Then
                    compareArray(i) = Replace(compareArray(i), Chr(34), "")
                    Sql = Sql & "'" & compareArray(i) & "'"
                    Else
                    Sql = Sql & "'" & compareArray(i) & "'"
                    End If
                Case 7 ''vbDate
                Sql = Sql & "#" & compareArray(i) & "#"
                Case 2 ''vbInteger
                Sql = Sql & compareArray(i)
            End Select
            If whereLen <> i Then
                Sql = Sql & " AND "
            End If
    Next i
     
End If

Set rs = sqlRun(accessFile, Sql)

Do Until rs.EOF
    setNum = setNum + 1
    rs.MoveNext
Loop

If setNum > 0 Then
    rs.MoveFirst
End If

itemNum = setNum * (colLen + 1)
j = 0
ReDim result(itemNum)

Do Until rs.EOF
    For i = 0 To colLen
        result(j) = rs.Fields(colArray(i))
        j = j + 1
    Next i
    rs.MoveNext
Loop

rs.Close
rs.ActiveConnection = Nothing
Set rs = Nothing

sqlSELECT = result

End Function

Public Function sqlInsert(accessFile, dbTable, colArray, valArray) As Boolean

Dim result As Boolean

colLen = UBound(colArray)

Sql = "INSERT INTO " & dbTable & " ("


For i = 0 To colLen
    If i = colLen Then
        Sql = Sql & colArray(i)
    Else
        Sql = Sql & colArray(i) & ", "
    End If
Next i

Sql = Sql & ") VALUES ("

For i = 0 To colLen
    Select Case VarType(valArray(i))
        Case vbString
            Sql = Sql & "'" & valArray(i) & "'"
        Case 7 ''vbDate
            Sql = Sql & "#" & valArray(i) & "#"
        Case 2 ''vbInteger
            Sql = Sql & valArray(i)
    End Select
    If i <> colLen Then
        Sql = Sql & ", "
    End If
Next i
Sql = Sql & ")"
    
result = sqlExecute(accessFile, dbTable, Sql)


sqlInsert = result

End Function

Public Function sqlUpdate(accessFile, dbTable, colArray, updateArray, whereArray, compareArray, oppArray) As Boolean

Dim result As Boolean

colLen = UBound(colArray)
whereLen = UBound(whereArray)
Sql = "UPDATE " & dbTable & " SET "


For i = 0 To colLen
    Sql = Sql & colArray(i) & " = '" & updateArray(i) & "'"
    If i <> colLen Then
        Sql = Sql & ", "
    End If
Next i

If whereLen > 0 Or whereArray(0) <> "" Then
    Sql = Sql & " WHERE "
    For i = 0 To whereLen
            Sql = Sql & whereArray(i) & " " & oppArray(i) & " "
            Select Case VarType(compareArray(i))
                Case vbString
                Sql = Sql & "'" & compareArray(i) & "'"
                Case 7 ''vbDate
                Sql = Sql & "#" & compareArray(i) & "#"
                Case 2 ''vbInteger
                Sql = Sql & compareArray(i)
            End Select
            If whereLen <> i Then
                Sql = Sql & " AND "
            End If
    Next i
     
End If

result = sqlExecute(accessFile, dbTable, Sql)

sqlUpdate = result

End Function

Public Function sqlDelete(accessFile, dbTable, whereArray, compareArray, oppArray) As Boolean

Dim result As Boolean
result = False
whereLen = UBound(whereArray)

Sql = "DELETE FROM " & dbTable

If whereLen > 0 Or whereArray(0) <> "" Then
    Sql = Sql & " WHERE "
    For i = 0 To whereLen
            Sql = Sql & whereArray(i) & " " & oppArray(i) & " "
            Select Case VarType(compareArray(i))
                Case vbString
                Sql = Sql & "'" & compareArray(i) & "'"
                Case 7 ''vbDate
                Sql = Sql & "#" & compareArray(i) & "#"
                Case 2 ''vbInteger
                Sql = Sql & compareArray(i)
            End Select
            If whereLen <> i Then
                Sql = Sql & " AND "
            End If
    Next i
       
result = sqlExecute(accessFile, dbTable, Sql)

End If

sqlDelete = result

End Function

Public Function sqlRun(accessFile, sqlIn) As Object

Dim con As Object
Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

fileExists = Dir(accessFile)
If fileExists <> "" Then '====
Set con = CreateObject("ADODB.connection")
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & accessFile

Sql = sqlIn

Set rs = CreateObject("ADODB.Recordset")

rs.Open Sql, con

Set sqlRun = rs

End If
End Function

Public Function sqlExecute(accessFile, dbTable, sqlIn) As Boolean
Dim con As Object
Dim rs As Object
Set rs = CreateObject("ADODB.Recordset")

fileExists = Dir(accessFile)
If fileExists <> "" Then '====
Set con = CreateObject("ADODB.connection")
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & accessFile

sqlGet = "Select * FROM " & dbTable
Sql = sqlIn

Set rs = CreateObject("ADODB.Recordset")

rs.Open sqlGet, con

con.Execute (Sql)

sqlExecute = True

rs.Close
rs.ActiveConnection = Nothing
Set rs = Nothing
End If

End Function
