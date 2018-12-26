Attribute VB_Name = "ModGen"
'*************************************
'            eBilling System
'             Version 1.0.0
'      Created by Vishnu Sivan
'          Date : 21-Aug-2018
'*************************************
'
'declare global variable and procedure
'
'*************************************
Option Explicit

Public Cn As New ADODB.Connection
Public CheckLogin As Boolean
Public UserName As String
Public UserType As String
Public CompanyName As String
Public Sub OpenCon()
    '>>> open connction
    If Cn.State = 1 Then Cn.Close
    Cn.ConnectionString = "provider=microsoft.jet.oledb.4.0; data source= " & App.Path & "\data.mdb"
    Cn.CursorLocation = adUseClient
    Cn.Open

End Sub

Public Function newsno(ByVal table As String) As Integer
    '>>> find max sno for passing table
    Dim Rs As New ADODB.Recordset
    If Rs.State = 1 Then Rs.Close
    Rs.Open "select max(sno) from " & table, Cn, adOpenStatic, adLockReadOnly
    If IsNull(Rs(0)) = True Then
        newsno = 1
    Else
        newsno = Val(Rs(0)) + 1
    End If

End Function



Public Function ReturnAlphabet(n As Integer) As String
    '>>> return alphabel as per supplied no
    '>>> like 1 - A,2-B, 26-Z, 27-AA, 256-IV
    '>>> this function is used to excel formatting to set column value in range
    If n < 0 Or n > 256 Then
        MsgBox "Invalid Invalid range is 1-256", vbQuestion
        Exit Function
    End If
    
    Dim i As Integer
    Dim r As Integer
    Dim s As String
    Dim R1 As Integer
    If n <= 26 Then
        s = Chr(n + 64)
    Else
        r = n Mod 26
        R1 = n / 26
        s = Chr(R1 + 64) & Chr(r + 64)
    End If
    ReturnAlphabet = s
End Function
