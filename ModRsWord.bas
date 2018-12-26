Attribute VB_Name = "ModRsWord"
'*************************************
'            eBilling System
'             Version 1.0.0
'      Created by Vishnu Sivan
'          Date : 21-Aug-2018
'*************************************
'
'Module to convert Rs to Word
'123.50 = One Hundred Twent Three and Ffty Paise Only
'*************************************
Option Explicit

'display digit to words
'123 -- > One hundred and twentythree only.

Public Function RsWord(t As String) As String
On Error Resume Next

    Dim la As Boolean
    Dim th As Boolean
    Dim l As Integer
    Dim l1 As Integer
    Dim i As String
    Dim r As String
    Dim P1 As String
    Dim NL As Integer
    Dim p As Boolean
    Dim pl As Integer
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim z As Boolean
    
    

        
    NL = InStr(t, ".")
    p = True
    pl = Len(t) - NL
    
    
    
    If NL = 0 Then
        NL = Len(t): p = False
        l = Len(t)
    Else
        l = NL - 1
    End If
    l1 = l
    
    While l > 0
        i = Mid(t, l1 - (l - 1), 1)
        
            If l = 7 Then la = True
            If l = 7 And i = "0" And Mid(t, l1 - (l - 2), 1) <> "0" Then
                r = r & SingleDigit(Mid(t, l1 - (l - 2), 1)) & " Lackhs "
            End If
            If l = 7 And (i <> "0" And i <> "1") Then
                r = r & Tenths(i) & " " & SingleDigit(Mid(t, l1 - (l - 2), 1)) & " Lackhs "
            End If
            If l = 7 And i = "1" Then
                r = r & TwoDigit(Mid(t, l1 - (l - 2), 1)) & " Lackhs "
            End If
            
            If l = 6 And i <> "0" And la = False Then
                r = r & SingleDigit(i) & " Lackhs "
            End If
            
            If l = 5 Then th = True
            If l = 5 And i = "0" And Mid(t, l1 - (l - 2), 1) <> "0" Then
                
                r = r & SingleDigit(Mid(t, l1 - (l - 2), 1)) & " Thousand "
            End If
            If l = 5 And (i <> "0" And i <> "1") Then
                r = r & Tenths(i) & " " & SingleDigit(Mid(t, l1 - (l - 2), 1)) & " Thousand "
            End If
            If l = 5 And i = "1" Then
                r = r & TwoDigit(Mid(t, l1 - (l - 2), 1)) & " Thousand "
            End If
            If l = 4 And i <> "0" And th = False Then
                r = r & SingleDigit(i) & " Thousand "
            End If

            If l = 3 And i <> "0" Then
                r = r & SingleDigit(i) & " Hundred "
            End If
            If l = 2 And (i <> "0" And i <> "1") Then
                r = r & Tenths(i)
            End If
            If l = 2 And i = "1" Then
                r = r & TwoDigit(Mid(t, l1 - (l - 1), 1)) & " "
            End If
            If l = 1 And Mid(t, l1 - (l - 2), 1) <> "1" Then
                r = r & " " & SingleDigit(i)
            End If
        l = l - 1
    Wend
    

    If p = True Then
        l1 = NL + 1
        While pl > 0
        i = Mid(t, Len(t) - pl + 1, 1)
            If pl = 2 And (i <> "0" And i <> "1") Then
                P1 = P1 & Tenths(i)
            End If
            If pl = 2 And i = "1" Then
                P1 = P1 & TwoDigit(Mid(t, Len(t) - pl + 2, 1)) & " "
            End If
            If pl = 1 And Mid(t, Len(t) - pl, 1) <> "1" Then
                P1 = P1 & " " & SingleDigit(i)
            End If
            pl = pl - 1
        Wend
    End If
    If r <> "" And P1 <> "" Then
        RsWord = "Rupees " & r & " and " & P1 & " Paise Only"
    End If
    If P1 = "" And r <> "" Then
            RsWord = "Rupees " & r & " Only"
    End If
    If r = "" And P1 <> "" Then
            RsWord = P1 & " Paise Only"
    End If
    If r = "" And P1 = "" Then
        RsWord = "Nil"
    End If
End Function
Private Function TwoDigit(d As String) As String
        Dim s As String
        If d = "1" Then s = "Eleven"
        If d = "2" Then s = "Twelve"
        If d = "3" Then s = "Thirteen"
        If d = "4" Then s = "Forteen"
        If d = "5" Then s = "Fifteen"
        If d = "6" Then s = "Sixteen"
        If d = "7" Then s = "Seventeen"
        If d = "8" Then s = "Eighteen"
        If d = "9" Then s = "Nineteen"
        If d = "0" Then s = "Ten"
        TwoDigit = s
End Function
Private Function Tenths(d As String) As String
        Dim s As String
        If d = "2" Then s = "Twenty"
        If d = "3" Then s = "Thirty"
        If d = "4" Then s = "Forty"
        If d = "5" Then s = "Fifty"
        If d = "6" Then s = "Sixty"
        If d = "7" Then s = "Seventy"
        If d = "8" Then s = "Eighty"
        If d = "9" Then s = "Ninety"
       Tenths = s
End Function

Private Function SingleDigit(d As String) As String
        Dim s As String
        If d = "1" Then s = "One"
        If d = "2" Then s = "Two"
        If d = "3" Then s = "Three"
        If d = "4" Then s = "Four"
        If d = "5" Then s = "Five"
        If d = "6" Then s = "Six"
        If d = "7" Then s = "Seven"
        If d = "8" Then s = "Eight"
        If d = "9" Then s = "Nine"
        If d = "0" Then s = ""
 
        SingleDigit = s
End Function

