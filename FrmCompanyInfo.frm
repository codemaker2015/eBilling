VERSION 5.00
Begin VB.Form FrmCompanyInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compnay Info"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtVatNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2280
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "FrmCompanyInfo.frx":0000
      Top             =   2820
      Width           =   4440
   End
   Begin VB.TextBox TxtPin 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1970
      Width           =   3405
   End
   Begin VB.TextBox TxtCity 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1545
      Width           =   3405
   End
   Begin VB.TextBox TxtAddress2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1120
      Width           =   4575
   End
   Begin VB.TextBox TxtAddress1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   695
      Width           =   4575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5565
      Picture         =   "FrmCompanyInfo.frx":0006
      TabIndex        =   15
      Top             =   4050
      Width           =   675
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4890
      Picture         =   "FrmCompanyInfo.frx":0448
      TabIndex        =   14
      Top             =   4050
      Width           =   675
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4215
      Picture         =   "FrmCompanyInfo.frx":088A
      TabIndex        =   13
      Top             =   4050
      Width           =   675
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3540
      Picture         =   "FrmCompanyInfo.frx":0CCC
      TabIndex        =   12
      Top             =   4050
      Width           =   675
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2865
      Picture         =   "FrmCompanyInfo.frx":110E
      TabIndex        =   11
      Top             =   4050
      Width           =   675
   End
   Begin VB.CommandButton Command5 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2250
      TabIndex        =   10
      Top             =   4050
      Width           =   465
   End
   Begin VB.CommandButton Command4 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1785
      TabIndex        =   9
      Top             =   4050
      Width           =   465
   End
   Begin VB.CommandButton Command3 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   660
      TabIndex        =   8
      Top             =   4050
      Width           =   465
   End
   Begin VB.CommandButton Command2 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   210
      TabIndex        =   7
      Top             =   4050
      Width           =   465
   End
   Begin VB.TextBox TxtTelephone 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2395
      Width           =   3405
   End
   Begin VB.TextBox TxtCompanyName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   270
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6240
      Picture         =   "FrmCompanyInfo.frx":1550
      TabIndex        =   16
      Top             =   4050
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "VAT No :"
      Height          =   210
      Left            =   315
      TabIndex        =   23
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Pin :"
      Height          =   210
      Left            =   315
      TabIndex        =   22
      Top             =   2022
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "City :"
      Height          =   210
      Left            =   315
      TabIndex        =   21
      Top             =   1597
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Address :"
      Height          =   210
      Left            =   315
      TabIndex        =   20
      Top             =   747
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   285
      X2              =   6885
      Y1              =   3885
      Y2              =   3885
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1365
      TabIndex        =   19
      Top             =   4155
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Telephone :"
      Height          =   210
      Left            =   315
      TabIndex        =   18
      Top             =   2460
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Compnay Name :"
      Height          =   210
      Left            =   315
      TabIndex        =   17
      Top             =   322
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   270
      X2              =   6840
      Y1              =   3900
      Y2              =   3900
   End
End
Attribute VB_Name = "FrmCompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'            eBilling System
'             Version 1.0.0
'      Created by Mr. Atanu Maity
'          Date : 21-Aug-2006
'*************************************
' add/edit/delete company details
'      Used Table : company_master
'open the company_master
'display first record in form load
'add edit save delete and navigation
'*************************************

Option Explicit
Dim RS1 As New ADODB.Recordset
Dim AddEdit As String

Private Sub Command1_Click()
    '>>> close the form
    Unload Me
End Sub

Private Sub Command11_Click()
'>>> delete the record
If RS1.State = adStateClosed Then Exit Sub
If RS1.RecordCount <= 0 Then Exit Sub

On Error GoTo myer1
    '>>> confirm before delete
    If MsgBox("Delete the Record ? ", vbCritical + vbYesNo) = vbYes Then
        RS1.Delete
        Call ClearText
        Command4_Click
    End If
    Exit Sub
myer1:
    MsgBox "Error Occured : " & Err.Description, vbCritical
End Sub

Private Sub Command2_Click()
    '>>> move record ponter to first record
    '>>> display first record
    On Error Resume Next
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub
    RS1.MoveFirst
    Call DisplayRecord
End Sub

Private Sub Command3_Click()
    '>>> move back the record pointer and display current record
    On Error Resume Next
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub
If RS1.AbsolutePosition > 1 Then
    RS1.MovePrevious
Else
    MsgBox "First Record ..", vbInformation

    RS1.MoveFirst
End If
    Call DisplayRecord

End Sub

Private Sub Command4_Click()
    '>>> move next the record pointer and display current record
    On Error Resume Next
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub
If RS1.AbsolutePosition < RS1.RecordCount Then
    RS1.MoveNext
Else
    MsgBox "Last Record ..", vbInformation

    RS1.MoveLast
End If
    Call DisplayRecord

End Sub

Private Sub Command5_Click()
    '>>> move last the record pointer and display current record
    On Error Resume Next
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub
    RS1.MoveLast
    Call DisplayRecord

End Sub

Private Sub Command6_Click()
    '>>> prepare for add record, clear all text box, set flag to ADD
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub
    AddEdit = "ADD"
    Call ClearText
    DE False, True
    TxtCompanyName.SetFocus
End Sub

Private Sub Command7_Click()
    '>>> prepare for edit record,  set flag to EDIT
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub

    AddEdit = "EDIT"
    DE False, True
    TxtCompanyName.SetFocus
End Sub

Private Sub Command8_Click()
    '>>> save the record
    '>>> check for validation
    '>>> check the flag for ADD/Edit
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub

    On Error GoTo myer1
    If Trim(TxtCompanyName.Text) = "" Then
        MsgBox "Enter Company Name  ", vbCritical
        TxtCompanyName.SetFocus
        Exit Sub
    End If
    If AddEdit = "ADD" Then
        RS1.AddNew
        RS1("company_name") = TxtCompanyName.Text
        RS1("Address1") = TxtAddress1.Text
        RS1("Address2") = TxtAddress2.Text
        RS1("city") = TxtCity.Text
        RS1("pin") = TxtPin.Text
        RS1("telephone") = TxtTelephone.Text
        RS1("vatno") = TxtVatNo.Text

        RS1.Update
        RS1.MoveLast
        Call DisplayRecord
    Else
        RS1("Address1") = TxtAddress1.Text
        RS1("Address2") = TxtAddress2.Text
        RS1("city") = TxtCity.Text
        RS1("pin") = TxtPin.Text
        RS1("telephone") = TxtTelephone.Text
        RS1("vatno") = TxtVatNo.Text
        RS1.Update
        
        '>>> if it is edit after requery show the edited record

        Dim p As Integer
        p = RS1.AbsolutePosition
        RS1.Requery
        RS1.MoveFirst
        RS1.Move p - 1
        Call DisplayRecord
    End If
    DE True, False
    Exit Sub
myer1:
    MsgBox "Error Occured : " & Err.Description, vbCritical
End Sub

Private Sub Command9_Click()
    '>>> cancel save
    DE True, False
End Sub

Private Sub Form_Load()
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    '>>> reset connection
    '>>> clear all text
    ClearText
    OpenCon
    '>>> load alreday saved clent data
    '>>> and show the first record

    If RS1.State = adStateOpen Then RS1.Close
    RS1.Open "select * from company_master order by company_name ", Cn, adOpenDynamic, adLockOptimistic
    If RS1.RecordCount > 0 Then
        RS1.MoveFirst
        Call DisplayRecord
    End If
    DE True, False
End Sub

Private Sub ClearText()
    '>>> clear all text box in the form
    Dim Ctl As Control
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is TextBox Then
            Ctl.Text = ""
        End If
    Next
End Sub

Private Sub DisplayRecord()
    '>>> display current record
    On Error Resume Next
    Call ClearText
    TxtCompanyName.Text = IIf(IsNull(RS1("company_name")) = True, "", RS1("company_name"))
    TxtAddress1.Text = IIf(IsNull(RS1("Address1")) = True, "", RS1("Address1"))
    TxtAddress2.Text = IIf(IsNull(RS1("Address2")) = True, "", RS1("Address2"))
    TxtCity.Text = IIf(IsNull(RS1("city")) = True, "", RS1("city"))
    TxtPin.Text = IIf(IsNull(RS1("pin")) = True, "", RS1("pin"))
    TxtTelephone.Text = IIf(IsNull(RS1("telephone")) = True, "", RS1("telephone"))
    TxtVatNo.Text = IIf(IsNull(RS1("vatno")) = True, "", RS1("vatno"))
    
    
    Label17.Caption = RS1.AbsolutePosition & "/" & RS1.RecordCount
End Sub

Private Sub DE(T1 As Boolean, T2 As Boolean)
    '>>> enable disable buttons
    Command2.Enabled = T1
    Command3.Enabled = T1
    Command4.Enabled = T1
    Command5.Enabled = T1
    Command6.Enabled = T1
    Command7.Enabled = T1
    Command11.Enabled = T1
    Command8.Enabled = T2
    Command9.Enabled = T2
End Sub


