VERSION 5.00
Begin VB.Form FrmDisconnect 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
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
      Left            =   2355
      MaxLength       =   75
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1125
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
      Left            =   2355
      MaxLength       =   75
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   645
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
      Left            =   5640
      Picture         =   "FrmDisconnect.frx":0000
      TabIndex        =   10
      Top             =   3765
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
      Left            =   4965
      Picture         =   "FrmDisconnect.frx":0442
      TabIndex        =   9
      Top             =   3765
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
      Left            =   4290
      Picture         =   "FrmDisconnect.frx":0884
      TabIndex        =   8
      Top             =   3765
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
      Left            =   3615
      Picture         =   "FrmDisconnect.frx":0CC6
      TabIndex        =   7
      Top             =   3765
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
      Left            =   2940
      Picture         =   "FrmDisconnect.frx":1108
      TabIndex        =   6
      Top             =   3765
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
      Left            =   2325
      TabIndex        =   5
      Top             =   3765
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
      Left            =   1860
      TabIndex        =   4
      Top             =   3765
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
      Left            =   735
      TabIndex        =   3
      Top             =   3765
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
      Left            =   285
      TabIndex        =   2
      Top             =   3765
      Width           =   465
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
      Left            =   2355
      MaxLength       =   75
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   195
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
      Left            =   6315
      Picture         =   "FrmDisconnect.frx":154A
      TabIndex        =   0
      Top             =   3765
      Width           =   675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   345
      X2              =   6915
      Y1              =   3615
      Y2              =   3615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Address :"
      Height          =   210
      Left            =   390
      TabIndex        =   15
      Top             =   705
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   6960
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1440
      TabIndex        =   14
      Top             =   3870
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Compnay Name :"
      Height          =   210
      Left            =   390
      TabIndex        =   13
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "FrmDisconnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS1 As New ADODB.Recordset
Dim AddEdit As String

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command11_Click()
If RS1.State = adStateClosed Then Exit Sub
If RS1.RecordCount <= 0 Then Exit Sub

On Error GoTo myer1
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
    On Error Resume Next
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub
    RS1.MoveFirst
    Call DisplayRecord
End Sub

Private Sub Command3_Click()
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
    On Error Resume Next
    If RS1.State = adStateClosed Then Exit Sub
    If RS1.RecordCount <= 0 Then Exit Sub
    RS1.MoveLast
    Call DisplayRecord

End Sub

Private Sub Command6_Click()
If RS1.State = adStateClosed Then Exit Sub
If RS1.RecordCount <= 0 Then Exit Sub
    AddEdit = "ADD"
    Call ClearText
    DE False, True
    TxtCompanyName.SetFocus
End Sub

Private Sub Command7_Click()
If RS1.State = adStateClosed Then Exit Sub
If RS1.RecordCount <= 0 Then Exit Sub

    AddEdit = "EDIT"
    DE False, True
    TxtCompanyName.SetFocus
End Sub

Private Sub Command8_Click()
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
        RS1("client_name") = TxtCompanyName.Text
        RS1("Address1") = TxtAddress1.Text
        RS1("Address2") = TxtAddress2.Text


        RS1.Update
        RS1.MoveLast
        Call DisplayRecord
    Else
        'rs1("company_name") = TxtCompanyName.Text
        RS1("Address1") = TxtAddress1.Text
        RS1("Address2") = TxtAddress2.Text

        RS1.Update
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
DE True, False
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    ClearText
    OpenCon
    If RS1.State = adStateOpen Then RS1.Close
    RS1.Open "select * from client_master order by client_name ", Cn, adOpenStatic, adLockBatchOptimistic
    If RS1.RecordCount > 0 Then
        RS1.MoveFirst
        Call DisplayRecord
    End If
    DE True, False
End Sub

Private Sub ClearText()
    Dim Ctl As Control
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is TextBox Then
            Ctl.Text = ""
        End If
    Next
End Sub

Private Sub DisplayRecord()
    'On Error Resume Next
    Call ClearText
    TxtCompanyName.Text = IIf(IsNull(RS1("client_name")) = True, "", RS1("client_name"))
    TxtAddress1.Text = IIf(IsNull(RS1("Address1")) = True, "", RS1("Address1"))
    TxtAddress2.Text = IIf(IsNull(RS1("Address2")) = True, "", RS1("Address2"))

    
    
    Label17.Caption = RS1.AbsolutePosition & "/" & RS1.RecordCount
End Sub

Private Sub DE(T1 As Boolean, T2 As Boolean)
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




