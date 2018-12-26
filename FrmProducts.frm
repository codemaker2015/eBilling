VERSION 5.00
Begin VB.Form FrmProducts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Master"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
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
   ScaleHeight     =   5535
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGST 
      Height          =   375
      Left            =   7920
      TabIndex        =   19
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   7005
      TabIndex        =   17
      Top             =   4620
      Width           =   1305
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   5670
      TabIndex        =   16
      Top             =   4620
      Width           =   1305
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   435
      Left            =   8340
      TabIndex        =   15
      Top             =   4740
      Width           =   1305
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   435
      Left            =   7000
      TabIndex        =   14
      Top             =   4740
      Width           =   1305
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New"
      Height          =   435
      Left            =   5660
      TabIndex        =   13
      Top             =   4740
      Width           =   1305
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "&Find"
      Height          =   435
      Left            =   4320
      TabIndex        =   12
      Top             =   4740
      Width           =   1305
   End
   Begin VB.CheckBox ChkActive 
      Caption         =   "Active Product"
      Height          =   315
      Left            =   5580
      TabIndex        =   11
      Top             =   3420
      Width           =   2025
   End
   Begin VB.TextBox TxtRate 
      Height          =   330
      Left            =   5580
      TabIndex        =   10
      Top             =   2760
      Width           =   1400
   End
   Begin VB.TextBox TxtName 
      Height          =   330
      Left            =   5580
      TabIndex        =   8
      Top             =   2040
      Width           =   4065
   End
   Begin VB.ComboBox CmbType 
      Height          =   330
      Left            =   5580
      TabIndex        =   6
      Text            =   "--Select--"
      Top             =   1380
      Width           =   4065
   End
   Begin VB.ListBox LstProductSubType 
      Height          =   4260
      Left            =   150
      TabIndex        =   2
      Top             =   990
      Width           =   3555
   End
   Begin VB.ComboBox CmbProductType 
      Height          =   330
      Left            =   150
      TabIndex        =   1
      Text            =   "--Select--"
      Top             =   480
      Width           =   3555
   End
   Begin VB.Label Label3 
      Caption         =   "GST"
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Rate :"
      Height          =   210
      Left            =   4320
      TabIndex        =   9
      Top             =   2820
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      Height          =   210
      Left            =   4320
      TabIndex        =   7
      Top             =   2100
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Type :"
      Height          =   210
      Left            =   4320
      TabIndex        =   5
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label LblSr 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   210
      Left            =   5610
      TabIndex        =   4
      Top             =   510
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Product Id :"
      Height          =   210
      Left            =   4320
      TabIndex        =   3
      Top             =   540
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Product Type :"
      Height          =   210
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   1365
   End
End
Attribute VB_Name = "FrmProducts"
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
'       Maintain Product Master
'       Used Table : product_master
'
'Module to allow user to select product
'add/modify product details
'*************************************

Option Explicit
'>> decalre form level valriable
Dim Rs As New ADODB.Recordset
Dim AddEdit As String

Private Sub CmbProductType_Change()
    '>>> as per product type fill the product list
    Dim QrStr As String
    If CmbProductType.Text = "ALL" Then
        QrStr = "select prod_sub_type from product_master order by prod_sub_type"
    Else
        QrStr = "select prod_sub_type from product_master where prod_type='" & CmbProductType.Text & "' order by prod_sub_type"
    End If
    LstProductSubType.Clear
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open QrStr, Cn, adOpenStatic, adLockReadOnly
    While Rs.EOF = False
        LstProductSubType.AddItem Rs("prod_sub_type")
        Rs.MoveNext
    Wend
    '>>> select the first product
    If LstProductSubType.ListCount > 0 Then
        LstProductSubType.ListIndex = 0
        DisplayRecord
    End If
End Sub

Private Sub CmbProductType_Click()
    '>>> call change event
    CmbProductType_Change
End Sub



Private Sub CmdCancel_Click()
    '>>> cancel update
    ED False, True
    DisplayRecord
End Sub

Private Sub CmdClose_Click()
    '>>> close the fron
    Unload Me
End Sub

Private Sub CmdEdit_Click()
    '>>> set flag to edit
    ED True, False
    AddEdit = "EDIT"
End Sub

Private Sub CmdNew_Click()
    '>>> set the flag to add
    '>>> claer text box
    LblSr.Caption = 0
    CmbType.Text = ""
    TxtName.Text = ""
    TxtRate.Text = 0
    ChkActive.Value = 1
    
    ED True, False
    
    AddEdit = "ADD"
End Sub

Private Sub CmdSave_Click()
    '>>> validate the entry
    If Trim(CmbType.Text) = "" Then
        MsgBox "Select or Enter product type.", vbExclamation
        CmbType.SetFocus
        Exit Sub
    End If
    If Trim(TxtName.Text) = "" Then
        MsgBox "Enter product name.", vbExclamation
        TxtName.SetFocus
        Exit Sub
    End If
    If InStr(1, TxtName.Text, Chr(34)) > 0 Then
        MsgBox "Don't use double qoute in product name.", vbExclamation
        TxtName.SetFocus
        Exit Sub
    End If
    If IsNumeric(TxtRate.Text) = False Then
        MsgBox "Enter rate, numeric only", vbExclamation
        TxtRate.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtGST.Text) = False Then
        MsgBox "Enter rate, numeric only", vbExclamation
        txtGST.SetFocus
        Exit Sub
    End If
    
    '>>> check the flag from add/edit
    If AddEdit = "ADD" Then
        If Rs.State = adStateOpen Then Rs.Close
        Rs.Open "select max(sno) +1 from product_master ", Cn, adOpenStatic, adLockReadOnly
        Dim sno As Integer
        sno = Rs(0)
        If Rs.State = adStateOpen Then Rs.Close
        Rs.Open "select * from product_master where 1=2", Cn, adOpenDynamic, adLockOptimistic
        Rs.AddNew
        Rs("sno") = sno
        Rs("prod_type") = CmbType.Text
        Rs("prod_sub_type") = TxtName.Text
        Rs("rate") = Val(TxtRate.Text)
        Rs("is_active") = Val(ChkActive.Value)
        Rs("GST") = Val(txtGST.Text)
        Rs.Update
        Rs.Close
    Else
        If Rs.State = adStateOpen Then Rs.Close
        Rs.Open "select * from product_master where sno=" & Val(LblSr.Caption), Cn, adOpenDynamic, adLockOptimistic
        Rs("prod_type") = CmbType.Text
        Rs("prod_sub_type") = TxtName.Text
        Rs("rate") = Val(TxtRate.Text)
        Rs("is_active") = Val(ChkActive.Value)
        Rs("GST") = Val(txtGST.Text)
        Rs.Update
        Rs.Close
    End If
    
    
    '>>> dispaly and update lists
    Dim OldPType As String
    OldPType = CmbType.Text
    Dim OldPName As String
    OldPName = TxtName.Text
    
    CmbType.Clear
    CmbProductType.Clear
    CmbProductType.AddItem "ALL"
    CmbProductType.Text = OldPType
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select distinct prod_type from product_master order by prod_type", Cn, adOpenStatic, adLockReadOnly
    While Rs.EOF = False
        CmbProductType.AddItem Rs("prod_type")
        CmbType.AddItem Rs("prod_type")
        Rs.MoveNext
    Wend
    '>>> fill the product list again with updated/inserted records
    Dim QrStr As String
    If CmbProductType.Text = "ALL" Then
        QrStr = "select prod_sub_type from product_master order by prod_sub_type"
    Else
        QrStr = "select prod_sub_type from product_master where prod_type='" & CmbProductType.Text & "' order by prod_sub_type"
    End If
    LstProductSubType.Clear
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open QrStr, Cn, adOpenStatic, adLockReadOnly
    While Rs.EOF = False
        LstProductSubType.AddItem Rs("prod_sub_type")
        Rs.MoveNext
    Wend
    '>>> show the first record
    If LstProductSubType.ListCount > 0 Then
        LstProductSubType.Text = OldPName
        DisplayRecord
    End If

    '>>> enable/diable button
    ED False, True
    
End Sub

Private Sub Form_Load()
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    '>>> reset connection
    OpenCon
    
    ED False, True
    CmbType.Clear
    
    '>>> fill the product type
    CmbProductType.Clear
    CmbProductType.AddItem "ALL"
    CmbProductType.Text = "ALL"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select distinct prod_type from product_master order by prod_type", Cn, adOpenStatic, adLockReadOnly
    While Rs.EOF = False
        'CmbProductType.AddItem Rs("prod_type")
        CmbType.AddItem Rs("prod_type")
        Rs.MoveNext
    Wend
    
    '>>> fill the product sub type
    Dim QrStr As String
    If CmbProductType.Text = "ALL" Then
        QrStr = "select prod_sub_type from product_master order by prod_sub_type"
    Else
        QrStr = "select prod_sub_type from product_master where prod_type='" & CmbProductType.Text & "' order by prod_sub_type"
    End If
    LstProductSubType.Clear
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open QrStr, Cn, adOpenStatic, adLockReadOnly
    While Rs.EOF = False
        LstProductSubType.AddItem Rs("prod_sub_type")
        Rs.MoveNext
    Wend
    '>>> select the first record
    If LstProductSubType.ListCount > 0 Then
        LstProductSubType.ListIndex = 0
        DisplayRecord
    End If
End Sub
Private Sub DisplayRecord()
    '>>> display record as per selected product name
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select * from product_master where prod_sub_type=" & Chr(34) & LstProductSubType.Text & Chr(34), Cn, adOpenStatic, adLockReadOnly
    If Rs.RecordCount > 0 Then
        LblSr.Caption = Rs("sno")
        CmbType.Text = Rs("prod_type")
        TxtName.Text = Rs("prod_sub_type")
        TxtRate.Text = Rs("rate")
        txtGST.Text = Rs("GST")
        ChkActive.Value = Rs("is_active")
    Else
        LblSr.Caption = ""
        CmbType.Text = ""
        TxtName.Text = ""
        TxtRate.Text = ""
        txtGST.Text = ""
        ChkActive.Value = 1
    
    End If
End Sub
Private Sub ED(T1 As Boolean, T2 As Boolean)
    '>>> enable/disable button
    CmdSave.Visible = T1
    CmdCancel.Visible = T1
    
    CmdFind.Visible = T2
    CmdNew.Visible = T2
    CmdEdit.Visible = T2
    CmdClose.Visible = T2
    
    CmbType.Locked = T2
    TxtName.Locked = T2
    TxtRate.Locked = T2
    ChkActive.Enabled = T1
End Sub

Private Sub LstProductSubType_Click()
    DisplayRecord
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    '>>> SELECT FROM LIST
    If CmbProductType.Text <> "ALL" Then
        CmbProductType.Text = "ALL"
    End If
    If KeyCode = vbKeyDown Then
        If LstProductSubType.ListIndex < LstProductSubType.ListCount - 1 Then
            LstProductSubType.ListIndex = LstProductSubType.ListIndex + 1
        End If
    End If
    If KeyCode = vbKeyUp Then
        If LstProductSubType.ListIndex > 0 Then
            LstProductSubType.ListIndex = LstProductSubType.ListIndex - 1
        End If
    End If
End Sub

