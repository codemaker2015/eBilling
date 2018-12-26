VERSION 5.00
Begin VB.Form FrmChangeCompany 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Current Company"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmChangeCompany.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3330
      TabIndex        =   3
      Top             =   2250
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Change"
      Default         =   -1  'True
      Height          =   390
      Left            =   1290
      TabIndex        =   2
      Top             =   2250
      Width           =   1365
   End
   Begin VB.ComboBox CmbCompanyName 
      Height          =   330
      Left            =   720
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   750
      Width           =   4425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Change current company to :"
      Height          =   210
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2760
   End
End
Attribute VB_Name = "FrmChangeCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'            eBilling System
'             Version 1.0.0
'      Created by Vishnu Sivan
'          Date : 21-Aug-2018
'*************************************
'     change the current company
'      Used Table : company_master
'Module to allow user to change the
'current comopany from the list
'set company name to global variable
'*************************************

Option Explicit
Dim Rs As New ADODB.Recordset


Private Sub Command1_Click()
    '>>> check the company nmae select by user
    '>>> frm the comapny_master table
    '>>> if record found set global variable
    '>>> otherwise warn user to select from the list.
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select * from company_master where company_name='" & CmbCompanyName.Text & "'", Cn, adOpenStatic, adLockReadOnly
    If Rs.RecordCount > 0 Then
        CompanyName = CmbCompanyName.Text
        FrmMain.LblCompanyName = CompanyName
        Unload Me
    Else
        MsgBox "Select company name from the list", vbExclamation
        CmbCompanyName.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Command2_Click()
    '>>> cloase the form
    Unload Me
End Sub

Private Sub Form_Load()
   
    '>>> reset the database connection
    If Cn.State = 1 Then Cn.Close
    OpenCon
    
    '>>> center the form

    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    '>>> fill the combo box with company name from company_master
    '>>> open record from company_master
    '>>> loop throgh recordset and add each company_name into combo box
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select * from company_master ", Cn, adOpenStatic, adLockReadOnly
    CmbCompanyName.Clear
    If Rs.RecordCount > 0 Then
        While Rs.EOF = False
            CmbCompanyName.AddItem Rs("company_name")
            Rs.MoveNext
        Wend
    End If
    If Rs.State = adStateOpen Then Rs.Close
    '>>> set the already selected company name from login form
    CmbCompanyName.Text = CompanyName
End Sub
