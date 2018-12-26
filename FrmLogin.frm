VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login to eBilling System"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   390
      Left            =   2250
      TabIndex        =   7
      Top             =   2550
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   3960
      TabIndex        =   6
      Top             =   2550
      Width           =   1470
   End
   Begin VB.TextBox TxtPassword 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2010
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1860
      Width           =   3555
   End
   Begin VB.TextBox TxtUserName 
      Height          =   330
      Left            =   2010
      TabIndex        =   4
      Text            =   "admin"
      Top             =   1260
      Width           =   3555
   End
   Begin VB.ComboBox CmbCompanyName 
      Height          =   330
      Left            =   2010
      TabIndex        =   1
      Top             =   360
      Width           =   3555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password :"
      Height          =   210
      Left            =   330
      TabIndex        =   3
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User Name :"
      Height          =   210
      Left            =   330
      TabIndex        =   2
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Company Name :"
      Height          =   210
      Left            =   330
      TabIndex        =   0
      Top             =   390
      Width           =   1620
   End
End
Attribute VB_Name = "FrmLogin"
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
'             Login Module
'      Used Table : user_master
'Module to check user login and load
'user rights as per user type.
'*************************************

Option Explicit
Dim Rs As New ADODB.Recordset
Private Sub Command1_Click()
    '>>> check wheather user name and password are blank
    '>>> if its is blan warn user to enter
    If TxtUserName.Text = "" Or TxtPassword.Text = "" Then
        MsgBox "Enter user name and password ...", vbExclamation
        TxtUserName.SetFocus
        Exit Sub
    End If
    
    '>>> check for entered company
    '>>> query to database and if no record found warn user to select company from the list.
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select * from company_master where company_name='" & CmbCompanyName.Text & "'", Cn, adOpenStatic, adLockReadOnly
    If Rs.RecordCount > 0 Then
        CompanyName = CmbCompanyName.Text
    Else
        MsgBox "Select company name from the list", vbExclamation
        CmbCompanyName.SetFocus
        Exit Sub
    End If

    '>>> check for username and password
    '>>> query to user_master with user_name and password
    '>>> if no record found check warn user for enter valid user namne and password
    '>>> if record found store user_nmae, user_type in global variable for future use.
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select * from user_master where USER_name ='" & TxtUserName.Text & "' and user_password ='" & TxtPassword & "'", Cn, adOpenStatic, adLockReadOnly
    If Rs.RecordCount > 0 Then
        CheckLogin = True
        UserName = IIf(IsNull(Rs("USER_name").Value) = True, "NA", Rs("USER_name").Value)
        UserType = IIf(IsNull(Rs("user_type").Value) = True, "NA", Rs("user_type").Value)
        
        
        Unload Me
        
    Else
        MsgBox "Invalid User Name and Password ... ", vbExclamation, "Login Error "
        TxtPassword.Text = ""
        TxtUserName.SetFocus
        Exit Sub
    End If

End Sub

Private Sub Command2_Click()
    '>>> close the application
    End
    Set FrmLogin = Nothing
End Sub


Private Sub Form_Load()
    '>>> open the global connection
    If Cn.State = 1 Then Cn.Close
    OpenCon
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    '>>> fill the combo box with all company_name from company master
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '>>> release all the object variable used by form
    Set FrmLogin = Nothing
End Sub
