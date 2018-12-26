VERSION 5.00
Begin VB.Form FrmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Billing System"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   405
      Picture         =   "FrmChangePassword.frx":030A
      ScaleHeight     =   225
      ScaleWidth      =   2265
      TabIndex        =   8
      Top             =   255
      Width           =   2265
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   285
         TabIndex        =   9
         Top             =   -15
         Width           =   1710
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3510
      TabIndex        =   7
      Top             =   2445
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   405
      Left            =   1020
      TabIndex        =   6
      Top             =   2445
      Width           =   1605
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2535
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1575
      Width           =   2670
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2535
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1185
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2535
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   645
      Width           =   2670
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00404040&
      Height          =   1875
      Left            =   210
      Top             =   330
      Width           =   5355
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   1875
      Left            =   210
      Top             =   330
      Width           =   5355
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   945
      Left            =   435
      Top             =   1050
      Width           =   4890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Confirm Password :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   540
      TabIndex        =   2
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "New Password :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   540
      TabIndex        =   1
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Old Password :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   540
      TabIndex        =   0
      Top             =   660
      Width           =   1425
   End
End
Attribute VB_Name = "FrmChangePassword"
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
' change the password for current user
'      Used Table : user_master
'check the oldpassword
'compare new password and confirm password
'update new password in user master
'*************************************

Option Explicit

Private Sub Command1_Click()
    '>>> validation for ole password it cannot blank
    If Text1.Text = "" Then
        MsgBox "Enter Old Password ...", vbExclamation
        Text1.SetFocus
        Exit Sub
    End If
    '>>> validation for new password it cannot blank
    If Text2.Text = "" Or Text3.Text = "" Then
        MsgBox "Enter New Passoed ...", vbExclamation
        Text3.Text = ""
        Text2.SetFocus
        Exit Sub
    End If
    '>>> compare new password and confirm password
    If Text2.Text <> Text3.Text Then
        MsgBox "Confirm password dosenot match with new password ...", vbExclamation
        Text3.Text = ""
        Text2.SetFocus
        Exit Sub
    End If
    '>>> reset the database connection
    OpenCon
    Dim Rs As New ADODB.Recordset
    If Rs.State = adStateOpen Then Rs.Close
    Dim P1 As String
    P1 = Text1.Text
    '>>> check the old password
    Rs.Open "select * from user_master where USER_name ='" & UserName & "' and user_password ='" & P1 & "'", Cn, adOpenDynamic, adLockOptimistic
    If Rs.RecordCount > 0 Then
        '>>> update old passwor dwith new password
        Rs("user_password") = Text2.Text
        Rs.Update
        MsgBox "Password Changed Successfully", vbInformation
        Unload Me
        
    
    Else
        MsgBox "Invalid Old Password cannot continue ...", vbExclamation, "Invalid Password"
        Text2.Text = ""
        Text3.Text = ""
        SendKeys "{home}+{end}"
        Text1.SetFocus
        Exit Sub
    End If
        
    
    
    
End Sub

Private Sub Command2_Click()
    '>>> cloase the form
    Unload Me
End Sub

Private Sub Form_Load()
    '>>> center form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
End Sub

