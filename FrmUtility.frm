VERSION 5.00
Begin VB.Form FrmUtility 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utility"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
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
   ScaleHeight     =   4785
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAddModifyClient 
      Caption         =   "Add/Modify Client Info"
      Height          =   555
      Left            =   90
      TabIndex        =   6
      Top             =   2768
      Width           =   4875
   End
   Begin VB.CommandButton CmdAddModifyCompany 
      Caption         =   "Add/Modify Company Info"
      Height          =   555
      Left            =   90
      TabIndex        =   3
      Top             =   2106
      Width           =   4875
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "CLOSE"
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   4095
      Width           =   4875
   End
   Begin VB.CommandButton CmdChangePassword 
      Caption         =   "Change Password"
      Height          =   555
      Left            =   90
      TabIndex        =   4
      Top             =   3430
      Width           =   4875
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "Restore Database"
      Height          =   555
      Left            =   90
      TabIndex        =   2
      Top             =   1444
      Width           =   4875
   End
   Begin VB.CommandButton CmdBackUp 
      Caption         =   "Backup Database"
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   782
      Width           =   4875
   End
   Begin VB.CommandButton CmdRepair 
      Caption         =   "Repair MDB "
      Height          =   555
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "FrmUtility"
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
'       Show options for utilities
'
'
'allow user to select diffrent options
'*************************************
Option Explicit


Private Sub CmdAddModifyClient_Click()
    FrmClientInfo.Show 1
End Sub

Private Sub CmdAddModifyCompany_Click()
    FrmCompanyInfo.Show 1
End Sub

Private Sub CmdBackUp_Click()
FrmBackupDatabase.Show 1
End Sub

Private Sub CmdChangePassword_Click()
    FrmChangePassword.Show 1
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdRepair_Click()
    FrmCompactAndRepairDatabase.Show 1
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub CmdRestore_Click()
    MsgBox "Copy last backup file, rename it to data.mdb and then paste it in application folder." & vbCrLf & vbCrLf & "Contact program vendor..it may cause lost of data, if backup is not current", vbCritical

End Sub

Private Sub Form_Load()
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

    '>>> check user type and show diffrent options as per user rights
    If LCase(UserType) <> "admin" Then
        CmdRepair.Enabled = False
        CmdBackUp.Enabled = False
        CmdAddModifyCompany.Enabled = False
        CmdRestore.Enabled = False
    End If
End Sub
