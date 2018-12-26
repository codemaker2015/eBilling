VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCompactAndRepairDatabase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compact And Repair Database"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCompactAndRepairDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   195
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6465
      Top             =   1965
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Start"
      Height          =   450
      Left            =   750
      TabIndex        =   3
      Top             =   2115
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   465
      Left            =   4200
      TabIndex        =   2
      Top             =   2130
      Width           =   1980
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1365
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "NOTE : Before continuing the process close all the activity of the database from all the workstation and server."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   570
      Left            =   285
      TabIndex        =   4
      Top             =   270
      Width           =   6330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Compact the Database"
      Height          =   225
      Left            =   2370
      TabIndex        =   1
      Top             =   1140
      Width           =   2250
   End
End
Attribute VB_Name = "FrmCompactAndRepairDatabase"
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
' comapct/shrink the access database
'      Used Table : NA
'check the repairdb.mdb file in application path
'if it is already their delete the file
'use DBENGINE CompactDatabase function to comapct the access database
'create new compacted tempdb.mdb from data.mdb.
'delete old data.mdb and rename tempdb.mdb to data.mdb
'*************************************

Option Explicit

Dim dbE As New DAO.DBEngine

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
'>>> if any connection open close all the connection
If Cn.State = 1 Then Cn.Close
Dim x As String
'>>> check allready file is there or not
x = Dir(App.Path & "\repairedDB.mdb")
'>>> if file present delete the file
If x <> "" Then Kill App.Path & "\repairedDB.mdb"
Timer1.Enabled = True
'>>> compact teh database
dbE.CompactDatabase App.Path & "\data.mdb", App.Path & "\RepairedDB"
'>>> delete old database
Kill App.Path & "\data.mdb"
'>>> rename the new database to old database
Name App.Path & "\repairedDB.mdb" As App.Path & "\data.mdb"
'>>> open connection
Call OpenCon
End Sub

Private Sub Form_Load()
'>>> center the form
Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
'>>> set the progress bar initial value
ProgressBar1.Min = 0
ProgressBar1.Max = 100
End Sub

Private Sub Timer1_Timer()
    '>> show the progress of compact process
    If ProgressBar1.Value < 100 Then
        ProgressBar1.Value = ProgressBar1.Value + 10
    Else
        MsgBox "Process Complete Successfully ..", vbInformation
        ProgressBar1.Value = 0 'Reset the min value
        Timer1.Enabled = False 'Disable the Timer
    End If
End Sub
