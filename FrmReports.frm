VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
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
   ScaleHeight     =   3225
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export Product List"
      Height          =   480
      Left            =   180
      TabIndex        =   4
      Top             =   2040
      Width           =   4650
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   3510
      Top             =   690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&CLOSE"
      Height          =   480
      Left            =   180
      TabIndex        =   3
      Top             =   2670
      Width           =   4650
   End
   Begin VB.CommandButton CmdBillSummary 
      Caption         =   "Bill Summary"
      Height          =   480
      Left            =   180
      TabIndex        =   2
      Top             =   1425
      Width           =   4650
   End
   Begin VB.CommandButton CmdProductSummary 
      Caption         =   "&Product Summary"
      Height          =   480
      Left            =   180
      TabIndex        =   1
      Top             =   810
      Width           =   4650
   End
   Begin VB.CommandButton CmdPrintBill 
      Caption         =   "Show/Print &Bill"
      Height          =   480
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4650
   End
End
Attribute VB_Name = "FrmReports"
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
'       Show report options
'
'
'allow user to select diffrent report
'*************************************

Option Explicit

Private Sub CmdBillSummary_Click()
    '>>> show bill summary
    FrmBillSummary.Show 1
End Sub

Private Sub CmdClose_Click()
    '>>> cloase the form
    Unload Me
End Sub

Private Sub CmdExport_Click()
    '>>> show export product form
    FrmExportData.Show 1
End Sub

Private Sub CmdPrintBill_Click()
    '>> show print bill
    FrmPrintBill.Show 1
End Sub

Private Sub CmdProductSummary_Click()
    '>>> show all product list report
    Cr1.WindowState = crptMaximized
    Cr1.ReportFileName = App.Path & "\reports\products.rpt"
    Cr1.DataFiles(0) = App.Path & "\data.mdb"
    Cr1.Action = 1
End Sub

Private Sub Form_Load()
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

End Sub
