VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmBillSummary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bill Summary"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
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
   ScaleHeight     =   1380
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Cr1 
      Left            =   2475
      Top             =   315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdGetBill 
      Caption         =   "Bill &Summary"
      Height          =   420
      Left            =   4125
      TabIndex        =   1
      Top             =   180
      Width           =   1875
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   4110
      TabIndex        =   0
      Top             =   660
      Width           =   1875
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1890
      TabIndex        =   2
      Top             =   165
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20512769
      CurrentDate     =   38951
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   1890
      TabIndex        =   3
      Top             =   720
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20512769
      CurrentDate     =   38951
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bill Date From :"
      Height          =   210
      Left            =   195
      TabIndex        =   5
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Bill Date From :"
      Height          =   210
      Left            =   195
      TabIndex        =   4
      Top             =   780
      Width           =   1425
   End
End
Attribute VB_Name = "FrmBillSummary"
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
'         Show summery of bill
'      Used Table : bill
'                 : bill_details
'show bill summary for seleted date
'range, show report in crystal report
'move data into temp table and show
'report from temp table
'*************************************
Option Explicit

Private Sub CmdClose_Click()
    '>>> close the form
    Unload Me
End Sub

Private Sub CmdGetBill_Click()
    'NOTE : it is not the right solution to call crystal report by temp using temp table
    'some time it is a good practice for complecated databse relation table
    'This may not run properly in multi user environment
    'Better approch is passing value by SelectionFormula in crystal report
    'but anyway it is a working solution
    '>>> find the bill sno from seleted invoice no
    '>>> if record found
    '>>> delete temp bill na dbill_details
    '>>> insert from bill,bill_details to temp_bill, teemp_bill_details
    
    Cn.Execute "delete from temp_bill_details"
    Cn.Execute "delete from temp_bill"
    Cn.Execute "insert into temp_bill select * from bill  where invoice_date>=#" & Format(DTPicker1.Value, "dd-mmm-yy") & "# and invoice_date<=#" & Format(DTPicker2.Value, "dd-mmm-yy") & "# and cname='" & CompanyName & "' "
    Cn.Execute "insert into temp_bill_details select * from bill_details where bill_sno in (  select sno from bill  where  invoice_date >=#" & Format(DTPicker1.Value, "dd-mmm-yy") & "# and invoice_date<=#" & Format(DTPicker2.Value, "dd-mmm-yy") & "# and cname='" & CompanyName & "')"
    Call OpenCon
    
    '>>> open crystal report
    Cr1.DataFiles(0) = App.Path & "\data.mdb"
    Cr1.WindowState = crptMaximized
    Cr1.ReportFileName = App.Path & "\reports\billsummary.rpt"
    Cr1.Action = 1

End Sub

Private Sub Form_Load()
    '>>> cnter the form
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    '>>> show the current date
    DTPicker1.Value = Date
    DTPicker2.Value = Date
End Sub
