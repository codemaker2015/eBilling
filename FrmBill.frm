VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmBill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bill - INVOICE"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
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
   ScaleHeight     =   8355
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1155
      ScaleHeight     =   675
      ScaleWidth      =   3840
      TabIndex        =   30
      Top             =   885
      Width           =   3840
   End
   Begin VB.ComboBox TxtCompany 
      Height          =   1020
      Left            =   1335
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "TxtCompany"
      Top             =   525
      Width           =   3555
   End
   Begin VB.TextBox TxtSerial 
      Height          =   375
      Left            =   6870
      MaxLength       =   50
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1110
      Width           =   2775
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   1830
      Top             =   7215
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox ChkPrint 
      Caption         =   "Print after Save"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   7905
      Width           =   1845
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   405
      Left            =   6921
      TabIndex        =   26
      Top             =   7845
      Width           =   1275
   End
   Begin VB.ComboBox cmbval 
      Height          =   1440
      Left            =   4680
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Text            =   "cmbval"
      Top             =   4155
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TXTVAL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D8FEFA&
      Height          =   315
      Left            =   2730
      TabIndex        =   10
      Top             =   4470
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   8400
      TabIndex        =   14
      Top             =   7845
      Width           =   1275
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New"
      Height          =   405
      Left            =   2490
      TabIndex        =   11
      Top             =   7815
      Width           =   1275
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   " &Print"
      Height          =   405
      Left            =   5444
      TabIndex        =   13
      Top             =   7845
      Width           =   1275
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   405
      Left            =   3975
      TabIndex        =   12
      Top             =   7830
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid Mf1 
      Height          =   3915
      Left            =   300
      TabIndex        =   8
      Top             =   2685
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6906
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtAddress2 
      Height          =   375
      Left            =   1320
      MaxLength       =   74
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2070
      Width           =   3555
   End
   Begin VB.TextBox TxtAddress1 
      Height          =   375
      Left            =   1320
      MaxLength       =   74
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1605
      Width           =   3555
   End
   Begin VB.TextBox TxtCompany_ 
      Height          =   375
      Left            =   2100
      MaxLength       =   74
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   -180
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.TextBox TxtChalanNo 
      Height          =   375
      Left            =   6870
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox TxtLocation 
      Height          =   375
      Left            =   6870
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1575
      Width           =   2775
   End
   Begin VB.CommandButton CmdGetBill 
      Caption         =   "&Get Bill"
      Height          =   405
      Left            =   8580
      TabIndex        =   15
      Top             =   660
      Width           =   1035
   End
   Begin VB.TextBox TxtInvoiceNo 
      Height          =   375
      Left            =   6870
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   660
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DtDate 
      Height          =   375
      Left            =   6870
      TabIndex        =   3
      Top             =   150
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   38950
   End
   Begin VB.Label lblGST 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8760
      TabIndex        =   32
      Top             =   6720
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "GST : "
      Height          =   210
      Left            =   6960
      TabIndex        =   31
      Top             =   6720
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Serial :"
      Height          =   210
      Left            =   5235
      TabIndex        =   28
      Top             =   1185
      Width           =   660
   End
   Begin VB.Label LblTotalAmount 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8730
      TabIndex        =   25
      Top             =   7035
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Amount :"
      Height          =   210
      Left            =   6090
      TabIndex        =   24
      Top             =   7035
      Width           =   1380
   End
   Begin VB.Label LblRsWord 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   360
      TabIndex        =   23
      Top             =   7035
      Width           =   5430
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   90
      X2              =   9870
      Y1              =   7665
      Y2              =   7665
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   90
      X2              =   9870
      Y1              =   7635
      Y2              =   7635
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   150
      X2              =   9930
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   150
      X2              =   9930
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Address :"
      Height          =   210
      Left            =   240
      TabIndex        =   22
      Top             =   1635
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Company :"
      Height          =   210
      Left            =   255
      TabIndex        =   21
      Top             =   555
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Chalan No :"
      Height          =   210
      Left            =   5235
      TabIndex        =   20
      Top             =   2115
      Width           =   1110
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Location :"
      Height          =   210
      Left            =   5235
      TabIndex        =   19
      Top             =   1650
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Invoice No :"
      Height          =   210
      Left            =   5235
      TabIndex        =   18
      Top             =   727
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Date :"
      Height          =   210
      Left            =   5235
      TabIndex        =   17
      Top             =   210
      Width           =   1320
   End
   Begin VB.Label LblCompanyName 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "MODERN FURNITURES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   255
      TabIndex        =   16
      Top             =   150
      Width           =   3015
   End
End
Attribute VB_Name = "FrmBill"
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
'         New/Modify Bill Module
'      Used Table : Bill
'                 : Bill Details
'                 : Product
'                 : Company
'Module to make new or modify bill,
'Print bill after save the record
'also we make editable flexgrid with
'text box and combobox
'*************************************

Option Explicit
'>>> declare form global variable

Dim AddEdit As String
Dim Rs As New ADODB.Recordset
Dim sno As Integer
Dim SavePrint As String
Dim gst As Integer

Private Sub CmdClose_Click()
    '>>> check the caption of the button
    '>>> close the form or cancel the save
    If CmdClose.Caption = "&Close" Then
        Unload Me
    Else
        DE True, False
    End If

End Sub

Private Sub CmdDelete_Click()
    '>>> confirm for deletion of bill
    '>>> if user select YES delete the record from bill and bill_details
    '>>> clear the seleted data from the screen for deleted bill

    If MsgBox("Record will delete permantley ?", vbYesNo + vbCritical) = vbYes Then
        Cn.Execute "delete from bill_details where bill_sno =" & sno
        Cn.Execute "delete from bill where sno =" & sno
        AddEdit = ""
        
        
        CmdGetBill.Caption = "&Find"
        CmdDelete.Enabled = False
        Call ClearField
        If TxtInvoiceNo.Enabled = True Then
        TxtInvoiceNo.SetFocus
        End If
    End If
    
End Sub

Private Sub CmdGetBill_Click()
    '>>> find the bill details
    '>>> find the bill by invoice no
    AddEdit = ""
    If CmdGetBill.Caption = "&Get Bill" Then
        TxtInvoiceNo.Enabled = True
        
        Call ClearField
        TxtInvoiceNo.BackColor = vbYellow
        TxtInvoiceNo.SetFocus
        CmdGetBill.Caption = "&Find"
        
    Else
        Dim RS1 As New ADODB.Recordset
        RS1.Open "select * from bill where invoice_no =" & Val(TxtInvoiceNo.Text) & " and cname='" & CompanyName & "'", Cn, adOpenStatic, adLockReadOnly
        If RS1.RecordCount > 0 Then
            '>>> show details from bill table
            DtDate.Value = RS1("invoice_date")
            TxtLocation.Text = RS1("location")
            TxtChalanNo.Text = RS1("chalan_no")
            TxtCompany.Text = RS1("customer_name")
            TxtAddress1.Text = RS1("customer_address1")
            TxtAddress2.Text = RS1("customer_address2")
            LblTotalAmount = RS1("total_amt")
            LblRsWord.Caption = RS1("amt_word")
            TxtSerial.Text = RS1("serial")
            
            '>>> show data from  bill_details
            Dim Rs2 As New ADODB.Recordset
            If Rs2.State = adStateOpen Then Rs2.Close
            Dim Rs3 As New ADODB.Recordset
            Rs2.Open "select * from bill_details where bill_sno=" & RS1("sno") & " order by sno ", Cn, adOpenStatic, adLockReadOnly
            If Rs2.RecordCount > 0 Then
                Dim i As Integer
                Rs2.MoveFirst
                For i = 0 To Rs2.RecordCount - 1
                    If Rs3.State = adStateOpen Then Rs3.Close
                    Rs3.Open "select * from product_master where sno =" & Rs2("prod_sno"), Cn, adOpenStatic, adLockReadOnly
                    If Rs3.RecordCount > 0 Then
                        Mf1.TextMatrix(i + 1, 1) = Rs3("prod_sub_type")
                    End If
                    If Rs3.State = adStateOpen Then Rs3.Close
                    Mf1.TextMatrix(i + 1, 2) = Rs2("qty")
                    Mf1.TextMatrix(i + 1, 3) = Rs2("rate")
                    Mf1.TextMatrix(i + 1, 4) = Rs2("amt")
                    
                    Rs2.MoveNext
                Next
            End If
            If Rs2.State = adStateOpen Then Rs2.Close
            
            CmdGetBill.Caption = "&Get Bill"
            CmdDelete.Enabled = True
            cmbval.Visible = False
            CmdPrint.Enabled = True
            CmdSave.Enabled = True
            
            '>>> locak the buttons
            Mf1.Enabled = True
            cmbval.Enabled = True
            TXTVAL.Enabled = True
            TxtCompany.Locked = False
            TxtAddress1.Locked = False
            TxtAddress2.Locked = False
            DtDate.Enabled = True
            TxtLocation.Locked = False
            TxtChalanNo.Locked = False
        Else
            MsgBox "No Previous Details found for invoice..." & TxtInvoiceNo.Text, vbExclamation
            CmdGetBill.Caption = "&Find"
            TxtInvoiceNo.Enabled = True
            TxtInvoiceNo.SetFocus
            CmdDelete.Enabled = False
            CmdPrint.Enabled = False
            CmdSave.Enabled = False
            
            Mf1.Enabled = False
            cmbval.Enabled = False
            TXTVAL.Enabled = False
            TxtCompany.Locked = True
            TxtAddress1.Locked = True
            TxtAddress2.Locked = True
            DtDate.Enabled = False
            TxtLocation.Locked = True
            TxtChalanNo.Locked = True
            
        End If
        If RS1.State = adStateOpen Then RS1.Close
    End If
End Sub

Private Sub CmdNew_Click()
    '>>> claer the screen for entering data for new bill
    '>>> enable/disable buttons
    lblGST.Caption = ""
    
    DE False, True
    
    AddEdit = "ADD"
    
    Call ClearField
    
    '>>> get the new system id from bill
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select max(sno),max(invoice_no),max(serial)  from bill ", Cn, adOpenStatic, adLockReadOnly
    If Rs.RecordCount > 0 Then
        sno = Rs(0) + 1
        TxtInvoiceNo.Text = Rs(1) + 1
        TxtSerial.Text = Rs(0) + 1
    End If
    If Rs.State = adStateOpen Then Rs.Close
    TxtCompany.SetFocus
End Sub

Private Sub CmdPrint_Click()
    'NOTE : it is not the right solution to call crystal report by temp using temp table
    'some time it is a good practice for complecated databse relation table
    'This may not run properly in multi user environment
    'Better approch is passing value by SelectionFormula in crystal report
    'but anyway it is a working solution
    '>>> find the bill sno from seleted invoice no
    '>>> if record found
    '>>> delete temp bill na dbill_details
    '>>> insert from bill,bill_details to temp_bill, teemp_bill_details
    
    Dim RS1 As New ADODB.Recordset
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select sno from bill where invoice_no=" & Val(TxtInvoiceNo.Text) & " and cname ='" & CompanyName & "'", Cn, adOpenStatic, adLockReadOnly
    If RS1.RecordCount > 0 Then
        Cn.Execute "delete from temp_bill_details"
        Cn.Execute "delete from temp_bill"
        Cn.Execute "insert into temp_bill select * from bill where sno=" & RS1("sno")
        Cn.Execute "insert into temp_bill_details select * from bill_details where bill_sno=" & RS1("sno")
        Call OpenCon
        
        '>>> call crystal report
        Cr1.WindowState = crptMaximized
        Cr1.ReportFileName = App.Path & "\reports\bill.rpt"
        Cr1.DataFiles(0) = App.Path & "\data.mdb"
        Cr1.Action = 1
    Else
        MsgBox "No Bill found select/enter invoice no for print", vbExclamation
        Exit Sub
    End If
    
    
    
End Sub

Private Sub CmdSave_Click()
    '>>> validation
    '>>> check the required field
    If Trim(TxtCompany.Text) = "" Then
        MsgBox "Enter Company Name...", vbExclamation
        TxtCompany.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(TxtInvoiceNo.Text) = False Then
        MsgBox "Enter only Numeric Invoice No...", vbExclamation
        TxtInvoiceNo.SetFocus
        Exit Sub
    End If
    
    If Trim(TxtLocation.Text) = "" Then
        MsgBox "Enter Location...", vbExclamation
        TxtLocation.SetFocus
        Exit Sub
    End If
    
    If Trim(TxtChalanNo.Text) = "" Then
        MsgBox "Enter Chalan No...", vbExclamation
        TxtChalanNo.SetFocus
        Exit Sub
    End If
    If IsNumeric(TxtChalanNo.Text) = False Then
        MsgBox "Enter only Numeric Chalan No...", vbExclamation
        TxtChalanNo.SetFocus
        Exit Sub
    End If
    
    '>>> reset the transaction
    Call OpenCon
    
    '>>> create transaction for insert bill and bil details
    Cn.BeginTrans
    Dim RS1 As New ADODB.Recordset
    
    '>>> check wheather we need to insert or edit the record
    '>>> if it is edit, then delete the old bill and insert new record
    If AddEdit <> "ADD" Then
        If RS1.State = adStateOpen Then RS1.Close
        RS1.Open "select * from bill where  invoice_no =" & Val(TxtInvoiceNo.Text) & " and cname='" & CompanyName & "'", Cn, adOpenStatic, adLockReadOnly
        If RS1.RecordCount > 0 Then
            sno = RS1("sno")
        End If
        If RS1.State = adStateOpen Then RS1.Close
        Cn.Execute "delete from bill_details where bill_sno =" & sno
        Cn.Execute "delete from bill where sno =" & sno
        AddEdit = ""
    End If
    
    '>>> check for product master
    '>>> check the grid
    '>>> wheather there is a product or not
    '>>> wheater they enter any quantity or not
    '>>> wheatehr there is any price or not
    '>>> if any thing goes wrong show message
    Dim cc As Integer
    Dim i As Integer
    For i = 1 To Mf1.Rows - 1
        Dim Ch As Boolean
        Ch = False
        If Trim(Mf1.TextMatrix(i, 1)) = "" Then
            Ch = True
        End If
        If RS1.State = adStateOpen Then RS1.Close
        RS1.Open "select sno from product_master where prod_sub_type =" & Chr(34) & Mf1.TextMatrix(i, 1) & Chr(34), Cn, adOpenStatic, adLockReadOnly
        If RS1.RecordCount <= 0 Then

            Ch = True
        End If
        If RS1.State = adStateOpen Then RS1.Close
        If Val(Mf1.TextMatrix(i, 2)) = 0 Then
            Ch = True
        End If
        If Val(Mf1.TextMatrix(i, 3)) = 0 Then
            Ch = True
        End If
        If Val(Mf1.TextMatrix(i, 4)) = 0 Then
            Ch = True
        End If
        If Ch = False Then
            cc = cc + 1
        End If
    Next
    If cc = 0 Then
        MsgBox "No Bill details found for save", vbExclamation
        Exit Sub
    End If
    
    '>>> check for duplicate invoice no
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select *  from bill where invoice_no=" & Val(TxtInvoiceNo.Text), Cn, adOpenStatic, adLockReadOnly
    If Rs.RecordCount > 0 Then
        MsgBox "Invalid Invoice No cannot save..", vbExclamation
        TxtInvoiceNo.SetFocus
        Exit Sub
    End If
    If Rs.State = adStateOpen Then Rs.Close
    
    '>>> save record in bill  table
    RS1.Open "select * from bill where 1=2", Cn, adOpenDynamic, adLockOptimistic
    RS1.AddNew
    RS1("sno") = sno
    RS1("invoice_no") = Val(TxtInvoiceNo.Text)
    RS1("cname") = CompanyName
    RS1("invoice_date") = Format(DtDate.Value, "dd-mmm-yy")
    RS1("location") = Trim(TxtLocation.Text)
    RS1("chalan_no") = Val(TxtChalanNo.Text)
    RS1("customer_name") = Trim(TxtCompany.Text)
    RS1("customer_address1") = Trim(TxtAddress1.Text)
    RS1("customer_address2") = Trim(TxtAddress2.Text)
    RS1("total_amt") = Val(LblTotalAmount.Caption)
    RS1("GST") = Val(lblGST.Caption)
    RS1("discount") = 0
    RS1("amt_word") = LblRsWord
    RS1("paid_type") = "NA"
    RS1("cheque_no") = "NA"
    RS1("entry_date") = Now
    RS1("serial") = TxtSerial.Text
    RS1.Update
    If RS1.State = 1 Then RS1.Close
    
    
    '>>> vaildate each row before save in details
    Dim LastSno As Integer
    Dim ProdSno As Integer
    Dim Rs2 As New ADODB.Recordset
    If Rs2.State = 1 Then Rs2.Close
    Rs2.Open "select max(sno) from bill_details", Cn, adOpenStatic, adLockReadOnly
    If Rs2.RecordCount > 0 Then
        LastSno = IIf(IsNull(Rs2(0)) = True, 0, Rs2(0)) + 1
    End If
    For i = 1 To Mf1.Rows - 1
        
        Ch = False
        If Trim(Mf1.TextMatrix(i, 1)) = "" Then
            Ch = True
        End If
        If RS1.State = adStateOpen Then RS1.Close
        RS1.Open "select sno from product_master where prod_sub_type =" & Chr(34) & Mf1.TextMatrix(i, 1) & Chr(34), Cn, adOpenStatic, adLockReadOnly
        If RS1.RecordCount > 0 Then
            ProdSno = RS1(0)
        Else
            Ch = True
        End If
        If RS1.State = adStateOpen Then RS1.Close
        If Val(Mf1.TextMatrix(i, 2)) = 0 Then
            Ch = True
        End If
        If Val(Mf1.TextMatrix(i, 3)) = 0 Then
            Ch = True
        End If
        If Val(Mf1.TextMatrix(i, 4)) = 0 Then
            Ch = True
        End If
        If Ch = False Then
            
            '>>> insert in bill details for each validated grid row
            If Rs2.State = 1 Then Rs2.Close
            Rs2.Open "select * from bill_details where 1=2", Cn, adOpenDynamic, adLockOptimistic
            Rs2.AddNew
            Rs2("sno") = LastSno
            Rs2("bill_sno") = sno
            Rs2("prod_sno") = ProdSno
            Rs2("qty") = Val(Mf1.TextMatrix(i, 2))
            Rs2("rate") = Val(Mf1.TextMatrix(i, 3))
            Rs2("amt") = Val(Mf1.TextMatrix(i, 4))
            Rs2.Update
            If Rs2.State = 1 Then Rs2.Close
            LastSno = LastSno + 1
        End If
    Next


    '>>> commit the transaction
    Cn.CommitTrans
    MsgBox "Bill Saved", vbInformation
    If ChkPrint.Value = 1 Then
        '>>> call the report for print of the saved bill
        CmdPrint_Click
    End If
    '>>> prepare for new bill entry
    AddEdit = ""
    Call ClearField
    Call CmdNew_Click
    
End Sub

Private Sub DtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    '>>> move the cursor to new field
   If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub DtDate_KeyPress(KeyAscii As Integer)
    '>>> move the cursor to new field
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
    '>>> cnter the form
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    '>>> reset connection
    Call OpenCon
    LblCompanyName.Caption = CompanyName
    DtDate.Value = Now
    
    '>>> clear the form for new bill entry
    Call ClearField
    lblGST.Caption = "0"
    '>>> enable/diable buttons
    DE True, False
    
    '>>> load last status for bill print after saved
    Dim ChkV As Integer
    ChkV = Val(GetSetting("billsystem", "print", "checkprint", "1"))
    ChkPrint.Value = ChkV
    
    '>>> load clent name from the table
    Rs.Open "select client_name from client_master where client_name is not null order by client_name", Cn, adOpenStatic, adLockReadOnly
    
    While Not Rs.EOF
        TxtCompany.AddItem Rs(0)
        Rs.MoveNext
    Wend
    
    
    Picture1.Visible = True
    
End Sub
Private Sub DE(T1 As Boolean, T2 As Boolean)

    '>>> enable/disable buttons
    CmdGetBill.Enabled = T1
    CmdNew.Enabled = T1
    CmdSave.Enabled = T2
    CmdPrint.Enabled = T2
    
    TXTVAL.Enabled = T2
    cmbval.Enabled = T2
    Mf1.Enabled = T2
    If T1 = True Then
        CmdClose.Caption = "&Close"
    Else
        CmdClose.Caption = "&Cancel"
    End If
    
    TxtCompany.Locked = T1
    TxtAddress1.Locked = T1
    TxtAddress2.Locked = T1
    DtDate.Enabled = T2
    TxtLocation.Locked = T1
    TxtChalanNo.Locked = T1
    TxtInvoiceNo.BackColor = vbWhite
    CmdDelete.Enabled = False
End Sub

Private Sub ClearField()
    '>>> clear the fields
    TxtCompany.Text = ""
    TxtAddress1.Text = ""
    TxtAddress2.Text = ""
    TxtInvoiceNo.Text = ""
    TxtLocation.Text = ""
    LblRsWord.Caption = ""
    LblTotalAmount.Caption = ""
    TxtChalanNo.Text = ""
    TxtSerial.Text = ""
    
    '>>> for grid edit
    Call set_heading
    Call move_textbox
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    '>>> save last status for bill print after saved in the registry

    If ChkPrint.Value = 1 Then
        SaveSetting "billsystem", "print", "checkprint", "1"
    Else
        SaveSetting "billsystem", "print", "checkprint", "0"
    End If
End Sub

Private Sub TxtAddress1_KeyPress(KeyAscii As Integer)
    '>>> move the focus to next control
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtAddress2_KeyPress(KeyAscii As Integer)
    '>>> move the focus to next control
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtChalanNo_KeyPress(KeyAscii As Integer)
    '>>> move the focus to next control

    If KeyAscii = 13 Then
        cmbval.SetFocus
    End If
End Sub

Private Sub TxtCompany_GotFocus()
    '>>> make listbox as editable textbox
    TxtCompany.Height = 1020
    Picture1.Visible = False
End Sub

Private Sub TxtCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    '>>> show the address of the seleted company
    Dim RS1 As New ADODB.Recordset
    If RS1.State = adStateOpen Then RS1.Close
    RS1.Open "select * from client_master where client_name='" & TxtCompany.Text & "'", Cn
    If RS1.RecordCount > 0 Then
        TxtAddress1.Text = IIf(IsNull(RS1("address1")) = True, "", RS1("address1"))
        TxtAddress2.Text = IIf(IsNull(RS1("address2")) = True, "", RS1("address2"))
    Else
        TxtAddress1.Text = ""
        TxtAddress2.Text = ""
    End If
    
    If RS1.State = adStateOpen Then RS1.Close
    
    RS1.Open "select max(sno) from bill", Cn
    If RS1.RecordCount > 0 Then
        TxtSerial.Text = Val(RS1(0)) + 1
    End If
    
    If RS1.State = adStateOpen Then RS1.Close
    RS1.Open "select max(invoice_no) from bill", Cn
    If RS1.RecordCount > 0 Then
        TxtInvoiceNo.Text = Val(RS1(0)) + 1
    End If
    
    If RS1.State = adStateOpen Then RS1.Close
End Sub

Private Sub TxtCompany_KeyPress(KeyAscii As Integer)
    '>>> move the focus to next control

    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtCompany_LostFocus()
    '>>> move the focus to next control

    TxtCompany.Height = 330
    Picture1.Visible = True
End Sub

Private Sub TxtInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
    '>>> call getbill by enter key in invoice no text box
    If KeyCode = 13 And CmdGetBill.Caption = "&Find" Then
        CmdGetBill_Click
    End If
        
End Sub

Private Sub TxtLocation_KeyPress(KeyAscii As Integer)
    '>>> move the focus to next control
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtSerial_KeyPress(KeyAscii As Integer)
    '>>> move the focus to next control
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TXTVAL_KeyPress(KeyAscii As Integer)
'>>> make flex gird editable move th textbox in gid cell for entering value
'>>> check wheathe we press enter key
'>>> if yes..move the control to next grod
Dim t As Integer

If KeyAscii = 13 And Mf1.Col = 2 Then
    If IsNumeric(TXTVAL.Text) = False Then
        KeyAscii = 0
        Exit Sub
    End If
    Mf1.Text = TXTVAL.Text
    '>>> show total
    Mf1.TextMatrix(Mf1.Row, Mf1.Col + 2) = Val(Mf1.TextMatrix(Mf1.Row, Mf1.Col + 1)) * TXTVAL.Text
    lblGST.Caption = Val(lblGST.Caption) + gst * Val(Mf1.TextMatrix(Mf1.Row, Mf1.Col + 2)) / 100
  
    Dim x As Integer
    Dim T1 As Double
    For x = 1 To Mf1.Rows - 1
        T1 = T1 + Val(Mf1.TextMatrix(x, 4))
    Next
    LblTotalAmount.Caption = T1 + Val(lblGST.Caption)
    LblRsWord = RsWord(LblTotalAmount.Caption)
    
    If Mf1.Col <> Mf1.Cols - 3 Then
        Mf1.Col = Mf1.Col + 1
    
           
    Else
    
        
        If Mf1.Row <> Mf1.Rows - 2 Then
            '>>> go to next row
            Mf1.Row = Mf1.Row + 1
        Else
            '>>> add new rows
            Mf1.Rows = Mf1.Rows + 1
            
            '>>> set the current row
            Mf1.Row = Mf1.Row + 1
            '>>> set sr no
            Mf1.TextMatrix(Mf1.Row, 0) = Val(Mf1.TextMatrix(Mf1.Row - 1, 0)) + 1
        End If
        Mf1.Col = 1
    End If
    move_textbox
    '>>> select the text in text box
    TXTVAL.SelStart = 0
    TXTVAL.SelLength = Len(TXTVAL.Text)
End If

End Sub

Private Sub cmbval_KeyPress(KeyAscii As Integer)
'>>> make flex gird editable move th combobox in gid cell for seleting value
'>>> check wheathe we press enter key
'>>> if yes..move the control to next grod

Dim t As Integer

If KeyAscii = 13 Then
    '>>> find prod rate
    Dim RS1 As New ADODB.Recordset
    If RS1.State = adStateOpen Then RS1.Close
    RS1.Open "select * from product_master where prod_sub_type=" & Chr(34) & cmbval.Text & Chr(34), Cn, adOpenStatic, adLockReadOnly
    If RS1.RecordCount > 0 Then
        Mf1.TextMatrix(Mf1.Row, Mf1.Col + 2) = RS1("rate")
        gst = RS1("GST")
    Else
        KeyAscii = 0
        Exit Sub
    End If

    Mf1.Text = cmbval.Text
    
    '>>> show total
    Mf1.TextMatrix(Mf1.Row, 4) = Val(Mf1.TextMatrix(Mf1.Row, 2)) * Val(Mf1.TextMatrix(Mf1.Row, 3))
    Dim x As Integer
    Dim T1 As Double
    
    For x = 1 To Mf1.Rows - 1
        T1 = T1 + Val(Mf1.TextMatrix(x, 4))
    Next
    LblTotalAmount.Caption = T1
    LblRsWord = RsWord(LblTotalAmount.Caption)
     
    If Mf1.Col <> Mf1.Cols - 2 Then
        
        Mf1.Col = Mf1.Col + 1
    Else
        If Mf1.Row <> Mf1.Rows - 1 Then
            Mf1.Row = Mf1.Row + 1
            
        Else
            '>>> add new rows
            Mf1.Rows = Mf1.Rows + 1
            
            '>>> set the current row
            Mf1.Row = Mf1.Row + 1
            
            '>>> set sr no
            Mf1.TextMatrix(Mf1.Row, 0) = Val(Mf1.TextMatrix(Mf1.Row - 1, 0)) + 1
        End If
        Mf1.Col = 1
    End If
    move_textbox
    cmbval.SelStart = 0
    cmbval.SelLength = Len(cmbval.Text)
End If
End Sub

Public Sub set_heading()
'>>> creating for the grid

Dim K As Integer
Dim t As Integer
    Mf1.Clear
    Mf1.Refresh
    Mf1.Rows = 30
    Mf1.Cols = 5
    
    Mf1.Row = 0
    Mf1.RowHeight(0) = 600
    
    Mf1.Col = 0
    Mf1.ColWidth(0) = 1000
    Mf1.CellForeColor = vbBlue
    Mf1.CellFontBold = True
    Mf1.CellAlignment = 4
    Mf1.Text = "Sr."
    
    Mf1.Col = 1
    Mf1.ColWidth(1) = 4200
    Mf1.CellForeColor = vbBlue
    Mf1.CellFontBold = True
    Mf1.CellAlignment = 4
    Mf1.Text = "Particulars"
    
    Mf1.Col = 2
    Mf1.ColWidth(2) = 1200
    Mf1.CellForeColor = vbBlue
    Mf1.CellFontBold = True
    Mf1.CellAlignment = 4
    Mf1.Text = "Quantity"
    
    Mf1.Col = 3
    Mf1.ColWidth(3) = 1200
    Mf1.CellForeColor = vbBlue
    Mf1.CellFontBold = True
    Mf1.CellAlignment = 4
    Mf1.Text = "Rate"
    
    Mf1.Col = 4
    Mf1.ColWidth(4) = 1200
    Mf1.CellForeColor = vbBlue
    Mf1.CellFontBold = True
    Mf1.CellAlignment = 4
    Mf1.Text = "Amount"
    
    Mf1.TextMatrix(1, 0) = "1"


Mf1.Row = 0
For K = 0 To Mf1.Cols - 1
    Mf1.Col = K
    Mf1.CellFontBold = True
Next

Mf1.Row = 1
Mf1.Col = 1

'>>> set serial from 1.2...
For K = 1 To Mf1.Rows - 1
    Mf1.TextMatrix(K, 0) = K
Next
Mf1.Row = 1
End Sub

Private Sub MF1_EnterCell()
    '>>> call appropriate control for edit the grid
        If Mf1.Col = 1 Then
        '>>> visble combo box for select product
        cmbval.Visible = True
        TXTVAL.Visible = False
        If cmbval.Visible = True Then
            If cmbval.Enabled = True Then
                cmbval.SetFocus
            End If
        End If
        
        cmbval.Clear
        Dim Rs As New ADODB.Recordset
        If Rs.State = 1 Then Rs.Close
        
        '>>>Fill item
        If Mf1.Col = 1 Then
            Rs.Open "select   prod_sub_type from product_master order by prod_sub_type", Cn, adOpenStatic, adLockReadOnly
            
            While Not Rs.EOF
                cmbval.AddItem Rs(0)
                Rs.MoveNext
            Wend
            
        ElseIf Mf1.Col = 3 Then
            cmbval.AddItem ""
        End If
        
    Else
        '>>> visble text box for entring quantity
        cmbval.Visible = False
        TXTVAL.Visible = True
        If TXTVAL.Visible = True Then
            If TXTVAL.Enabled = True Then
                TXTVAL.SetFocus
            End If
        End If
        
    End If
    
    
    Call move_textbox
End Sub

Public Sub move_textbox()
    '>>align textbox as per grid cell and set text
    TXTVAL.Left = Mf1.CellLeft + Mf1.Left
    TXTVAL.Top = Mf1.CellTop + Mf1.Top
    TXTVAL.Width = Mf1.CellWidth
    TXTVAL.Height = Mf1.CellHeight
    TXTVAL.Text = Mf1.Text
    
    '>>align combo box as per grid cell and set text
    cmbval.Left = Mf1.CellLeft + Mf1.Left
    cmbval.Top = Mf1.CellTop + Mf1.Top
    cmbval.Width = Mf1.CellWidth
    cmbval.Text = Mf1.Text
End Sub
