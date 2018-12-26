VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmExportData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Product List"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExcel 
      Caption         =   "&Export to Excel"
      Height          =   405
      Left            =   4830
      TabIndex        =   3
      Top             =   5955
      Width           =   2415
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   7365
      TabIndex        =   2
      Top             =   5955
      Width           =   1650
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mf1 
      Height          =   5220
      Left            =   180
      TabIndex        =   1
      Top             =   570
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   9208
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdProductMaster 
      Caption         =   "Product Master"
      Height          =   405
      Left            =   180
      TabIndex        =   0
      Top             =   105
      Width           =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4155
      TabIndex        =   4
      Top             =   150
      Width           =   75
   End
End
Attribute VB_Name = "FrmExportData"
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
' Export product_master in grid and ms-excel
'      Used Table : product_master
'open the product_master in grid
'display record in flex grid with gropu by product type
'*************************************
Option Explicit

Private Sub CmdClose_Click()
    '>>> close the form
    Unload Me
End Sub

Private Sub CmdExcel_Click()
    '>>>export data into ms excel from grid with formatting
    '>>> check the grid
    If Mf1.TextMatrix(0, 0) = "" Then
        MsgBox "No Records Available for Exporting ... ", vbExclamation
        Exit Sub
    End If
    Label1.Caption = "WAIT ... Generate Excel "
    Label1.Refresh
    
    '>>> creating excel object variable
    Dim ex As New Excel.Application
    Dim wb As New Workbook
    Dim Es As New Worksheet
    Set wb = ex.Workbooks.Add
    Set Es = wb.Worksheets(1)
    Dim i As Integer
    Dim j As Integer
    '>>> set excel columns width as per flex grid columns width
    For i = 0 To Mf1.Cols - 1
        Mf1.Row = 1
        Mf1.Col = i
        Es.Columns(ReturnAlphabet(i + 1) & ":" & ReturnAlphabet(i + 1)).ColumnWidth = Mf1.CellWidth / 110
    Next
    '>>> set data from grid to excel row, column wise
    Dim K As Integer
    For i = 0 To Mf1.Rows - 1
        For j = 0 To Mf1.Cols - 1
            ex.Cells(i + 1, j + 1) = Mf1.TextMatrix(i, j)
        Next
    Next
    
    Dim R1 As String
    Dim R2 As String
    R2 = ReturnAlphabet(Mf1.Cols) & "1"
    
    '>>> formatting excel
    Dim x As Range
    '>>>head
    Set x = Es.Range("A1:" & R2)
    x.Font.Bold = True
    x.Font.ColorIndex = 40
    x.Interior.ColorIndex = 9
    x.Interior.Pattern = xlSolid
    x.HorizontalAlignment = xlCenter
    x.VerticalAlignment = xlBottom

    '>>>border
    R2 = ReturnAlphabet(Mf1.Cols) & Mf1.Rows - 1
    Set x = Es.Range("A1:" & R2)

    With x.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With x.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With x.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With x.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With x.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With x.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    '>>> FILL DETAILS
    Set x = Es.Range("A2:" & R2)
    x.Interior.ColorIndex = 40
    
    '>>TOTAL
    
    R2 = ReturnAlphabet(Mf1.Cols) & Mf1.Rows
    Set x = Es.Range("A" & Mf1.Rows & ":" & R2)
    x.Font.Bold = True
    x.Font.ColorIndex = 9
'
    Es.Name = "Report"
    
    ex.Visible = True
    
    ex.Quit
    Set wb = Nothing
    Set Es = Nothing
    Set ex = Nothing

    
    '>>> process complete

    
    Label1.Caption = "Ready"
    Label1.Refresh
    

End Sub

Private Sub CmdProductMaster_Click()
    '>>> reset the grid
    
    Mf1.Rows = 2
    Mf1.Cols = 3
    Mf1.Clear
    Mf1.Refresh
    
    Mf1.Row = 0
    
    Mf1.Col = 0
    Mf1.ColWidth(0) = 800
    Mf1.Text = "Sr"
    Mf1.CellAlignment = 4
    Mf1.CellFontName = "Arial"
    Mf1.Font.Size = 10
    Mf1.Font.Bold = True
    Mf1.CellForeColor = vbBlue
    Mf1.CellBackColor = vbCyan
    
    
    Mf1.Col = 1
    Mf1.ColWidth(1) = 2500
    Mf1.Text = "Product Type"
    Mf1.CellAlignment = 4
    Mf1.CellFontName = "Arial"
    Mf1.Font.Size = 10
    Mf1.Font.Bold = True
    Mf1.CellForeColor = vbBlue
    Mf1.CellBackColor = vbCyan
    
    Mf1.Col = 2
    Mf1.ColWidth(2) = 5000
    Mf1.Text = "Product"
    Mf1.CellAlignment = 4
    Mf1.CellFontName = "Arial"
    Mf1.Font.Size = 10
    Mf1.Font.Bold = True
    Mf1.CellForeColor = vbBlue
    Mf1.CellBackColor = vbCyan
    
    '>>> find distinct product type from product master
    '>>> loop all product type
    Dim RS1 As New ADODB.Recordset
    Dim Rs2 As New ADODB.Recordset
    RS1.Open "select distinct prod_type from product_master", Cn, adOpenStatic, adLockReadOnly
    Dim i As Integer
    Dim j As Integer
    For i = 0 To RS1.RecordCount - 1
        Me.Caption = i + 1
        Mf1.Row = Mf1.Rows - 1
        Mf1.Col = 0
        Mf1.Text = i + 1
        
        Mf1.Col = 1
        Mf1.Text = RS1("prod_type")
        '>>> query product master for each prod type from outer loop
        If Rs2.State = adStateOpen Then Rs2.Close
        Rs2.Open "select prod_sub_type from product_master where prod_type ='" & RS1("prod_type") & "' order by prod_sub_type", Cn, adOpenStatic, adLockReadOnly
        For j = 0 To Rs2.RecordCount - 1
            Mf1.Row = Mf1.Rows - 1
            Mf1.Col = 2
            Mf1.Text = Rs2(0)
            
            Mf1.Rows = Mf1.Rows + 1
            Rs2.MoveNext
        Next
        RS1.MoveNext
    Next
    
End Sub




Private Sub Form_Load()
    '>>> cnter the form
    Me.Left = (Screen.Width - Me.Width)
    Me.Top = (Screen.Height - Me.Height)
    
End Sub
