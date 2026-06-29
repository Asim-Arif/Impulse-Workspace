VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmPaymentsLedger 
   Appearance      =   0  'Flat
   Caption         =   "Payment Ledger"
   ClientHeight    =   7155
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaymentsLedger.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   10455
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      Height          =   7140
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   10140
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&R e f r e s h"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8805
         TabIndex        =   19
         Top             =   915
         Width           =   1200
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2985
         TabIndex        =   11
         Top             =   6375
         Width           =   5235
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   1770
            TabIndex        =   14
            Top             =   195
            Width           =   1665
            Caption         =   "Print"
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   3465
            TabIndex        =   13
            Top             =   195
            Width           =   1665
            Caption         =   "Close            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAdd 
            Height          =   360
            Left            =   75
            TabIndex        =   12
            Top             =   195
            Width           =   1665
            Caption         =   "New"
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   10140
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Ledger"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   465
            Index           =   8
            Left            =   30
            TabIndex        =   9
            Top             =   135
            Width           =   10035
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Ledger"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   150
            Width           =   10035
         End
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   975
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         ListBackColor   =   16777215
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBackColor   =   16777215
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
      Begin MSFlexGridLib.MSFlexGrid FGrid 
         Height          =   4950
         Left            =   120
         TabIndex        =   2
         Top             =   1425
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   8731
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777215
         ForeColorFixed  =   12582912
         BackColorBkg    =   -2147483633
         GridColor       =   6244673
         GridColorFixed  =   6244673
         GridLinesFixed  =   1
         Appearance      =   0
         FormatString    =   "< Cust Code  |^  Country  |^      Invoice #       |^          Date          |^    Amount   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   300
         Left            =   5640
         TabIndex        =   15
         Top             =   975
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22740995
         CurrentDate     =   37770
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   300
         Left            =   7200
         TabIndex        =   16
         Top             =   975
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22740995
         CurrentDate     =   37770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5640
         TabIndex        =   18
         Top             =   780
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   7200
         TabIndex        =   17
         Top             =   780
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2025
         TabIndex        =   6
         Top             =   780
         Width           =   585
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   300
         Left            =   2010
         TabIndex        =   5
         Top             =   975
         Width           =   1635
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "2884;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice #"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3705
         TabIndex        =   4
         Top             =   765
         Width           =   690
      End
      Begin MSForms.ComboBox cmbInv 
         Height          =   300
         Left            =   3675
         TabIndex        =   3
         Top             =   975
         Width           =   1950
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "3440;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmPaymentsLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbBy_Change()
    Call cmbBy_Click
End Sub

Private Sub cmbBy_Click()

    If Not cmbBy.MatchFound Then Exit Sub
    
    Dim Sql As String
    If cmbBy.Text = "Customer" Then
        Sql = "Distinct CustCode"
    Else
        Sql = "Distinct Curr"
    End If
    
    Call cmbCust.ClearVals
    Call cmbCust.AddItem("<All>", "0")
    Call cmbCust.AddVals(con, Sql, "Customers")
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    On Error GoTo err
    
    
    Call ClearRows(FGrid)
    If cmbCountry.MatchFound = False Then Exit Sub
    'cmbCust.Tag = GetAutoCustID(cmbCust.Text & cmbCountry)
    
    Dim iRow As Integer
    Dim Cond As String
    Dim rs As New ADODB.Recordset
    
    If cmbCountry.ListIndex = 0 Then
        If cmbCust.ID = "" Then
            Cond = ""
        Else
            Cond = " Where CustCode='" & cmbCust.Text & "'"
        End If
    Else
        If cmbCust.ID = "" Then
            Cond = " Where Country='" & cmbCountry.Text & "'"
        Else
            Cond = " Where CustCode='" & cmbCust.Text & "' and Country='" & cmbCountry.Text & "'"
        End If
    End If
    
    
    With rs
        
        .Open "Select Distinct CustomInvoice From VPaymentRcvd" & Cond & " ", con, adOpenForwardOnly, adLockReadOnly
        cmbInv.Clear
        cmbInv.AddItem "<All Invoices>"
        Do Until .EOF
            cmbInv.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbCust_matched()
    On Error GoTo err
    
    'Call FillCountries(cmbCust.Text, cmbCountry)
    Dim Sql As String
    If cmbCust.ID = "" Then
        Sql = " Where CustCode='" & cmbCust.Text & "'"
    Else
        Sql = ""
    End If
        
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Distinct Country from ForeignCustomers " & Sql, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
    End With
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
        
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmdAdd_Click()
    frmNewPayment.Show 1
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdPrint_Click()

    On Error GoTo err
    Exit Sub
    con.Execute "Delete From rptTotalExport"
    With FGrid
        For i = 1 To .Rows - 1
            'con.Execute
        Next i
    End With
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmdRefresh_Click()
    On Error GoTo err
    Call ClearRows(FGrid)
    If cmbCountry.MatchFound = False Then Exit Sub
    'cmbCust.Tag = GetAutoCustID(cmbCust.Text & cmbCountry)
    
    Dim iRow As Integer
    Dim Cond As String
    Dim rs As New ADODB.Recordset
    
    If cmbCountry.ListIndex = 0 Then
        If cmbCust.ID = "" Then
            Cond = ""
        Else
            Cond = " Where CustCode='" & cmbCust.Text & "'"
        End If
    Else
        If cmbCust.ID = "" Then
            Cond = " Where Country='" & cmbCountry.Text & "'"
        Else
            Cond = " Where CustCode='" & cmbCust.Text & "' and Country='" & cmbCountry.Text & "'"
        End If
    End If
    
    If cmbInv.ListIndex <> 0 Then
        If Cond = "" Then
            Cond = " Where CustomInvoice='" & cmbInv & "'"
        Else
            Cond = " and CustomInvoice='" & cmbInv & "'"
        End If
    End If
    
    If Cond = "" Then
        Cond = Cond & " Where DT Between #" & DtFrom & "# and #" & DtTo & "#"
    Else
        Cond = Cond & " and DT Between #" & DtFrom & "# and #" & DtTo & "#"
    End If
    With rs
        .Open "Select * From VPaymentRcvd" & Cond & " Order By DT,EntryID", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
        
            iRow = .AbsolutePosition
            If iRow > 1 Then FGrid.AddItem ""
            FGrid.TextMatrix(iRow, 0) = ![CustCode] & ""
            FGrid.TextMatrix(iRow, 1) = ![Country] & ""
            
            FGrid.TextMatrix(iRow, 2) = ![CustomInvoice] & ""
            FGrid.TextMatrix(iRow, 3) = Format(![DT], "dd-MM-yyyy") & ""
            FGrid.TextMatrix(iRow, 4) = ![Amt] & " " & ![Curr] & ""
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    
    DtFrom = DateAdd("m", -1, Date)
    DtTo = Date
    
    With FGrid
    
        .ColWidth(0) = 1400
        .ColWidth(1) = 1400
        .ColWidth(2) = 1200
        .ColWidth(3) = 1400
        .ColWidth(4) = 1200
        
    
    End With
    
    cmbCust.ListHeight = 1500
    'cmbCat.ListHeight = 1500
    cmbCust.AddItem "<All>", 0
    cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers"
    'Call GetItems
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub
