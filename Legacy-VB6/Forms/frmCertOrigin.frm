VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCertOrigin 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate Of Origin"
   ClientHeight    =   7170
   ClientLeft      =   4545
   ClientTop       =   2640
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8955
   Begin VB.TextBox txtEntryID 
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   8700
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Certificate Of Origin"
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
         Height          =   390
         Index           =   8
         Left            =   60
         TabIndex        =   17
         Top             =   165
         Width           =   8775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Certificate Of Origin"
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
         Index           =   9
         Left            =   90
         TabIndex        =   18
         Top             =   180
         Width           =   8760
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   5865
      Left            =   30
      TabIndex        =   13
      Top             =   1200
      Width           =   7320
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   44
         Top             =   5400
         Width           =   3495
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   43
         Top             =   4620
         Width           =   3495
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   1110
         Index           =   8
         Left            =   3720
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   4620
         Width           =   3495
      End
      Begin VB.ComboBox cmbAddress 
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "frmCertOrigin.frx":0000
         Left            =   4530
         List            =   "frmCertOrigin.frx":000A
         TabIndex        =   38
         Text            =   "Combo2"
         Top             =   1605
         Width           =   1875
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   555
         Index           =   7
         Left            =   2085
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   3840
         Width           =   1590
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   1110
         Index           =   5
         Left            =   3705
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   3285
         Width           =   3495
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   32
         Top             =   4095
         Width           =   1860
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   315
         Index           =   2
         Left            =   2085
         TabIndex        =   30
         Top             =   3300
         Width           =   1590
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   315
         Index           =   6
         Left            =   180
         TabIndex        =   28
         Top             =   3300
         Width           =   1860
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   1110
         Index           =   0
         Left            =   3705
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   465
         Width           =   3495
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Combo2"
         Top             =   2745
         Width           =   1875
      End
      Begin VB.ComboBox cmbCust 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1965
         Width           =   1875
      End
      Begin VB.ComboBox cmbCountry 
         Height          =   315
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1965
         Width           =   1590
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   1110
         Index           =   1
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   465
         Width           =   3495
      End
      Begin VB.TextBox TBox 
         DataSource      =   "Data1"
         Height          =   1110
         Index           =   3
         Left            =   3705
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1935
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   315
         Left            =   2085
         TabIndex        =   26
         Top             =   2745
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   22347779
         CurrentDate     =   37924
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Designation"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   42
         Top             =   5160
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "User Name"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   41
         Top             =   4440
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Description of goods"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   40
         Top             =   4440
         Width           =   1470
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark n Numbers"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   2085
         TabIndex        =   37
         Top             =   3645
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Number and Kind of Packages"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3705
         TabIndex        =   35
         Top             =   3105
         Width           =   2130
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Weight"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   3885
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Weight"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   2070
         TabIndex        =   31
         Top             =   3090
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Value"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   210
         TabIndex        =   29
         Top             =   3105
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   2115
         TabIndex        =   27
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Via"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3705
         TabIndex        =   25
         Top             =   270
         Width           =   210
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Consignee"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3720
         TabIndex        =   21
         Top             =   1635
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Country"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   2115
         TabIndex        =   20
         Top             =   1740
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Cust Code"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   1725
         Width           =   750
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   2520
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Name && Address Of Exporter"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   255
         Width           =   2085
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000004&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4410
      Width           =   1335
   End
   Begin VB.CommandButton clos 
      BackColor       =   &H80000004&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3930
      Width           =   1335
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H80000004&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3090
      Width           =   1335
   End
   Begin VB.CommandButton ok 
      BackColor       =   &H80000004&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2595
      Width           =   1335
   End
   Begin VB.CommandButton edit 
      BackColor       =   &H80000004&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3330
      Width           =   1335
   End
   Begin VB.CommandButton delete 
      BackColor       =   &H80000004&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2850
      Width           =   1335
   End
   Begin VB.CommandButton add 
      BackColor       =   &H80000004&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2370
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   3690
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Entry No :"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   2895
      TabIndex        =   12
      Top             =   885
      Width           =   735
   End
End
Attribute VB_Name = "frmCertOrigin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private customersel As String
Dim IsEdit As Boolean
Private Sub add_Click()

    On Error GoTo eh
    Dim NewSNo As String
    
    
    
    'getshipping

    'Data1.Recordset.AddNew
    TBox(0) = ""
    TBox(4) = ""
    TBox(5) = ""
    TBox(6) = ""
    TBox(7) = ""
    TBox(2) = ""
    
    vis (False)
    loc (False)
    
    
    'Set rec = data.OpenRecordset("Select Max(Val(Mid(SNo,Instr(SNo,'-')+1))) From SepAdvancePayments")
    NewSNo = "CO-"
    'NewSNo = NewSNo & Format(Val(rec.Fields(0) & "") + 1, "0000")
    TBox(9).Text = "NAVEED"
    TBox(10).Text = "PARTNER"
    
    With txtEntryID
        .Text = NewSNo
        .SetFocus
        .SelStart = 3
        '.SelLength = 0
        
    End With
    
    Exit Sub

eh:
    MsgBox err.Description

End Sub

Private Sub cancel_Click()

    On Error GoTo eh
    'Data1.Recordset.CancelUpdate



    vis (True)
    loc (True)
    Exit Sub
    
eh:
    MsgBox err.Description

End Sub

Private Sub clos_Click()
    Unload Me
End Sub

Private Sub cmbAddress_Change()
    Call cmbAddress_Click
End Sub

Private Sub cmbAddress_Click()
    If cmbAddress.ListIndex = -1 Then Exit Sub
    If Not add.Visible Then
        If cmbAddress.ListIndex = 0 Then
            TBox(3) = cmbAddress.Tag
        ElseIf cmbAddress.ListIndex = 1 Then
            TBox(3) = TBox(3).Tag
        End If
    End If
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    On Error GoTo err
    
    If cmbCust.ListIndex = -1 Or cmbCountry.ListIndex = -1 Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
    
        .Open "Select Address From ForeignCustomers Where CustCode='" & cmbCust & "' And Country='" & cmbCountry & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            TBox(3).Text = .Fields(0) & ""
        End If
        .Close
        
        .Open "Select CustomInvoice From CustomInvoice Where CustCode='" & cmbCust & "' And Country='" & cmbCountry & "' And CustomInvoice Not In(Select CustomInvoice From CertificateOfOrigin)", con, adOpenForwardOnly, adLockReadOnly
        Combo2.Clear
        Do Until .EOF
            Combo2.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
        
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Combo1_Click()

    On Error GoTo eh
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT  * from VCertificateOfOrigin Where SNo='" & Combo1.Text & "'", con, adOpenForwardOnly, adLockReadOnly
    
        If .RecordCount > 0 Then
            TBox(0) = ![Via] & ""
            cmbCust = ![CustCode] & ""
            cmbCountry = ![Country] & ""
            Combo2.Text = ![CustomInvoice] & ""
            
            TBox(2) = ![GWeight] & ""
            TBox(4) = ![NetWeight] & ""
            TBox(7) = ![MarknNumber] & ""
            TBox(5) = ![NonKind] & ""
            
            TBox(6) = Val(![TotalAmt] & "")
            TBox(3) = ![Address] & ""
            
            TBox(8) = ![DescriptionOfGoods] & ""
            TBox(9) = ![UserName] & ""
            TBox(10) = ![UserDesignation] & ""
                        
            .Close
            .Open "Select ComNotify,ComConsignee From CustomInvoice Where CustomInvoice='" & Combo2.Text & "'"
    
            If .RecordCount > 0 Then
                cmbAddress.Tag = ![ComNotify] & ""
                TBox(3).Tag = ![ComConsignee] & ""
            End If
            
        End If
        .Close
        
    End With
    
    txtEntryID.Text = Combo1.Text
    
    Exit Sub
eh:
    MsgBox err.Description
 
End Sub

Private Sub Combo2_Change()
    Call Combo2_click
End Sub

Private Sub Combo2_click()
    
    On Error GoTo err
    
    
    If Combo2.ListIndex = -1 Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select DT,ComNotify,ComConsignee From CustomInvoice Where CustomInvoice='" & Combo2.Text & "'", con, adOpenForwardOnly, adLockReadOnly
    
        If .RecordCount > 0 Then
            DT1 = .Fields(0) & ""
            cmbAddress.Tag = ![ComNotify] & ""
            TBox(3).Tag = ![ComConsignee] & ""
        End If
    
        If add.Visible Then Exit Sub
    
        'txtEntryID = Replace(Combo2.Text, "VTL ", "CO-")
        txtEntryID = "CO-" & Mid(Combo2.Text, 4)
        .Close
        .Open "Select TotalAmt From VCustomInvoiceAmt Where CustomInvoice='" & Combo2.Text & "'", con, adOpenForwardOnly, adLockReadOnly
    
        If .RecordCount > 0 Then
            
            TBox(6) = .Fields(0) & ""
        End If
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtEntryID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If Len(txtEntryID.Text) = 3 Then KeyAscii = 0
    End If
End Sub

Private Sub Command2_Click()

    On Error GoTo eh
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptCertOfOrigin.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VCertificateOfOrigin.SNo}='" & Combo1.Text & "' ", rpt
    
    Exit Sub
    
eh:
     MsgBox err.Description
     
End Sub

Private Sub edit_Click()

    On Error GoTo eh
    IsEdit = True
    

    vis (False)
    loc (False)
    Exit Sub
eh:
    MsgBox err.Description

End Sub

Private Sub Form_Load()


    On Error GoTo eh
    IsEdit = False
    
    Call FillCmbs
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select SNo from CertificateOfOrigin", con, adOpenForwardOnly, adLockReadOnly
     
        Combo1.Clear
        While Not .EOF
            Combo1.AddItem .Fields(0).value
            .MoveNext
        Wend
        .Close
    End With
    Set rs = Nothing
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
     
    loc (True)
    vis (True)

    Exit Sub
eh:
    MsgBox err.Description
End Sub
Private Sub cmbCust_Change()
    Call cmbCust_Click
End Sub

Private Sub cmbCust_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    
    If cmbCust.ListIndex = 0 Then
        StrSQL = "select distinct(Country) from ForeignCustomers"
    Else
        StrSQL = "select distinct(Country) from ForeignCustomers Where CustCode='" & cmbCust.Text & "'"
    End If
    
    With rs
        .Open StrSQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
    
        Do Until .EOF
            cmbCountry.AddItem .Fields(0).value & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    cmbCountry.ListIndex = 0
        
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub FillCmbs()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select distinct(custcode) from ForeignCustomers", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCust.Clear
    
        Do Until .EOF
            cmbCust.AddItem .Fields(0).value
            .MoveNext
        Loop
        .Close
    End With
    
    cmbCust.ListIndex = 0
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub ok_Click()

    On Error GoTo eh
    Dim rs As New ADODB.Recordset
    
    If txtEntryID.Text = "CO-" Or Left(txtEntryID.Text, 3) <> "CO-" Then
        MsgBox "Invalid ID", vbInformation
        Exit Sub
    End If
    
    If IsEdit Then
        con.Execute "Delete From CertificateOfOrigin Where SNo='" & Combo1.Text & "'"
    End If
    con.Execute "Insert Into CertificateOfOrigin(SNo,CustomInvoice,CustCode,Country,GWeight,NetWeight,MarknNumber,NonKind,Via,Address," & _
    "DescriptionOfGoods,UserName,UserDeSignation) Values('" & _
     txtEntryID.Text & "','" & Combo2.Text & "','" & cmbCust & "','" & cmbCountry & "'," & _
     Val(TBox(2)) & "," & Val(TBox(4)) & ",'" & TBox(7) & "','" & TBox(5) & "','" & TBox(0) & "','" & TBox(3) & _
     "','" & TBox(8) & "','" & TBox(9) & "','" & Replace(TBox(10), "'", "''") & "')"

    With rs
        .Open "Select SNo From CertificateOfOrigin", con, adOpenForwardOnly, adLockReadOnly
        
        Combo1.Clear
        While Not .EOF
            Combo1.AddItem .Fields(0).value
            .MoveNext
        Wend
        .Close
    End With
    
    Set rs = Nothing
    
    vis (True)
    loc (True)

    Exit Sub
eh:

    MsgBox err.Description
    
End Sub

Private Sub vis(bool As Boolean)

    clos.Visible = bool
    add.Visible = bool
    delete.Visible = bool
    edit.Visible = bool
    ok.Visible = Not bool
    cancel.Visible = Not bool
    Combo1.Visible = bool
    txtEntryID.Visible = Not bool
    
End Sub

Private Sub loc(bool As Boolean)

    On Error Resume Next
    
    For a = 0 To TBox.count - 1
        TBox(a).Locked = bool
    Next a
    
    Combo2.Locked = bool
    
    'If IsEdit = False And bool Then
    '    Combo1.RemoveItem Combo1.ListCount - 1
    'End If
    
    If bool Then
        IsEdit = False
    End If

End Sub

Sub getshipping()


Set rec = Data.OpenRecordset("select name from shipping ")
  
  Combo4.Clear
  While Not rec.EOF
  Combo4.AddItem rec.Fields(0)
  rec.MoveNext
  Wend
  
 Set rec = Nothing
 
End Sub




