VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTestStats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "THTC Test Statistics"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport cr1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   150
      TabIndex        =   16
      Top             =   -15
      Width           =   5775
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Statistics"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Index           =   21
         Left            =   2025
         TabIndex        =   17
         Top             =   150
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Statistics"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   22
         Left            =   2040
         TabIndex        =   18
         Top             =   165
         Width           =   1845
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   2145
      Left            =   150
      TabIndex        =   0
      Top             =   510
      Width           =   5775
      Begin VB.ComboBox cmbRange 
         Height          =   315
         ItemData        =   "frmTestStats.frx":0000
         Left            =   90
         List            =   "frmTestStats.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   465
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   4110
         TabIndex        =   2
         Top             =   465
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22806531
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   2430
         TabIndex        =   3
         Top             =   465
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22806531
         CurrentDate     =   37055
      End
      Begin MSForms.CommandButton CommandButton2 
         Default         =   -1  'True
         Height          =   375
         Left            =   4110
         TabIndex        =   15
         Top             =   1665
         Width           =   1545
         Caption         =   "Close"
         PicturePosition =   327683
         Size            =   "2725;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   4110
         TabIndex        =   14
         Top             =   1260
         Width           =   1545
         Caption         =   "Print"
         PicturePosition =   327683
         Size            =   "2725;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2715
         TabIndex        =   13
         Top             =   870
         Width           =   1320
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   375
         Left            =   4110
         TabIndex        =   12
         Top             =   855
         Width           =   1545
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2725;661"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2715
         TabIndex        =   11
         Top             =   1770
         Width           =   1320
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2715
         TabIndex        =   10
         Top             =   1170
         Width           =   1320
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2715
         TabIndex        =   9
         Top             =   1470
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total  No Of CardiacT's Done : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   5
         Left            =   105
         TabIndex        =   8
         Top             =   1770
         Width           =   2580
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total  No Of ETT's Done          : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   4
         Left            =   105
         TabIndex        =   7
         Top             =   1470
         Width           =   2580
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total  No Of ECHo's Done        : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   3
         Left            =   105
         TabIndex        =   6
         Top             =   1170
         Width           =   2580
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total  No Of ECG's Done          : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   870
         Width           =   2580
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Date Range                              From                           To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   5565
      End
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Test Statistics"
      End
      Begin VB.Menu mnuPrintMonthly 
         Caption         =   "Print Monthly Test Statistics"
      End
   End
End
Attribute VB_Name = "frmTestStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblECG_Click(Index As Integer)

End Sub


Private Sub cmbRange_Click()

    DtTo = Date
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
    Case 0
        DtFrom = Date
    Case 1
        DtFrom = DateAdd("d", -7, Date)
    Case 2
        DtFrom = DateAdd("d", -15, Date)
    Case 3
        DtFrom = DateAdd("d", -30, Date)
    Case 4
        DtFrom.Enabled = True
        DtTo.Enabled = True
        DtFrom.SetFocus
        Exit Sub
    End Select
    Call cmdRefresh_Click


End Sub





Private Sub cmdPrint_Click()
    PopupMenu mnupop, , FAcc.Left + cmdPrint.Left, FAcc.Top + cmdPrint.Top + cmdPrint.Height
End Sub

Private Sub cmdRefresh_Click()
Dim rs As New ADODB.Recordset

With rs
    .Open "Select count(ECG) as TotalECGs,count(ECHO) as TotalECHOs," & _
     "count(ETT) as TotalETTs,count(CardiacT) as TotalCardiacTs From DailyIncomesheet Where DT Between '" & _
     Format(DtFrom, "dd-MMM-yyyy") & "' and '" & Format(DtTo, "dd-MMM-yyyy") & "'", con, adOpenForwardOnly, adLockReadOnly
     
     
    lbls(0) = Val(!TotalECGS & "")
    lbls(1) = Val(!TotalEchoS & "")
    lbls(2) = Val(!TotalETTs & "")
    lbls(3) = Val(!TotalCardiacTs & "")
    
    .Close
    
End With

Set rs = Nothing
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    mnupop.Visible = False
    cmbRange.ListIndex = 0
End Sub

Private Sub mnuPrint_Click()
    With cr1
        .ReportFileName = RptPath & "\teststats.rpt"
        .Formulas(0) = "ECGS=" & lbls(0)
        .Formulas(1) = "ECHOS=" & lbls(1)
        .Formulas(2) = "ETTS=" & lbls(2)
        .Formulas(3) = "CardiacTS=" & lbls(3)
        .Formulas(4) = "fromto='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
        .Action = 1
        .PageZoomNext
    End With
End Sub

Private Sub mnuPrintMonthly_Click()
    con.Execute "Delete From Dates "
    For i = 1 To DateDiff("M", DtFrom, DtTo)
        con.Execute "Insert into Dates (DT,SortDT) values('" & Format(DateAdd("M", i, DtFrom), "MMM yyyy") & "','" & DateAdd("m", i, DtFrom) & "')"
    Next
    
    With cr1
        .ReportFileName = RptPath & "\monthlytests.rpt"
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
        .Formulas(4) = "fromto='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
        '.SelectionFormula = "{dailyincomesheet.Dt}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") To date(" & Format(DtTo, "yyyy,MM,dd") & ")"
        .Action = 1
        .PageZoomNext
    End With
    
    
End Sub
