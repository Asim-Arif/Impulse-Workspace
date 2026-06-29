VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendAdvSalary 
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkExternal 
      Caption         =   "External Salary Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9750
      TabIndex        =   12
      Top             =   885
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   15
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Payment"
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
         Index           =   1
         Left            =   4680
         TabIndex        =   9
         Top             =   120
         Width           =   2760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Payment"
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
         Height          =   390
         Index           =   2
         Left            =   4695
         TabIndex        =   10
         Top             =   135
         Width           =   2760
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Top             =   870
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   330
      Left            =   2265
      TabIndex        =   5
      Top             =   870
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Frame FLV 
      Height          =   6255
      Left            =   45
      TabIndex        =   0
      Top             =   1230
      Width           =   11835
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   900
         Left            =   45
         TabIndex        =   19
         Top             =   5340
         Width           =   4935
         Begin VB.CommandButton cmdAccounts 
            Caption         =   "Pay All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2310
            TabIndex        =   21
            Top             =   450
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Unpay All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3615
            TabIndex        =   20
            Top             =   450
            Visible         =   0   'False
            Width           =   1275
         End
         Begin MSForms.CommandButton cmdUpdate 
            Height          =   360
            Left            =   15
            TabIndex        =   22
            Top             =   450
            Width           =   2265
            ForeColor       =   0
            Caption         =   "Make This Sheet Final"
            PicturePosition =   327683
            Size            =   "3995;635"
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   900
         Left            =   5445
         TabIndex        =   14
         Top             =   5325
         Width           =   6345
         Begin MSForms.CommandButton cmdAddDetail 
            Height          =   360
            Left            =   2565
            TabIndex        =   27
            Top             =   495
            Visible         =   0   'False
            Width           =   1215
            ForeColor       =   0
            Caption         =   "Add Detail"
            PicturePosition =   327683
            Size            =   "2143;635"
            Accelerator     =   68
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total Credit :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   26
            Top             =   405
            Width           =   1080
         End
         Begin VB.Label lblCredit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1350
            TabIndex        =   25
            Top             =   390
            Width           =   1170
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total Debit :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   255
            TabIndex        =   24
            Top             =   90
            Width           =   1020
         End
         Begin VB.Label lblDebit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1350
            TabIndex        =   23
            Top             =   75
            Width           =   1170
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            Height          =   285
            Left            =   4830
            TabIndex        =   16
            Top             =   75
            Width           =   1425
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Salary Amount :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2865
            TabIndex        =   15
            Top             =   90
            Width           =   1875
         End
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   5055
            TabIndex        =   18
            Top             =   495
            Width           =   1215
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "2143;635"
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
            Height          =   360
            Left            =   3810
            TabIndex        =   17
            Top             =   495
            Width           =   1215
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2143;635"
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   5070
         TabIndex        =   11
         Top             =   690
         Visible         =   0   'False
         Width           =   1155
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5130
         Left            =   60
         TabIndex        =   1
         Top             =   195
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   9049
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Vend ID"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vender Name"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Work Done"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "ST Loan"
            Object.Width           =   1508
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Loan Ded."
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Res Ded."
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Balance"
            Object.Width           =   1667
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Per%"
            Object.Width           =   1111
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Payable"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Db/Cr"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendAdvSalary.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendAdvSalary.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8055
      TabIndex        =   2
      Top             =   855
      Width           =   1605
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2831;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbDays 
      Height          =   300
      Left            =   7050
      TabIndex        =   13
      Top             =   900
      Width           =   990
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1746;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   5685
      TabIndex        =   7
      Top             =   885
      Width           =   1320
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2328;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   4530
      TabIndex        =   6
      Top             =   885
      Width           =   1110
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1958;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   615
      Width           =   7935
      BackColor       =   11517387
      Caption         =   "  Department                         Venders                                 Year                  Month                Total Days"
      Size            =   "13996;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuSalSheet 
         Caption         =   "Print Salary Sheet"
      End
      Begin VB.Menu mnuPrintChqs 
         Caption         =   "Print Cheques"
      End
      Begin VB.Menu mnuPrintVouchers 
         Caption         =   "Print Vouchers"
      End
   End
End
Attribute VB_Name = "frmVendAdvSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ColIndex As Integer
Dim rsLedger As New ADODB.Recordset
Dim bSavedSal As Boolean

Private Sub chkExternal_Click()
    Call cmdRefresh_Click
End Sub

Private Sub cmbMonth_Click()

    Dim maxday As Integer
    cmbDays = TotalMonthDays
    If Format(ServerDate, "MMM-yyyy") = cmbMonth & "-" & cmbYear Then
        maxday = Day(ServerDate)
    Else
        maxday = TotalMonthDays
    End If
    cmbDays.Clear
    
    For i = 1 To maxday
        cmbDays.AddItem Format(i, "00")
    Next
    
    cmbDays.ListIndex = maxday - 1
    
End Sub

Private Sub cmbYear_Change()
    txtTotal = TotalMonthDays
End Sub

Private Sub cmdAccounts_Click()
For i = 1 To LV.ListItems.Count
    LV.ListItems(i).Checked = True
Next

End Sub

Private Sub cmdCancel_Click()
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End Sub

Private Sub cmdAddDetail_Click()

    On Error GoTo err
    Dim F As New frmVendAddPayDetail
    F.ShowMe CInt(cmbMonth.ListIndex + 1), CInt(cmbYear), 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub PrintReport(PaySlips As Boolean)

    On Error GoTo err
    If LV.ListItems.Count = 0 Then Exit Sub
    If Not bSavedSal Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Dim DtFrom As Date
    Dim DtTo As Date
    DtFrom = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    DtTo = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), cmbDays)
    
    Dim rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendAdvSalSheet.rpt")
    'rpt.RecordSelectionFormula =
    
    Dim F As New frmPrevRpt
    
    F.ShowReport "{VVendMonthlyPayments.DT}=#" & DtTo & "# AND {VVendMonthlyPayments.PaidAmt}>0 AND {VVendMonthlyPayments.AdvSalary}=True AND {VVendMonthlyPayments.Payed}=True ", rpt
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub PrintReportOLD(PaySlips As Boolean)

    On Error GoTo err
    If LV.ListItems.Count = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    
    Dim i As Integer
    Dim j As Integer
    con.Execute "Delete from PrintSalary"
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select * from PrintSalary", con, adOpenStatic, adLockOptimistic
        For i = 1 To LV.ListItems.Count
         If LV.Checkboxes And LV.ListItems(i).Checked Then
            .AddNew
            ![VendID] = Val(LV.ListItems(i).Key)
            ![VendID1] = LV.ListItems(i).Text & ""
            
            For j = 1 To LV.ColumnHeaders.Count - 1
                
                'If j = 19 Then
                '    .Fields(j) = LV.ListItems(i).SubItems(j + 1) & ""
                'ElseIf j = 20 Then
                '    .Fields(j) = LV.ListItems(i).SubItems(j - 1) & ""
                'Else
                '    .Fields(j) = LV.ListItems(i).SubItems(j) & ""
                'End If
                Debug.Print .Fields(j).Name
                .Fields(j + 1) = LV.ListItems(i).SubItems(j) & ""
            Next j
            ![DT] = DT & ""
            .Update
         End If
        Next i
    End With
    
    'Sleep (1000)
    'Dim TDate As Date
    'Dim MonthDays As Integer
    'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    'MonthDays = DateAdd("m", 1, TDate) - TDate
    
    'With cr1
    '
    '    If Not PaySlips Then
    '      .ReportFileName = RptPRPath & "\AdvSalSheet.rpt"
    '      '.DataFiles(0) = DatabasePath
    '      .Connect = ConnectStr
    '      .Formulas(0) = "Fromto=' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
    '    Else
    '      .ReportFileName = RptPRPath & "\PaySlip.rpt"
    '      '.DataFiles(0) = DatabasePath
    '      .Connect = ConnectStr
    '      .Formulas(0) = ""
    '    End If
    '
    '    '.Formulas(1) = "ForMonthDays=" & MonthDays
    '    .Action = 1
    '    .PageZoomNext
    'End With
    Exit Sub
    'Dim F As New frmPrevRpt
    'Dim rpt As New rptAdvSalSheet
    'F.ShowReport " ", rpt
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
    'Me.PopupMenu mnupop, 8, cmdPrint.Left, cmdPrint.Top + FLV.Top + cmdPrint.Height
    'Call PrintReport(False)
    Me.PopUpMenu mnupop, 8, fButs1.Left + cmdPrint.Left, cmdPrint.Top + fButs1.Top + cmdPrint.Height
End Sub

Public Sub cmdRefresh_Click()

    
    'Check If salary Has been Saved or Not
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Count(*) From VendMonthlyPayments Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear & " AND AdvSalary=1", con, adOpenForwardOnly, adLockReadOnly
        If .Fields(0).value > 0 Then
            bSavedSal = True
            LV.Checkboxes = True
            cmdAccounts.Visible = True
            Command1.Visible = True
            cmdUpdate.Caption = "Save Changes"
            cmdAddDetail.Visible = True
        Else
            bSavedSal = False
            LV.Checkboxes = False
            cmdAccounts.Visible = False
            Command1.Visible = False
            cmdUpdate.Caption = "Make This Sheet Final"
            cmdAddDetail.Visible = False
        End If
        .Close
        
        If Not bSavedSal Then
            .Open "Select Count(*) From VVendRcvdItems Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear & " AND Day(DT)<=" & cmbDays & " And ReqAuth=1", con, adOpenForwardOnly, adLockReadOnly
            If .Fields(0).value > 0 Then
                MsgBox "There Are Some Entries Waiting For Authorization." & vbNewLine & "You Should Do Something With Them Before Continuing.", vbInformation
                .Close
                Exit Sub
            End If
            .Close
        End If
        
        
    End With
    
    Me.MousePointer = vbHourglass
    Call RefreshLedger
     
    'cmdAccounts.Visible = cmdUpdate.Visible And AccountsLinked
    Me.MousePointer = vbNormal
    'Call Form_Resize
    
End Sub



Private Sub cmdSave_Click()
    'Call SaveIt(rsLedger)
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End Sub

Private Sub cmdUpdate_Click()
    
    On Error GoTo err
    
    If cmdUpdate.Caption = "Make This Sheet Final" Then

        With rsLedger
            If .RecordCount = 0 Then Exit Sub Else .MoveFirst
           
            con.BeginTrans
           
            For i = 1 To LV.ListItems.Count
                .Find "VendID=" & Val(LV.ListItems(i).Key), 0, adSearchForward, 1
                Call SaveIt(rsLedger, 0, LV.ListItems(i).Checked, Val(LV.ListItems(i).SubItems(8)), Val(LV.ListItems(i).SubItems(7)))
           
                .MoveNext
            Next
           
            con.CommitTrans
            Call cmdRefresh_Click
        End With
    
        
   
    Else
    
        Dim setVal As Integer
        Dim VendID As Long
        Dim TheMonth As Byte
        Dim TheYear As Integer
        Dim MonthInd As Integer
        
        con.BeginTrans
        
        For i = 1 To LV.ListItems.Count
            setVal = Abs(LV.ListItems(i).Checked)
            VendID = Val(LV.ListItems(i).Key)
            MonthInd = InStr(LV.ListItems(i).Key, "|") + 1
            TheMonth = Month(CDate(Right(LV.ListItems(i).Key, Len(LV.ListItems(i).Key) - MonthInd)))
            'TheMonth = Month(CDate(Right(LV.ListItems(i).Key, Len(LV.ListItems(i).Key) - 7)))
            TheYear = year(CDate(Right(LV.ListItems(i).Key, Len(LV.ListItems(i).Key) - MonthInd)))
'            If Val(LV.ListItems(i).SubItems(3)) > 0 Then
'                con.Execute "Update VendMonthlyPayments Set STAmt=" & Val(LV.ListItems(i).SubItems(3)) & " Where VendID=" & VendID & " and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear & " and AdvSalary=1"
'            End If
            con.Execute "Update VendMonthlyPayments Set Payed=" & setVal & ",CashPaid=" & _
             IIf(LV.ListItems(i).SubItems(10) = "Cash", 1, 0) & ",ChequeNo='" & _
             LV.ListItems(i).SubItems(11) & "' Where VendID=" & VendID & " and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear & " and AdvSalary=1", a
             
        Next i
        
        con.CommitTrans
    
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Command1_Click()
For i = 1 To LV.ListItems.Count
    LV.ListItems(i).Checked = False
Next
End Sub

Private Sub Form_Load()

'cmdUpdate.Visible = UserHasAccess("FinalAdvSal")
 
 'Hide Columns Which are Not required
    mnupop.Visible = False
    DtTo = ServerDate
    DtFrom = DateAdd("M", -1, ServerDate)
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    cmbMonth.AddItem "Jan"
    cmbMonth.List(0, 1) = 1
    cmbMonth.AddItem "Feb"
    cmbMonth.List(1, 1) = 2
    cmbMonth.AddItem "Mar"
    cmbMonth.List(2, 1) = 3
    cmbMonth.AddItem "Apr"
    cmbMonth.List(3, 1) = 4
    cmbMonth.AddItem "May"
    cmbMonth.List(4, 1) = 5
    cmbMonth.AddItem "Jun"
    cmbMonth.List(5, 1) = 6
    cmbMonth.AddItem "Jul"
    cmbMonth.List(6, 1) = 7
    cmbMonth.AddItem "Aug"
    cmbMonth.List(7, 1) = 8
    cmbMonth.AddItem "Sep"
    cmbMonth.List(8, 1) = 9
    cmbMonth.AddItem "Oct"
    cmbMonth.List(9, 1) = "10"
    cmbMonth.AddItem "Nov"
    cmbMonth.List(10, 1) = 11
    cmbMonth.AddItem "Dec"
    cmbMonth.List(11, 1) = 12
    
    Dim CYear As Long
    CYear = Format(ServerDate, "yyyy")
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Processes>", "0"
    cmbDept.AddVals con, "Description", "Processes", "ProcessID"
    cmbDept.Text = "<All Processes>"
    
    
    'Set LV.SmallIcons = ImageList1
    cmbMonth.ListIndex = Val(Format$(ServerDate, "MM")) - 1
    cmbYear = Format$(ServerDate, "yyyy")
    
    
    'Call cmdRefresh_Click
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Venders>", "0"
    If cmbDept.Text = "<All Processes>" Then
        cmbEmp.AddVals con, "VenderName", "Venders", "VendID"
    Else
        cmbEmp.AddVals con, "VenderName", "Venders", "VendID", "Where PhaseID=" & cmbDept.ID
    End If
    cmbEmp.Text = "<All Venders>"
        
End Sub

Private Sub Form_Resize()

On Error Resume Next

 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

 FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 fButs2.Move LV.Left, LV.Top + LV.Height

 
End Sub



Private Sub RefreshLedger()

    Dim GrandTotal As Double
    Dim TotalDebit As Double
    Dim TotalCredit As Double
    
    Dim dAdvSalPer As Double
    
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    If cmbYear.MatchFound = False Or cmbMonth.MatchFound = False Then Exit Sub
    
    Dim VendID As Long
    Dim ProcessID As Long
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    'Dim rsLedger As New ADODB.Recordset
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            ProcessID = 0
        Else
            ProcessID = cmbDept.ID
        End If
        'VendID = IIf(cmbDept.ID = "0", "", cmbDept.ID)
    Else
       VendID = cmbEmp.ID
    End If
    
    Dim rsCheck As New ADODB.Recordset
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(ServerDate))
    
    SDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(ServerDate))
    
    'TotalSalDays = DateAdd("m", 1, SDate) - SDate
    
    rs.Open "Select Distinct(DT) from VendMonthlyPayments Where convert(char(3),dt,7)='" & cmbMonth & "' and Year(DT)=" & cmbYear & " and AdvSalary=1", con
    Dim AdvSalLocked As Boolean
    If rs.RecordCount > 0 Then
        cmbDays.Enabled = False
        cmbDays.Text = Format(Day(rs.Fields(0)), "00")
        LV.Checkboxes = True
        TDate = rs.Fields(0)
        TotalSalDays = Day(rs.Fields(0))
    Else
        cmbDays.Enabled = True
        LV.Checkboxes = False
        'TotalSalDays = DateAdd("m", 1, SDate) - SDate
        TotalSalDays = cmbDays
    End If
      
    rs.Close
    Set rs = Nothing
    
    LastDay = DateAdd("m", 1, SDate) - SDate
    
    Dim JDate As Date
    If TDate <= ServerDate Then
        TotalSalDays = cmbDays
    ElseIf Month(ServerDate) < Month(TDate) Then
        MsgBox "Can't Show Entries Of Future Months.", vbInformation
        Exit Sub
    End If
    
    TempTotalSaldays = TotalSalDays    'Store For Later Use
    
    JDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), TotalSalDays)
        
    'Sql = "Select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(JDate, "dd-MMM-yyyy") & "' Order By DeptID,EmpID"
    
    'cmd.CommandText = Sql
    
    cmd.CommandType = adCmdStoredProc
    
    cmd.CommandText = "SP_VendAdvSalary"
    
    Set cmd.ActiveConnection = con
    
    cmd.Parameters("@ParDate").value = JDate
    cmd.Parameters("@VendID").value = VendID
    cmd.Parameters("@ProcessID").value = ProcessID
    cmd.Parameters("@TotMonDays").value = cmbDays
    
    If rsLedger.State = 1 Then rsLedger.Close
    rsLedger.CursorType = adOpenKeyset
    rsLedger.LockType = adLockReadOnly
    Set rsLedger = cmd.Execute
    
    With rsLedger
        
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        
        
        Dim total As Double
        Dim Payed As Double
        Dim STLoan As Double
        Dim LongLoan As Double
        Dim VendWorkDone As Double
        Dim IsComputed As Boolean
        
        Dim DT As String
        Dim IsPayed As Boolean
        Dim IsTransfered As Boolean
        
        
        LV.ListItems.Clear
        LV.Visible = False
        
        
        
        If LV.Checkboxes = False And LV.ColumnHeaders.Count = 12 Then
                LV.ColumnHeaders.Remove 12
                LV.ColumnHeaders.Remove 11
        End If
        
        If LV.Checkboxes And LV.ColumnHeaders(LV.ColumnHeaders.Count).Text <> "Cheque #" Then
            LV.ColumnHeaders.Add , , "Pay By", 945.0709
            LV.ColumnHeaders.Add , , "Cheque #", 1094.74
        End If
        
        For i = 0 To .RecordCount - 1
                    
                    
            Call Computed(![VendID], IsComputed, STLoan, LongLoan, DT, IsPayed, IsTransfered, JDate, VendWorkDone)
            
            If Not IsComputed Then
                
                'Long Term Loan Deduction
                Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
                If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
                If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
                If IsNull(![ClsAmt]) Then lClsAmt = 0 Else lClsAmt = ![ClsAmt]
                If IsNull(![total]) Then VendWorkDone = 0 Else VendWorkDone = Val(![total] & "")
                
                If lDedAmt > (lTotAmt - lClsAmt) Then
                    LongLoan = (lTotAmt - lClsAmt)
                Else
                    LongLoan = lDedAmt
                End If
                
            End If
            
            'Dim theKey As String
            If IsComputed Then
                theKey = ![VendID] & "|" & Format(SDate, "dd-MMM-yyyy")
            Else
                theKey = ![VendID] & "'"
            End If
            Set ITM = LV.ListItems.Add(, theKey, ![VendID1] & "")
            If IsComputed Then
                ITM.Checked = IsPayed
                ITM.Bold = IsTransfered
            End If
            
            'itm.Tag = Val(![EmpType] & "")
            
            ITM.SubItems(1) = ![VenderName] & ""
            ITM.SubItems(2) = Format(Round(VendWorkDone), "0.00")
            
            ITM.SubItems(3) = Val(![Payed] & "")
            ITM.SubItems(4) = LongLoan
            ITM.SubItems(5) = Val(![TotalResDeduction] & "")
            
            ITM.SubItems(6) = Format(Round(Val(ITM.SubItems(2)) - Val(ITM.SubItems(3)) - Val(ITM.SubItems(5))), "0.00") '- LongLoan
            
            If IsNull(![AdvSalPer]) Then
                dAdvSalPer = 0
            Else
                dAdvSalPer = Val(![AdvSalPer] & "")
            End If
            ITM.SubItems(7) = dAdvSalPer 'Val(![AdvSalPer] & "")
            
            Dim NetPayable As Long
            
            'NetPayable = Round(Val(itm.SubItems(5)) * Val(![AdvSalPer] & "") / 10000, 0) * 100
            
            NetPayable = Round(Val(ITM.SubItems(6)) * dAdvSalPer / 10000, 0) * 100
            
            ITM.SubItems(8) = Format(NetPayable, "0.00")
            If NetPayable = 0 Then
                ITM.SubItems(9) = ""
            ElseIf NetPayable > 0 Then
                ITM.SubItems(9) = "DB"
                TotalDebit = TotalDebit + NetPayable
            Else
                ITM.SubItems(9) = "CR"
                TotalCredit = TotalCredit - NetPayable
            End If
            
            If Not IsNull(![CashPaid]) Then
                ITM.SubItems(10) = IIf((![CashPaid] & ""), "Cash", "Cheque")
                ITM.SubItems(11) = ![ChequeNo] & ""
            End If
            
            'Calculate Grand Total Payable
            GrandTotal = GrandTotal + NetPayable
            'End If
NEXT_EMPLOYEE:
            .MoveNext
        Next
        '.Close
        LV.Visible = True
    End With
    cmbMonth.Tag = cmbMonth.Text
    cmbYear.Tag = cmbYear.Text
    'Set rsLedger = Nothing
    
    Command1.Visible = LV.Checkboxes
    cmdAccounts.Visible = LV.Checkboxes
    
'    If LV.Checkboxes Then
'        LV.ColumnHeaders.Add , , "Pay By", 945.0709
'        LV.ColumnHeaders.Add , , "Cheque #", 1094.74
'    Else
'        If LV.ColumnHeaders.Count = 12 Then
'            LV.ColumnHeaders.Remove 12
'            LV.ColumnHeaders.Remove 11
'        End If
'    End If
    'Show Grand Total
    lblTotal = Format(GrandTotal, "#,##0.00")
    lblDebit = Format(TotalDebit, "#,##0.00")
    lblCredit = Format(TotalCredit, "#,##0.00")
    
    Sql = ""
    Cond = ""
    ICO = ""
    
End Sub
Private Sub Computed(VendID As Long, ByRef IsComputed As Boolean, ByRef STLoan As Double, ByRef LongLoan As Double, ByRef DT As String, ByRef Payed As Boolean, ByRef Transfered As Boolean, ToDate As Date, ByRef VendWorkDone As Double)

    Dim rsCheck As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        '.CommandText = "Select * from VSalarySettings where empid='" & Emp & "'"
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendSalarySettings"
        Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters("@ParDate").value = ToDate
        
        .Parameters("@VendID").value = VendID
        
        .Parameters("@AdvSalary").value = 1
        
        rsCheck.CursorType = adOpenForwardOnly
        rsCheck.LockType = adLockReadOnly
        Set rsCheck = .Execute
    End With
    
    With rsCheck
        
        If .RecordCount > 0 Then
            IsComputed = True
            
            
            DT = Format(![DT], "dd-MMM-yyyy") & ""
            ', AbsentDays, OverTimeHours, NTime, LTime)
            LV.Checkboxes = True
            'cmdUpdate.Visible = UserHasAccess("FinalAdvSal")
            
            Payed = ![Payed] & ""
            Transfered = ![Transfered] & ""
            
            If IsNull(![LongLoan]) Then LongLoan = 0 Else LongLoan = ![LongLoan]
            STLoan = Val(![STAmt] & "")
            
            'Get The Salary Depending On Salary Sheet Type
            'VendWorkDone = IIf(chkExternal = vbUnchecked, Val(![Total] & ""), Val(![ExtSalary] & ""))
            VendWorkDone = Val(![WorkDone] & "")
            
            'EmpSal = Val(![StartingSalary] & "")
            LV.Checkboxes = True
            cmdAccounts.Visible = True
            Command1.Visible = True
            cmdUpdate.Caption = "Save Changes"
            'cmdUpdate.Caption = "Save Changes"
            
        Else
            
            IsComputed = False
            
            cmdUpdate.Caption = "Make This Sheet Final"
            
        End If
        
        
    End With
    
End Sub

Private Sub SaveIt(rs As ADODB.Recordset, PerDeduct As Double, Optional Paid As Boolean, Optional PaidAmt As Double, Optional dPer As Double)

    Dim DT As Date
    Dim LongLoan As Double

    Dim lDedAmt As Long, lClsAmt1 As Long, lTotAmt As Long
    If IsNull(rs![DedAmt]) Then lDedAmt = 0 Else lDedAmt = rs![DedAmt]
    If IsNull(rs![totamt]) Then lTotAmt = 0 Else lTotAmt = rs![totamt]
    If IsNull(rs![ClsAmt]) Then lClsAmt1 = 0 Else lClsAmt1 = rs![ClsAmt]
                
    If lDedAmt > (lTotAmt - lClsAmt1) Then
        LongLoan = (lTotAmt - lClsAmt1)
    Else
        LongLoan = lDedAmt
    End If

    'If IsNull(rs![LongLoan]) Then LongLoan = 0 Else LongLoan = rs![LongLoan]
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), cmbDays)
    con.Execute "Insert into VendMonthlyPayments (VendID,DT,WorkDone" & _
     ",Payed,LongLoan,ResDed,Transfered,AdvSalary,PaidAmt,ShortTermPer,STAmt)" & _
     " Values('" & rs![VendID] & "','" & DT & _
     "'," & Val(rs![total] & "") & "," & _
     "0," & LongLoan & "," & Val(rs![TotalResDeduction] & "") & ",0,1," & _
     IIf(Paid, PaidAmt, 0) & "," & dPer & "," & Val(rs!Payed & "") & ")"
     
    'Update Long Term ...
    'Dim LongAmt As Double
    'LongAmt = Val(LV.ListItems(LV.ListItems.Count).SubItems(17))
    
    'Con.Execute "Insert Into AmtCleared (EmpID,DT,AmtClrd) Values('" & rs![EmpID] & "','" & DT & "'," & LongAmt & ")"
    
    
    If LV.Checkboxes = False Then
        LV.Checkboxes = True
        'cmdUpdate.Visible = UserHasAccess("FinalAdvSal")
    End If
    
End Sub
Private Function TotalMonthDays() As Integer
    Dim TDate As Date
    Dim Tot As Integer
    Dim year As Integer
    If cmbYear = "" Then
        year = 2002
    Else
        year = cmbYear
    End If
    
    If Month(ServerDate) = cmbMonth.List(cmbMonth.ListIndex, 1) Then
        Tot = Day(ServerDate)
    Else
        TDate = DateSerial(year, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
        Tot = DateAdd("m", 1, TDate) - TDate
    End If
    TotalMonthDays = Tot
End Function


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Exit Sub 'This Will Be Done ON Separate Form...
    If Item.Checked Then
        Item.SubItems(10) = "Cash"
        Item.SubItems(11) = ""
    Else
        Item.SubItems(10) = ""
        Item.SubItems(11) = ""
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Set Flag that mouse is in Issue column or In Rcvd.
    If Not LV.Checkboxes Then Exit Sub
    If CBool(x >= LV.ColumnHeaders(11).Left And x <= (LV.ColumnHeaders(11).Left + LV.ColumnHeaders(11).Width)) Then
        ColIndex = 11
    ElseIf CBool(x >= LV.ColumnHeaders(12).Left And x <= (LV.ColumnHeaders(12).Left + LV.ColumnHeaders(12).Width)) Then
        ColIndex = 12
    Else
        ColIndex = 0
    End If
End Sub
Private Sub LV_DblClick()
    
    Exit Sub 'This Will Be Done On Separate Form....
    If LV.ListItems.Count = 0 Then Exit Sub
    
    If (ColIndex = 0) Then Exit Sub
    
    
        If LV.SelectedItem.SubItems(10) = "Cash" Then
            LV.SelectedItem.SubItems(10) = "Cheque"
            GoTo EnterChequeNo
        
        Else
            LV.SelectedItem.SubItems(10) = "Cash"
            Exit Sub
        End If
    
EnterChequeNo:
    With txtEdit
        .Move LV.ColumnHeaders(12).Left + LV.Left + 50, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(12).Width
        .Text = LV.SelectedItem.SubItems(11) & ""
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub mnuPaySlips_Click()
   Call PrintReport(True)
End Sub

Private Sub mnuPrintChqs_Click()

    Call PrintCheques

End Sub


Private Sub PrintCheques(Optional bPrintVoucher As Boolean = False)

    On Error GoTo err
    
    If LV.ListItems.Count = 0 Then Exit Sub
        
    
    On Error GoTo err
    
    '****************************************************************
    Dim rs As New ADODB.Recordset
    Dim myDT As Date
    With rs
        .Open "Select * From ChequePrintingDates Where Year(SalaryDT)=" & cmbYear & " AND Month(SalaryDT)=" & cmbMonth.ListIndex + 1 & " AND AdvSalary=1", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            myDT = !DT
        Else
            myDT = Date
        End If
        .Close
    End With
    
    Set rs = Nothing
    
    If bPrintVoucher = False Then
        If MsgBox("Do you want to print cheques on " & Format(myDT, "dd-MMM-yyyy") & " ?", vbQuestion + vbYesNo) = vbNo Then
            myDT = frmDT.GetDate(myDT)
        End If
        
        'DT = DateSerial(DTMonth.year, DTMonth.Month + 1, 0)
        con.Execute "Delete From ChequePrintingDates Where Year(SalaryDT)=" & cmbYear & " AND Month(SalaryDT)=" & cmbMonth.ListIndex + 1 & " AND AdvSalary=1"
        con.Execute "Insert Into ChequePrintingDates(SalaryDT,DT,AdvSalary) Values('" & DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0) & "','" & myDT & "',1)"
    End If
    '******************************************************************
    
    Screen.MousePointer = vbHourglass
    
    Dim DtFrom As Date
    Dim DtTo As Date
    DtFrom = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    DtTo = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), cmbDays)
    
    Dim rpt As craxddrt.Report
    If bPrintVoucher = False Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptVendChqs.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptVendJV.rpt")
    End If
    
    Dim F As New frmPrevRpt
    rpt.FormulaFields(2).Text = "'" & Format(myDT, "dd-MMM-yyyy") & "'"
    F.ShowReport "{VVendMonthlyPayments.DT}=#" & DtTo & "# AND {VVendMonthlyPayments.PaidAmt}>0 AND {VVendMonthlyPayments.AdvSalary}=True AND {VVendMonthlyPayments.Payed}=True ", rpt
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuPrintVouchers_Click()

    Call PrintCheques(True)
    
End Sub

Private Sub mnuSalSheet_Click()
   Call PrintReport(False)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(11) = txtEdit.Text
        txtEdit.Text = ""
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Text = ""
        txtEdit.Visible = False
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
End Sub

Private Function AddSalaryEntries() As Boolean

   Dim SSAdv As Double, SLadv As Double, SSal As Double, STax As Double     'Staff Member Vars
   Dim DSAdv As Double, DLadv As Double, DSal As Double, DTax As Double     'Directors Vars
   Dim PSAdv As Double, PLadv As Double, PSal As Double, PTax As Double     'Production Staff Vars
   
   
   
   Dim Dated As Date
   Dated = ServerDate
   
   '''''''''''''''''' A D V A N C E   C L E A R E N C E ''''''''''''''''''
   ConAcc.BeginTrans
   
   Call GetTotal(SSAdv, DSAdv, PSAdv, "SADV")        'Short Term Advances
   Call AddAdvClearVoucher(SSAdv, DSAdv, PSAdv, Dated, False, cmbMonth & ", " & cmbYear)
   
   Call GetTotal(SLadv, DLadv, PLadv, "LADV")        'Long  Term Advances
   Call AddAdvClearVoucher(SLadv, DLadv, PLadv, Dated, True, cmbMonth & ", " & cmbYear)
   
   
   '''''''''''''''''' S A L R A R Y   &   T A X    E N T R Y ''''''''''''''''''
   
   Call GetTotal(SSal, DSal, PSal, "SAL")            'Salaries
   Call GetTotal(STax, DTax, PTax, "TAX")            'Taxes
   
   Call addSalaryVoucher(SSal, DSal, PSal, Dated, STax, DTax, PTax, cmbMonth & ", " & cmbYear)
   
   ConAcc.CommitTrans
   
   
   
    For i = 1 To LV.ListItems.Count
    
        con.Execute "Update MonthlySalaries Set Transfered=" & CBool(LV.ListItems(i).Checked Or LV.ListItems(i).Bold) & " Where EmpID+Format(DT,'dd-MMM-yyyy')='" & LV.ListItems(i).Key & "'"
    
    Next i
      
   
   AddSalaryEntries = True
   Exit Function
err:
   AddSalaryEntries = False
   MsgBox "Cannot Add Salary Vouchers Due To Error in " & err.Source & " :" & vbNewLine & err.Description, vbCritical
   ConAcc.RollbackTrans
End Function

Private Sub GetTotal(Stotal As Double, DTotal As Double, PTotal As Double, TotalType As String)

   Dim StaffTotal As Double, ProductionTotal As Double, DirectorTotal As Double, Amount As Double



   For i = 1 To LV.ListItems.Count
      'if Paid or Transfered then
      If Not LV.ListItems(i).Checked Or LV.ListItems(i).Bold Then GoTo NEXT_ITEM
         
   
         Select Case TotalType
            Case "LADV"
               'Long Term Advances
               Amount = Val(LV.ListItems(i).SubItems(17))
            Case "SADV"
               'Short Tem Advances
               Amount = Val(LV.ListItems(i).SubItems(16))
            Case "SAL"
               'Salary amount
               Amount = Val(LV.ListItems(i).SubItems(12))
            Case "TAX"
               Amount = Val(LV.ListItems(i).SubItems(14))
         End Select
   
   
   
   
    Select Case LV.ListItems(i).Tag
      Case "0"
         StaffTotal = StaffTotal + Amount
      Case "1"
         DirectorTotal = DirectorTotal + Amount
      Case "2"
         ProductionTotal = ProductionTotal + Amount
    End Select


NEXT_ITEM:
   Next
   
   PTotal = ProductionTotal
   Stotal = StaffTotal
   DTotal = DirectorTotal
   
End Sub

