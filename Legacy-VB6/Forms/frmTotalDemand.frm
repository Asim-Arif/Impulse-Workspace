VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmTotalDemand 
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12225
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12225
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   12120
      Begin VB.ComboBox cmbRange 
         Height          =   315
         ItemData        =   "frmTotalDemand.frx":0000
         Left            =   2445
         List            =   "frmTotalDemand.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1035
         Width           =   2010
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1050
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   3
         Top             =   15
         Width           =   12120
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Demand"
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
            TabIndex        =   4
            Top             =   150
            Width           =   12000
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Demand"
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
            TabIndex        =   5
            Top             =   165
            Width           =   11985
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6585
         Left            =   285
         TabIndex        =   10
         Top             =   1365
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   11615
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cust Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Order No."
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Exch Rate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Rs."
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   6195
         TabIndex        =   16
         Top             =   1035
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   109707267
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   4470
         TabIndex        =   17
         Top             =   1035
         Width           =   1710
         _ExtentX        =   3016
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
         Format          =   109707267
         CurrentDate     =   37055
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   8655
         TabIndex        =   14
         Top             =   8535
         Width           =   1665
         Caption         =   "Print             "
         PicturePosition =   327683
         Size            =   "2937;635"
         Picture         =   "frmTotalDemand.frx":0066
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkCustom 
         Height          =   300
         Left            =   9465
         TabIndex        =   13
         Top             =   1020
         Width           =   2250
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3969;529"
         Value           =   "0"
         Caption         =   "Show Custom Amount"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   7830
         TabIndex        =   12
         Top             =   990
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   360
         Left            =   10350
         TabIndex        =   11
         Top             =   8535
         Width           =   1665
         Caption         =   "      Close    "
         PicturePosition =   327683
         Size            =   "2937;635"
         Picture         =   "frmTotalDemand.frx":0178
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox lblTotWords 
         Height          =   270
         Left            =   2100
         TabIndex        =   9
         Top             =   7980
         Width           =   7545
         VariousPropertyBits=   746604575
         BackColor       =   15199215
         BorderStyle     =   1
         Size            =   "13309;476"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount :"
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
         Height          =   270
         Index           =   1
         Left            =   9660
         TabIndex        =   8
         Top             =   7980
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount (In Words) :"
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
         Height          =   270
         Index           =   2
         Left            =   285
         TabIndex        =   7
         Top             =   7980
         Width           =   1830
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10530
         TabIndex        =   6
         Top             =   7980
         Width           =   1500
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer                             Range                                 From                            To"
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
         Index           =   7
         Left            =   285
         TabIndex        =   2
         Top             =   810
         Width           =   7455
      End
   End
End
Attribute VB_Name = "frmTotalDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String
Dim strSQL_Report As String
Private Sub RefreshLV()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    Dim ITM As ListItem
    Dim lColor As Long
    Dim Cond As String
    Dim dTotal As Long
    Dim Amt As Long
    Cond = " WHERE DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND CustCode='" & cmbCust.ID & "'"
    End If
'    Cond = Cond & " AND OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList)"
    Cond = Cond & " AND OrderNo NOT IN(SELECT OrderNo FROM FCustomerFinalOrders WHERE Cancelled IN(1,2))"
    Cond = Cond & " AND OrderType='Customer Order' "
    Cond = Cond & " AND CustCode NOT IN('HMP')"
    LV.ListItems.Clear
        
    With rs
        strSQL_Report = "SELECT * FROM VFOrderList" & Cond
        .Open strSQL_Report, con, adOpenForwardOnly, adLockReadOnly
        dTotal = 0
        Do Until .EOF
            
            
            Amt = Val(![OrderAmt] & "")
           
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            
            ITM.ListSubItems.add , , ![CustCode] & ""
            ITM.ListSubItems.add , , ![Country] & ""
            
            ITM.ListSubItems.add , , ![OrderNo] & ""
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Amt
            'ITM.ListSubItems.Add , , Val(![ExchRate] & "")
'            ITM.ListSubItems.Add , , Amt * Val(![ExchRate] & "")
            dTotal = dTotal + Amt '* Val(![ExchRate] & "")
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    lblTotal = dTotal
    lblTotWords = ConvertInWords(dTotal, "")
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub cmbRange_Change()
    Call cmbRange_Click
End Sub

Private Sub cmbRange_Click()
    
    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report 'New rptTotalExport
    Set rpt = rptApp.OpenReport(RptPath & "\TotalDemand.rpt") '
    Dim f As New frmPrevRpt
    Dim Cond As String
    
    rpt.FormulaFields.GetItemByName("CustomAmt").Text = chkCustom.Value
    rpt.FormulaFields.GetItemByName("SubHeading").Text = "'" & Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy") & "'"
    
    Dim strSelection As String
    strSelection = "{VTotalExport.DT}=#" & DtFrom.Value & "# To #" & DtTo.Value & "#"
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {VTotalExport.CustCode}='" & cmbCust.ID & "' "
    End If
    
    Dim rptSub As CRAXDDRT.Report
    Set rptSub = rpt.OpenSubreport("currencytotalexport")
    'rptSub.RecordSelectionFormula = strSelection
    rptSub.SQLQueryString = strSQL_Report
    rptSub.EnableParameterPrompting = False
    rpt.EnableParameterPrompting = False
    f.ShowReport "", rpt, strSQL_Report

    Set rpt = Nothing
    Set f = Nothing
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub



Private Sub Form_Load()
   
    On Error GoTo err
    

    
    Call SaveLV(LV)
    
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    cmbRange.ListIndex = 4
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
    
End Sub
