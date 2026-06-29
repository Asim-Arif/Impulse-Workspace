VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmAdvancePaymentList 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12825
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   12825
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   8670
      Index           =   0
      Left            =   0
      ScaleHeight     =   8670
      ScaleWidth      =   12840
      TabIndex        =   0
      Tag             =   "ADVANCE PAYMENTS LIST"
      Top             =   0
      Width           =   12840
      Begin VB.Frame FraUsage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail of Usage :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   4785
         TabIndex        =   15
         Top             =   1905
         Visible         =   0   'False
         Width           =   4980
         Begin VB.CommandButton cmdHide 
            Caption         =   "Hide Detail"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3480
            TabIndex        =   17
            Top             =   3060
            Width           =   1425
         End
         Begin MSComctlLib.ListView LVDetail 
            Height          =   2715
            Left            =   90
            TabIndex        =   16
            Top             =   300
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   4789
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   16777215
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Invoice #"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin VB.TextBox txtEdit 
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
         Left            =   2940
         TabIndex        =   14
         Top             =   2520
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9690
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1395
         Width           =   1545
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   11175
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8100
         Width           =   1545
      End
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmAdvancePaymentList.frx":0000
         Left            =   4305
         List            =   "frmAdvancePaymentList.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1530
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8055
         TabIndex        =   2
         Top             =   1530
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
         Format          =   114819075
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6330
         TabIndex        =   3
         Top             =   1530
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
         Format          =   114819075
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   4
         Top             =   1530
         Width           =   2310
         _ExtentX        =   4075
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
      Begin MSComctlLib.ListView LV 
         Height          =   6135
         Left            =   120
         TabIndex        =   5
         Top             =   1860
         Width           =   12630
         _ExtentX        =   22278
         _ExtentY        =   10821
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16777215
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
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   2619
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   3360
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Advice #"
            Object.Width           =   3625
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Exch Rate"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Status"
            Object.Width           =   4789
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B54646&
         Height          =   195
         Index           =   3
         Left            =   8055
         TabIndex        =   12
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B54646&
         Height          =   195
         Index           =   2
         Left            =   6330
         TabIndex        =   11
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Range :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B54646&
         Height          =   195
         Index           =   1
         Left            =   4305
         TabIndex        =   10
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B54646&
         Height          =   195
         Index           =   0
         Left            =   2415
         TabIndex        =   9
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B54646&
         Height          =   195
         Index           =   7
         Left            =   75
         TabIndex        =   8
         Top             =   1320
         Width           =   915
      End
      Begin VB.Image Image2 
         Height          =   1200
         Left            =   150
         Picture         =   "frmAdvancePaymentList.frx":0066
         Top             =   60
         Width           =   1200
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   1530
         Width           =   1860
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Popup"
      Begin VB.Menu mnuAddNew 
         Caption         =   "Add New Advance Payment"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Advance Payment"
      End
      Begin VB.Menu mnuUpdateStatus 
         Caption         =   "<<Update Status>>"
      End
      Begin VB.Menu mnuDashPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print This Advance Payment"
      End
      Begin VB.Menu mnuDashUsage 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowUsageDetail 
         Caption         =   "Show Usage Detail"
      End
      Begin VB.Menu mnuDashPosttoPRC 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPosttoPRC 
         Caption         =   "Post to PRC"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAdvancePaymentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "Select Distinct Country From ForeignCustomers "
    
    If cmbCust.ID <> "0" Then
        SQL = SQL & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        
        cmbCountry.Clear
        cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            'cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = GetServerDate
        Case 1
            DtFrom = DateAdd("d", -15, GetServerDate)
        Case 2
            DtFrom = DateAdd("d", -30, GetServerDate)
        Case 3
            DtFrom = DateAdd("d", -60, GetServerDate)
        Case 4
            DtFrom = DateAdd("d", -90, GetServerDate)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdHide_Click()
    FraUsage.Visible = False
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub Form_Load()
    
    Call DrawPicCaptions
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode")
    cmbCust.ID = "0"
    'DtFrom = GetServerDate
    'DtTo = GetServerDate
    cmbRange.ListIndex = 0
    mnuPop.Visible = False
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
     
    Cond = " WHERE DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND CustCode='" & cmbCust.ID & "' "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        If Cond = "" Then
            Cond = " AND Country='" & cmbCountry & "'"
        Else
            Cond = Cond & " AND Country='" & cmbCountry & "'"
        End If
    End If
  
    Dim strStatus As String
    With rs
        
        .Open "Select * From FCustAdvancePayments " & Cond & " Order By DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![EntryID] & "'", .AbsolutePosition)
            'ITM.Tag = ![Curr] & ""
            ITM.ListSubItems.add(, , ![CustCode] & "").Tag = Val(!EntryID & "")
            
            ITM.ListSubItems.add(, , ![Country] & "").Tag = ""
            
            ITM.ListSubItems.add , , ![AdviceNo] & ""
            
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            
            ITM.SubItems(5) = Val(![Amount] & "")
            
            ITM.SubItems(6) = Val(![ExchRate] & "")
            
            If Val(!ExchRate & "") = 0 Then
                strStatus = "Advice Feeded Only."
            ElseIf !VchrNo & "" = "" Then
                strStatus = "Ready to Post"
            Else
                strStatus = "Posted to Financial"
            End If
            ITM.SubItems(7) = strStatus
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LV.SelectedItem.SubItems(7) = "Posted to Financial" Then
        MsgBox "Already Posted, Can't Edit.", vbInformation
        Exit Sub
    End If
    Load frmAdvancePayment
    
    Call frmAdvancePayment.EditMe(Val(LV.SelectedItem.ListSubItems(1).Tag))
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then
        mnuEdit.Enabled = False
        mnuPrint.Enabled = False
        mnuUpdateStatus.Enabled = False
    Else
        mnuEdit.Enabled = True
        mnuPrint.Enabled = True
        mnuUpdateStatus.Enabled = True
        If Val(LV.SelectedItem.SubItems(6)) = 0 Then
            strStatus = "Update Exchange Rate"
        ElseIf LV.SelectedItem.SubItems(7) = "Ready to Post" Then
            strStatus = "Post to Financial"
        Else
            mnuUpdateStatus.Enabled = False
        End If
        mnuUpdateStatus.Caption = strStatus
    End If
    
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    For i = 0 To PicWC.count - 1
        Call DrawCaptions(PicWC(i))
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox)

    Dim lBorderColor As Long, lFillColor
    lBorderColor = RGB(70, 70, 181)
    lFillColor = RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 350)-(Pic.Width, 350), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 600), lFillColor, BF
        Pic.ForeColor = vbWhite
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 180
        Pic.Print Pic.Tag
    End If
        
    Pic.Line (0, Pic.Height - 20)-(Pic.Width, Pic.Height - 20), lBorderColor 'Bottom

    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 30, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right

End Sub

Private Sub mnuAddNew_Click()
    Dim f As New frmAdvancePayment
    Load f
    Call f.ShowMe
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuPosttoPRC_Click()
'    If LV.ListItems.count = 0 Then Exit Sub
'    If MsgBox("Are you sure ?", vbQuestion + vbYesNo) = vbYes Then
'        Dim f As New frmPRC
'        Load f
'        f.ShowMe Val(LV.SelectedItem.key)
'    End If
End Sub

Private Sub mnuPrint_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptAdvancePayments.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustAdvancePayments.EntryID}=" & Val(LV.SelectedItem.key), rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuShowUsageDetail_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem, dTotal As Double
    With rs
        .Open "SELECT * FROM FCustAdvancePaymentsAdjusted WHERE RefID=" & Val(LV.SelectedItem.key) & " ORDER BY EntryID ", con, adOpenForwardOnly, adLockReadOnly
        LVDetail.ListItems.Clear
        dTotal = 0
        Do Until .EOF
            Set ITM = LVDetail.ListItems.add(, , LVDetail.ListItems.count + 1)
            ITM.ListSubItems.add , , !CustomInvoice & ""
            ITM.ListSubItems.add , , Val(!Amount & "")
            dTotal = dTotal + Val(!Amount & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    If LVDetail.ListItems.count > 0 Then
        Set ITM = LVDetail.ListItems.add
        ITM.ListSubItems.add(, , "Total").Bold = True
        ITM.ListSubItems.add(, , dTotal).Bold = True
        ITM.ListSubItems(2).ForeColor = vbRed
        FraUsage.Visible = True
    Else
        MsgBox "No usage Details found.", vbInformation
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuUpdateStatus_Click()
    
    On Error GoTo err
    If mnuUpdateStatus.Caption = "Update Exchange Rate" Then
        Call EditExchangeRate
    ElseIf mnuUpdateStatus.Caption = "Post to Financial" Then
        'Post Voucher
        If Trim(LV.SelectedItem.SubItems(3)) = "" Then
            MsgBox "Please Enter Advice No.", vbInformation
            Exit Sub
        End If
        Dim f As New frmAdvancePaymentPosting
        Load f
        f.ShowMe (Val(LV.SelectedItem.key))
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub EditExchangeRate()
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(7).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(7).Width
        .Text = Val(LV.SelectedItem.SubItems(6))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtEdit) <= 0 Then
            MsgBox "Invalid Exch. Rate", vbInformation
            Exit Sub
        End If
        con.Execute "UPDATE FCustAdvancePayments SET ExchRate=" & Val(txtEdit) & " WHERE EntryID=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.SubItems(6) = Val(txtEdit)
        LV.SelectedItem.SubItems(7) = "Ready to Post"
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
