VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmPRCList 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   8670
   ClientLeft      =   15
   ClientTop       =   0
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
   ScaleHeight     =   8670
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
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
      Tag             =   "PRC LIST"
      Top             =   0
      Width           =   12840
      Begin VB.CheckBox chkUnposted 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unposted"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   11280
         TabIndex        =   15
         Top             =   1560
         Width           =   1275
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
         ItemData        =   "frmPRCList.frx":0000
         Left            =   4305
         List            =   "frmPRCList.frx":0016
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
         Format          =   53215235
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
         Format          =   53215235
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
         Left            =   75
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
         NumItems        =   9
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
            Text            =   "Invoice #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "PRC #"
            Object.Width           =   3625
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Exch Rate"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Amoun (Rs.)"
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
         Picture         =   "frmPRCList.frx":0066
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
End
Attribute VB_Name = "frmPRCList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean

Public Sub ShowMe()
    Me.Show
End Sub


Private Sub cmbCust_matched()

    On Error GoTo Err
    
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
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
Err:
    MsgBox Err.Description
    
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

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub Form_Load()
    
    Call DrawPicCaptions
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode")
    
    'DtFrom = GetServerDate
    'DtTo = GetServerDate
    cmbRange.ListIndex = 0

    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo Err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    If cmbCust.ID <> "0" Then
        Cond = Cond & " WHERE CustCode='" & cmbCust.ID & "' "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        If Cond = "" Then
            Cond = " Where Country='" & cmbCountry & "'"
        Else
            Cond = Cond & " and Country='" & cmbCountry & "'"
        End If
    End If
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " BillNoDT Between '" & DateAdd("d", 0, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " BillNoDT Between '" & DateAdd("d", -15, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " BillNoDT Between '" & DateAdd("d", -30, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " BillNoDT Between '" & DateAdd("d", -60, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " BillNoDT Between '" & DateAdd("d", -90, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " BillNoDT Between '" & DtFrom & "' and '" & DtTo & "'"
    End If
    
    If Cond = "" Then
        Cond = " Where "
    Else
        Cond = Cond & " and "
    End If
    If chkUnposted.value = vbChecked Then
        Cond = Cond & " CustomInvoice NOT IN(SELECT CustomInvoice FROM PRC) AND CAST(SUBSTRING(CustomInvoice,5,LEN(CustomInvoice)-4) AS INT)>1600"
    Else
        Cond = Cond & DTRange
    End If
    
    Dim strStatus As String
    With rs
        If chkUnposted.value = vbChecked Then
            .Open "SELECT CustCode,Country,OurBankID,CustomInvoice,AmtInPakRs,ExchRate FROM VCustomInvoiceAuthorized " & Cond & " Order By CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select * From VPRC " & Cond & " Order By BillNoDT", con, adOpenForwardOnly, adLockReadOnly
        End If
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , .AbsolutePosition)
            'ITM.Tag = ![Curr] & ""
            ITM.ListSubItems.add(, , ![CustCode] & "").Tag = 0 'Val(!EntryID & "")
            
            ITM.ListSubItems.add(, , ![Country] & "").Tag = ""
            ITM.ListSubItems.add , , ![CustomInvoice] & ""
            
            If chkUnposted.value = vbUnchecked Then
                ITM.ListSubItems.add , , ![BillNo] & ""
                
                ITM.ListSubItems.add , , Format(![RealizationDT] & "", "dd-MMM-yyyy")
                
                ITM.SubItems(6) = Val(![AmtRealized] & "")
                
                ITM.SubItems(7) = Val(![ExchRate] & "")
                ITM.SubItems(8) = Round(Val(![ExchRate] & "") * Val(![AmtRealized] & ""), 2)
            End If
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
Err:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
    
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

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If Button <> 2 Then Exit Sub
    Exit Sub
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


Private Sub mnuEdit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuPrint_Click()

    On Error GoTo Err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptExportPath & "\rptAdvancePayments.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustAdvancePayments.EntryID}=" & Val(LV.SelectedItem.key), rpt
    
    Exit Sub
Err:
    MsgBox Err.Description
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
