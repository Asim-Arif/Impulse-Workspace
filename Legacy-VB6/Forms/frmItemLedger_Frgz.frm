VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmItemLedger_Frgz 
   ClientHeight    =   9660
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11370
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleMode       =   0  'User
   ScaleWidth      =   10184.26
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   105
      TabIndex        =   19
      Top             =   0
      Width           =   11190
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Ledger"
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
         Index           =   2
         Left            =   4650
         TabIndex        =   20
         Top             =   135
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Ledger"
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
         Index           =   3
         Left            =   4665
         TabIndex        =   21
         Top             =   150
         Width           =   1920
      End
   End
   Begin VB.PictureBox ImgLoading 
      Height          =   3600
      Left            =   1515
      ScaleHeight     =   3540
      ScaleWidth      =   7845
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   7905
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Please wait While Loading Data"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1860
         Left            =   1350
         TabIndex        =   18
         Top             =   960
         Width           =   5310
      End
   End
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   105
      TabIndex        =   10
      Top             =   855
      Width           =   6180
      _ExtentX        =   10901
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
   Begin ComboList.Usercontrol1 cmbMaterialID 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   855
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
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
   Begin VB.Frame FLV 
      Height          =   7710
      Left            =   105
      TabIndex        =   9
      Top             =   1920
      Width           =   11265
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3090
         TabIndex        =   11
         Top             =   2115
         Visible         =   0   'False
         Width           =   3615
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7020
         Left            =   75
         TabIndex        =   4
         Top             =   180
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   12383
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Transaction Type"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Customer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Order No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Location"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remarks"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Stock"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Lot #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Batch #"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrintBinCard 
         Height          =   360
         Left            =   6765
         TabIndex        =   22
         Top             =   7260
         Width           =   1425
         ForeColor       =   64
         Caption         =   "Print Bin Card"
         PicturePosition =   327683
         Size            =   "2514;635"
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
         Left            =   9765
         TabIndex        =   6
         Top             =   7260
         Width           =   1425
         ForeColor       =   64
         Caption         =   " Close      "
         PicturePosition =   327683
         Size            =   "2514;635"
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
         Left            =   8250
         TabIndex        =   5
         Top             =   7260
         Width           =   1425
         ForeColor       =   64
         Caption         =   "Print         "
         PicturePosition =   327683
         Size            =   "2514;635"
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
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   1410
      Left            =   105
      TabIndex        =   8
      Top             =   510
      Width           =   11190
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   7965
         TabIndex        =   2
         Top             =   345
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   146014211
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6210
         TabIndex        =   1
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
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
         Format          =   104071171
         CurrentDate     =   37055
      End
      Begin VB.Label lblUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   9900
         TabIndex        =   16
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label lblStock 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   7620
         TabIndex        =   15
         Top             =   885
         Width           =   2265
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   2715
         TabIndex        =   14
         Top             =   885
         Width           =   4890
      End
      Begin VB.Label lblSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   45
         TabIndex        =   13
         Top             =   885
         Width           =   2655
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmItemLedger_Frgz.frx":0000
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
         Height          =   210
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   660
         Width           =   11130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   9615
         TabIndex        =   3
         Top             =   300
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmItemLedger_Frgz.frx":00D7
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
         Height          =   210
         Index           =   0
         Left            =   15
         TabIndex        =   7
         Top             =   120
         Width           =   11130
      End
   End
End
Attribute VB_Name = "frmItemLedger_Frgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrintSQL As String

Private Sub cmbMaterial_matched()
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ItemSize,Unit,InHand,ItemColor,SizeUnit FROM VItems WHERE ItemID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lblColor.Caption = !ItemColor & ""
            lblSize.Caption = !ItemSize & " " & !SizeUnit
            lblStock.Caption = GetSingleLongValue("SUM(NetQty)", "VItemsWithShelfWiseStock", " WHERE ItemID='" & cmbMaterial.ID & "'")
            lblUnit.Caption = !Unit & ""
        Else
            lblColor.Caption = ""
            lblSize.Caption = ""
            lblStock.Caption = ""
            lblUnit.Caption = ""
        End If
        .Close
    End With
    Set rs = Nothing
    
    Call cmdRefresh_Click
End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt
    

    Set rpt = rptApp.OpenReport(RptPath & "\ItemLedger.rpt")
    rpt.ParameterFields.GetItemByName("@ItemID").SetCurrentValue cmbMaterial.ID
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    
    rpt.EnableParameterPrompting = False
    
    rpt.FormulaFields.GetItemByName("ForPeriod").Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
    f.ShowReport "", rpt, "", False
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrintBinCard_Click()
    
    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt
    

    Set rpt = rptApp.OpenReport(RptPath & "\ItemLedger_BC.rpt")
    rpt.ParameterFields.GetItemByName("@ItemID").SetCurrentValue cmbMaterial.ID
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    
    rpt.EnableParameterPrompting = False
    
    rpt.FormulaFields.GetItemByName("ForPeriod").Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
    f.ShowReport "", rpt, "", False
    Exit Sub
    
err:
    MsgBox err.Description

End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLedger
    cmbMaterial.Tag = cmbMaterial.ID
    cmbMaterial.SetFocus
End Sub

Private Sub Form_Load()

    DtTo = Date
    DtFrom = DateAdd("M", -1, Date)
    cmbMaterialID.ListHeight = 6000
    cmbMaterial.ListHeight = 6000
    'cmbMaterialID.AddVals con, "ItemID", "Items", "ItemID", " Order By ItemID"
    
   'Call GetBankAccountsInMenu
    
    
End Sub


Private Sub RefreshLedger()

    If Not cmbMaterial.MatchFound Then Exit Sub
    
    Dim rsLedger As New ADODB.Recordset
    
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "FinishItemLedger_SP"
        .Parameters("@ItemID").Value = cmbMaterial.ID
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value
        Set rsLedger = .Execute
    End With
    
    Dim lOpeningBalance As Long, lBalance As Long
    Dim ITM As ListItem
    
    With rsLedger
         
        LV.ListItems.Clear
        Dim sItem As String
    
        For i = 0 To .RecordCount - 1
            If i = 0 Then
                Set ITM = LV.ListItems.add(, "Opening", "")
                ITM.ListSubItems.add , , "Opening Balance"
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , Val(!OpeningBalance & "")
                lBalance = Val(!OpeningBalance & "")
            End If
            
            If IsNull(!RcvNo) Then GoTo TheEnd
            Set ITM = LV.ListItems.add(, "", Format(![DT], "MMM dd, yyyy"))
            
            ITM.SubItems(1) = !RcvdFrom & " " & !RcvdBy
            ITM.SubItems(2) = Round(Val(![QtyIn] & ""), 2)
            ITM.SubItems(3) = Round(Val(![QtyOut] & ""), 2)
             
            ITM.SubItems(4) = !CustCode & ""
            ITM.SubItems(5) = !OrderNo & ""
            ITM.SubItems(6) = !Location & ""
            ITM.SubItems(7) = !Remarks & ""
            
            lBalance = lBalance + Val(!QtyIn & "") - Val(!QtyOut & "")
            ITM.SubItems(8) = Round(lBalance, 2)
            ITM.SubItems(9) = !LotNo_Manual & ""
            ITM.SubItems(10) = !BatchNo_Manual & ""
       
TheEnd:
            .MoveNEXT
        
        Next
    
    End With

    DtTo.Tag = DtTo
    DtFrom.Tag = DtFrom
     
End Sub

Private Sub refreshItemLedger()

    Set LVItems.SmallIcons = ImageList1
    Dim ListColor As Long
    Dim CrsLedger As New ADODB.Recordset
    
    With CrsLedger
        .CursorLocation = adUseClient
        .Open "Select OrderNo,ItemID,ItemName,Price,Qty,Rcvd from VOrdItems Where (OrderNo='" & LV.SelectedItem & "') ", con, adOpenForwardOnly '& _
            '"and (CDate between #" & DtFrom & "# and #" & DtTo & "#) " & _
            '"Order by CDate,Sno", con, 0, 1
            
        LVItems.ListItems.Clear
        If .RecordCount > 0 Then
            Dim ITM As ListItem
            Do Until .EOF
                If ![Qty] - ![Rcvd] = 0 Then
                  ListColor = vbBlack
                
                Else
                  ListColor = vbRed
                End If
                
                Set ITM = LVItems.ListItems.add(, ![OrderNo] & ![ITemID], ![ITemID])
                ITM.ForeColor = ListColor
                
                For i = 1 To 4
                    With ITM.ListSubItems.add(, , "")
                     .ForeColor = ListColor
                    End With
                Next i
                
                ITM.ListSubItems(1) = ![ItemName] & ""
                ITM.ListSubItems(2) = ![Price] & ""
                ITM.ListSubItems(3) = ![Qty] & ""
                ITM.ListSubItems(4) = ![Rcvd] & ""
                
                'Check For The Chq Status
                Dim Status As String, img As String
                If ![Qty] - ![Rcvd] = 0 Then
                  
                  img = "CLR"
                Else
                  
                  img = "DUE"
                
                End If
                 
                
                 ITM.SmallIcon = img
                 
                'Set Flag That it Was Payment Or Recept
                'If ![Payment] Then itm.Tag = "P" Else itm.Tag = "R"
                
                
                
                .MoveNEXT
                'LVItems.Visible = True
                'FDrag.Visible = True
            Loop
            
        Else
            'LVItems.Visible = False
            'FDrag.Visible = False
        End If
        .Close
    End With
    
    Set CrsLedger = Nothing
End Sub

Private Sub List1_Click()
If UCase(Screen.ActiveControl.Name) <> "CMBITEM" Then
   cmbItem = List1
   List1.Visible = False
   cmbItem.SetFocus
End If
End Sub

Private Sub Form_Resize()
    'FAcc.Left = ((Me.Width - FAcc.Width) / 2) - MainForm.ExpBar.Width
    'cmbMaterialID.Left = FAcc.Left
    'cmbMaterial.Left = cmbMaterialID.Left + cmbMaterialID.Width + 50
    'FLV.Move (Me.Width - FLV.Width - MainForm.ExpBar.Width) / 2, (Me.Height - FLV.Height) / 2
    'ImgLoading.Move (Me.ScaleWidth - ImgLoading.Width) / 2, (Me.ScaleHeight - ImgLoading.Height) / 2
End Sub

Private Sub LV_DblClick()
    If LV.SelectedItem Is Nothing Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(6).Left + 85, LV.Top + LV.SelectedItem.Top + 25, LV.ColumnHeaders(6).Width ', LV.SelectedItem.Height
        .Text = LV.SelectedItem.SubItems(5)
        Set .Font = LV.Font
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

'If Not LVItems.Visible Then Exit Sub
' If Val(ITEM.Tag) = 0 Then
'   LVItems.HideSelection = True
' Else
'    LVItems.HideSelection = False
'    LVItems.ListItems(Val(ITEM.Tag) & "'").Selected = True
' End If
 'Call refreshItemLedger
 
 
 
End Sub

Private Sub lvitems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'ITEM.Checked = Not ITEM.Checked
    
    
End Sub

Private Sub LVItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    refreshdetail
End Sub

Private Sub lvitems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    With LVItems
    
        If .ListItems.count = 0 Or Button <> 2 Then Exit Sub
        
        If .SelectedItem.ForeColor <> vbBlue Then Exit Sub
        
        If .SelectedItem.ListSubItems(5).Text = "Open" Then
            mnuToCash.Visible = True
        Else
            mnuToCash.Visible = False
        End If
        If .SelectedItem.Tag = "R" Then         ' Receipt Chq
         Me.PopUpMenu mnuPop
        ElseIf .SelectedItem.Tag = "P" Then         ' Payment Chq
          Me.PopUpMenu mnuClear
        End If
    End With
    
End Sub

Private Sub PrintLedger()
With Cr1
        

End With
End Sub

Private Sub refreshdetail()

    'Set LVItems.SmallIcons = ImageList1
    'Dim ListColor As Long
    Dim CrsLedger As New ADODB.Recordset
    
    With CrsLedger
        .CursorLocation = adUseClient
        .Open "Select EntryID,Dated,Quantity,ItemID,NetAmt from InOutDetail Where (OrderNo='" & LV.SelectedItem & "' and ItemID='" & LVItems.SelectedItem & "') ", con, adOpenForwardOnly '& _
            '"and (CDate between #" & DtFrom & "# and #" & DtTo & "#) " & _
            '"Order by CDate,Sno", con, 0, 1
            
        LVDetail.ListItems.Clear
        If .RecordCount > 0 Then
            Dim ITM As ListItem
            Do Until .EOF
                Set ITM = LVDetail.ListItems.add(, ![EntryID] & ![ITemID], ![EntryID])
                For i = 1 To 3
                    ITM.ListSubItems.add , , ""
                Next i
                ITM.ListSubItems(1) = Format(![Dated], "dddd dd MMMM ,yyyy")
                ITM.ListSubItems(2) = Abs(![Quantity])
                ITM.ListSubItems(3) = ![NetAmt]
                
                
                
                'Set Flag That it Was Payment Or Recept
                'If ![Payment] Then itm.Tag = "P" Else itm.Tag = "R"
                .MoveNEXT
                LVDetail.Visible = True
                'FDrag.Visible = True
            Loop
            
        Else
            'LVItems.Visible = False
            'FDrag.Visible = False
        End If
        .Close
    End With
    
    Set CrsLedger = Nothing
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    Dim a As Integer
    If KeyAscii = 13 Then
        con.Execute "UPDATE ItemLedgerSimple SET location='" & txtEdit & "' WHERE EntryID=" & Val(LV.SelectedItem.key), a
        If Val(LV.SelectedItem.SubItems(2)) > 0 Then    'Receving
            con.Execute "UPDATE RcvItemsSimpleDetail SET RcvItemsSimpleDetail.Location='" & txtEdit & "' FROM ItemLedgerSimple " & _
             "INNER JOIN RcvItemsSimpleDetail ON ItemLedgerSimple.EntryID=RcvItemsSimpleDetail.LedgerRefID " & _
             "WHERE ItemLedgerSimple.EntryID=" & Val(LV.SelectedItem.key)
        Else    'Issuance
            con.Execute "UPDATE IssItemsSimpleDetail SET IssItemsSimpleDetail.Location='" & txtEdit & "' FROM ItemLedgerSimple " & _
             "INNER JOIN IssItemsSimpleDetail ON ItemLedgerSimple.EntryID=IssItemsSimpleDetail.LedgerRefID " & _
             "WHERE ItemLedgerSimple.EntryID=" & Val(LV.SelectedItem.key)
        End If
        
        If a = 1 Then LV.SelectedItem.SubItems(5) = txtEdit
        LV.SetFocus
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Public Sub ShowMe()
    
    Me.Show
    
    'ImgLoading.Visible = True
    Dim f As New frmWaitAnimation
    f.Show
    DoEvents
    '-----------------------------------------------------------------
     
    '-----------------------------------------------------------------
    cmbMaterial.DropDownOffLine = True
    cmbMaterial.AddVals con, "'{' + ItemID + '} ' + ItemName", "Items", "ItemID", " ORDER BY ItemID"
    
    'ImgLoading.Visible = False
    Unload f
    
End Sub
