VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendRcvItems_Lots 
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11880
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
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   9870
      Left            =   105
      TabIndex        =   5
      Top             =   45
      Width           =   11685
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00B54646&
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Top             =   945
         Width           =   1395
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear All"
         Height          =   255
         Left            =   5625
         TabIndex        =   20
         Top             =   7815
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdCheckAll 
         Caption         =   "Check &All"
         Height          =   255
         Left            =   5610
         TabIndex        =   19
         Top             =   7575
         Visible         =   0   'False
         Width           =   975
      End
      Begin ComboList.Usercontrol1 cmbVend 
         Height          =   285
         Left            =   5445
         TabIndex        =   11
         Top             =   945
         Width           =   4695
         _ExtentX        =   8281
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
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   1530
         TabIndex        =   12
         Top             =   945
         Width           =   3915
         _ExtentX        =   6906
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
         Height          =   645
         Left            =   0
         TabIndex        =   8
         Top             =   75
         Width           =   11685
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Receive Items"
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
            Index           =   0
            Left            =   60
            TabIndex        =   9
            Top             =   165
            Width           =   11565
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Receive Items"
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
            Index           =   1
            Left            =   75
            TabIndex        =   10
            Top             =   180
            Width           =   11550
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   10155
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   945
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   148963331
         CurrentDate     =   37250
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   8505
         Left            =   120
         TabIndex        =   6
         Top             =   1275
         Width           =   11415
         Begin VB.CheckBox chkPrintPTC 
            Caption         =   "Print PTC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            TabIndex        =   33
            Top             =   7875
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.TextBox txtMaterialCertNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00B54646&
            Height          =   285
            Left            =   1740
            TabIndex        =   32
            Top             =   7395
            Width           =   2325
         End
         Begin VB.TextBox txtTemperValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00B54646&
            Height          =   285
            Left            =   4095
            TabIndex        =   31
            Top             =   7395
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.Frame FraEmp 
            Height          =   3795
            Left            =   75
            TabIndex        =   27
            Top             =   3360
            Width           =   11325
            Begin VB.CheckBox chkOverTime 
               Caption         =   "Over Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Left            =   60
               TabIndex        =   30
               Top             =   3465
               Width           =   1485
            End
            Begin ComboList.Usercontrol1 cmbEmp 
               Height          =   285
               Left            =   60
               TabIndex        =   1
               Top             =   465
               Width           =   10245
               _ExtentX        =   18071
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
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   10350
               TabIndex        =   2
               Top             =   405
               Width           =   870
            End
            Begin MSComctlLib.ListView LVEmp 
               Height          =   2610
               Left            =   60
               TabIndex        =   29
               Top             =   795
               Width           =   11160
               _ExtentX        =   19685
               _ExtentY        =   4604
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   8388608
               BackColor       =   15199215
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Employee"
                  Object.Width           =   17639
               EndProperty
            End
            Begin VB.Label lblEmp 
               AutoSize        =   -1  'True
               Caption         =   "Employee :"
               Height          =   195
               Left            =   60
               TabIndex        =   28
               Top             =   225
               Width           =   795
            End
         End
         Begin VB.CheckBox chkPrint 
            Caption         =   "Print When Saved"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            TabIndex        =   24
            Top             =   8235
            Width           =   1920
         End
         Begin VB.CheckBox chkPrintDirect 
            Caption         =   "Print Directly to Printer"
            ForeColor       =   &H00000040&
            Height          =   315
            Left            =   4095
            TabIndex        =   23
            Top             =   8160
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.ComboBox cmbCountedBy 
            Height          =   315
            Left            =   2430
            TabIndex        =   22
            Text            =   "cmbCountedBy"
            Top             =   2370
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtCountedBy 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2535
            TabIndex        =   21
            Top             =   735
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   6150
            TabIndex        =   18
            Top             =   7890
            Width           =   5280
            Begin VB.CommandButton cmdSave 
               Caption         =   "Save && &New   "
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
               Left            =   60
               TabIndex        =   3
               Top             =   165
               Width           =   1665
            End
            Begin VB.CommandButton cmdSaveClose 
               Caption         =   "&Save && Close"
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
               Left            =   1755
               TabIndex        =   14
               Top             =   165
               Width           =   1665
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Cancel"
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
               Left            =   3480
               TabIndex        =   15
               Top             =   165
               Width           =   1665
            End
         End
         Begin VB.PictureBox Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1350
            ScaleHeight     =   255
            ScaleWidth      =   945
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1110
            Visible         =   0   'False
            Width           =   945
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Width           =   950
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   3225
            Left            =   75
            TabIndex        =   13
            Top             =   0
            Width           =   11325
            _ExtentX        =   19976
            _ExtentY        =   5689
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   15199215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Priority"
               Object.Width           =   1296
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Item No."
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Name"
               Object.Width           =   4233
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Size"
               Object.Width           =   1191
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Type"
               Object.Width           =   1799
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Color"
               Object.Width           =   1905
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Rcvd. Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Weight"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Ret. Date"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Counted By"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Text            =   "Rcv ID."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Order No."
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Label lblTemperValue 
            AutoSize        =   -1  'True
            Caption         =   "Temper Value :"
            Height          =   195
            Left            =   4065
            TabIndex        =   26
            Top             =   7185
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lblMatCertNo 
            AutoSize        =   -1  'True
            Caption         =   "Mil No."
            Height          =   195
            Left            =   1755
            TabIndex        =   25
            Top             =   7200
            Width           =   480
         End
         Begin VB.Image imgItem 
            BorderStyle     =   1  'Fixed Single
            Height          =   1275
            Left            =   75
            Stretch         =   -1  'True
            Top             =   7185
            Width           =   1620
         End
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   735
         Width           =   11325
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmVendRcvItems_Lots.frx":0000
         Size            =   "19976;397"
         BorderColor     =   10860487
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop up"
      Begin VB.Menu mnuShowPicture 
         Caption         =   "Show Picture"
      End
   End
End
Attribute VB_Name = "frmVendRcvItems_Lots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dWastage As Double
Dim dRcvdQty As Double
Dim Col_OrdQty As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean
Dim InitLotProcess As Boolean
Dim lLotNo As String
Dim GotRcvID As Long
Dim RcvingLot As Boolean
Dim UserReqAuth As Integer
Dim iColNo As Integer
Dim bByLotNo As Boolean

Private Sub cmbCountedBy_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
'        If cmbCountedBy.ListIndex = -1 Then
'            MsgBox "Invlaid Name.", vbInformation
'            Exit Sub
'        End If
        LV.SelectedItem.SubItems(iColNo - 1) = cmbCountedBy.Text
        cmbCountedBy.Visible = False
    ElseIf KeyAscii = 27 Then
        cmbCountedBy.Visible = False
    End If
    
End Sub

Private Sub cmbProcess_matched()
   
    On Error GoTo err
'    If cmbProcess.ID = 9 Then
'        cmdCheckAll.Visible = True
'        cmdClearAll.Visible = True
'    Else
'        cmdCheckAll.Visible = False
'        cmdClearAll.Visible = False
'    End If
    
    Dim cmd As New ADODB.Command
    
    Call cmbVend.ClearVals
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_ReqAuthForRcv"
        Set .ActiveConnection = con
        .Parameters("@UserID").Value = UserID
        .Parameters("@ProcessID").Value = cmbProcess.ID
        .Execute
        UserReqAuth = Val(.Parameters("@ReqAuth").Value & "")
        
        '.CommandText = "SP_VendersForRcving"
        '.Parameters("@ProcessID").value = cmbProcess.ID
        '.Parameters("@HicoVisible").value = HicoVisible
        
        'Set rs = .Execute
    End With
    
    If RcvLot(cmbProcess.ID) Then
        RcvingLot = True
    Else
        RcvingLot = False
    End If
    
    LV.ListItems.Clear
    GotRcvID = 0
    
    
    'If HicoVisible Then
    '    Call cmbVend.AddVals(con, "VenderName", "VVendIssItems", "VendID", " Where PhaseID=" & cmbProcess.ID & " Or PhaseID=-1")
    'Else
    '    Call cmbVend.AddVals(con, "VenderName", "VVendIssItems", "VendID", " Where PhaseID=" & cmbProcess.ID)
    'End If
    
    
'    If HicoVisible Then
'        If cmbProcess.ID = 7 Then
'            Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1 ) And VendID In(Select VendID From VMakerIssItems Where ProcessID=" & cmbProcess.ID & "  Group By VendID Having Sum(RemQty)>0)")
'        Else
'            Call cmbVend.AddVals(con, "VenderName", "Makers", "VendID", " Where (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1) And VendID In(Select VendID From VMakerIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
'        End If
'    Else
'        If cmbProcess.ID = 7 Then
'            Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where PhaseID=" & cmbProcess.ID & "  And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
'        Else
'            Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where PhaseID=" & cmbProcess.ID & " And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
'        End If
'    End If
    
    'Call cmbVend.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID In(SELECT VendID From VMakerIssItems WHERE RcvProcessID=" & cmbProcess.ID & " GROUP BY VendID HAVING SUM(RemQty)>0)")
    
    Dim lSNo1 As Long, lSNo2 As Long
    'lSNo1 = Val(con.Execute("Select SNo From VInitialLotProcess").Fields(0).value & "")
    
    lSNo2 = Val(con.Execute("SELECT SNo FROM Processes WHERE ProcessID=" & cmbProcess.ID).Fields(0).Value & "")
    
'    If lSNo2 >= lSNo1 Then
'        cmbProcess.Tag = "Lot No"
'    Else
'        cmbProcess.Tag = ""
'    End If
    'cmbProcess.Tag = ""
    'Set Focus To CmbVenders And Also Drop It Down!
'    cmbVend.SetFocus

    'SendKeys "a"
    'SendKeys Chr(8)
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmbVend_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    
    InitLotProcess = False
    'If LV.ColumnHeaders.count = 8 Then LV.ColumnHeaders.Remove "OrderNo"
    LV.Checkboxes = False
    Dim dExtraPer As Double, lExtraQty As Long
    With rs
        If Not (strCompany = "Weldon Instruments" Or strCompany = "QEL") Or bByLotNo Then
            'SQL = "SELECT EntryID,OrderNo,ItemCode,ItemSize,SizeUnit,Type,ItemColor,ItemNameUrdu,ItemName,IssQty-RcvdQty-ISNULL(IssRcvBalance,0)-UnAuthRcvdQty+ISNULL(Wastage_Return,0) AS Quantity,RcvProcessID,ReturnDT,Priority,ForgingWeight,Unit,IssQty,RcvdQty,UnAuthRcvdQty,ExcessQtyPercentage,LotNo,ReWorkLot,Repair_RefID FROM VMakerIssItems WHERE IssQty-RcvdQty-ISNULL(IssRcvBalance,0)-UnAuthRcvdQty+ISNULL(Wastage_Return,0)>0 AND RefID=" & txtID.Tag & " ORDER BY Priority"
            SQL = "SELECT EntryID,OrderNo,ItemCode,ItemSize,SizeUnit,Type,ItemColor,ItemNameUrdu,ItemName,IssQty-RcvdQty-UnAuthRcvdQty+CASE WHEN IssRcvBalance<0 THEN ABS(IssRcvBalance) ELSE 0 END+ISNULL(Wastage_Return,0) AS Quantity,RcvProcessID,ReturnDT,Priority,ForgingWeight,Unit,IssQty,RcvdQty,UnAuthRcvdQty,ExcessQtyPercentage,LotNo,ReWorkLot,Repair_RefID FROM VMakerIssItems WHERE IssQty-RcvdQty-UnAuthRcvdQty+CASE WHEN IssRcvBalance<0 THEN ABS(IssRcvBalance) ELSE 0 END+ISNULL(Wastage_Return,0)>0 AND RefID=" & txtLotNo.Tag & " ORDER BY Priority"
        Else
            SQL = "SELECT EntryID,OrderNo,ItemCode,ItemSize,SizeUnit,Type,ItemColor,ItemNameUrdu,ItemName,IssQty+ISNULL(Wastage_Return,0)-RcvdQty-UnAuthRcvdQty AS Quantity,RcvProcessID,ReturnDT,Priority,ForgingWeight,Unit,IssQty,RcvdQty,UnAuthRcvdQty,ExcessQtyPercentage,LotNo,ReWorkLot,Repair_RefID FROM VMakerIssItems WHERE RefID=" & txtID.Tag & " ORDER BY Priority"
        End If
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        '.Filter = "Quantity<>0"
        LV.ListItems.Clear
        GotRcvID = 0
        Dim FColor As Long
        Dim lTotalQty As Long
        Do Until .EOF
            FColor = LV.ForeColor
            Set ITM = LV.ListItems.add(, , Val(!Priority & ""))
            ITM.ForeColor = FColor
            ITM.ListSubItems.add(, , ![ItemCode] & "").ForeColor = FColor
            ITM.ListSubItems.add(, , !ItemName & "").ForeColor = FColor
            
            dExtraPer = Val(!ExcessQtyPercentage & "")
            lExtraQty = (Val(!IssQty & "") / 100) * dExtraPer
            
            ITM.ListSubItems.add(, , ![ItemSize] & " " & !SizeUnit & "").ForeColor = FColor
            ITM.ListSubItems.add(, , !Type & "").ForeColor = FColor
            ITM.ListSubItems.add(, , !ItemColor & "").ForeColor = FColor
            ITM.ListSubItems(1).Tag = !OrderNo & ""
            ITM.ListSubItems(2).Tag = Val(![Quantity] & "") + lExtraQty
            ITM.ListSubItems(4).Bold = True
            ITM.ListSubItems(3).Tag = Val(!EntryID & "")
            ITM.ListSubItems(4).Tag = Val(!RcvProcessID & "")
            
            If bByLotNo Then
                ITM.ListSubItems.add , , Val(!IssQty & "") 'Rcvd Qty...
            Else
                ITM.ListSubItems.add , , ""
            End If
            
            ITM.ListSubItems.add        'Rcvd Weight...
            ITM.ListSubItems.add , , Format(!ReturnDT & "", "dd-MM-yy")
            ITM.ListSubItems.add    'Counted By
            ITM.ListSubItems.add().Tag = Val(!IssQty & "")  'RcvID ... Hidden Column
            ITM.ListSubItems.add , , !OrderNo & ""
            
            'ITM.ListSubItems.add , , DateDiff("d", !ReturnDT, Date) & " days"
            ITM.ListSubItems(5).Tag = Val(!ForgingWeight & "")
            ITM.ListSubItems(6).Tag = !Unit & ""
            ITM.ListSubItems(7).Tag = (!LotNo & "")
            ITM.ListSubItems(8).Tag = !ReWorklot
            ITM.ListSubItems(9).Tag = Val(!Repair_RefID & "")
            lTotalQty = lTotalQty + Val(ITM.ListSubItems(2).Tag)
            .MoveNEXT
        Loop
        .Close
        LV.Tag = lTotalQty
        If LV.ListItems.count > 0 Then
            Dim strItemCode As String
            strItemCode = LV.ListItems(LV.ListItems.count).SubItems(1) 'LV.ListItems(LV.ListItems.count).SubItems(2)
            Dim rsPic As New ADODB.Recordset
            rsPic.Open "SELECT ItemID,ItemPic FROM Items WHERE ItemID='" & strItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(rsPic!ItemPic) Then
                Set imgItem.DataSource = rsPic
                imgItem.DataField = "ItemPic"
            End If
            rsPic.Close
            
            If Val(ITM.ListSubItems(7).Tag) = 0 Then
                LV.ListItems(1).Selected = True
                Call LV_DblClick
            End If
        End If
    End With
    
    Set rs = Nothing
    
    Dim myDT As String
    If cmbProcess.Tag = "Lot No" Then
        If LV.ListItems.count > 0 Then
            myDT = LV.ListItems(1).Tag
            For i = 1 To LV.ListItems.count
                Set ITM = LV.ListItems(i)
                If myDT = ITM.Tag Then
                    'Set The ForeColor Of all The ListSubItems To Red.
                    ITM.ForeColor = vbRed
                    For j = 1 To ITM.ListSubItems.count
                        ITM.ListSubItems(j).ForeColor = vbRed
                    Next j
                End If
            Next i
        End If
    End If
    
    If cmbVend.ID = iFactoryMakerID And cmbVend.MatchFound Then
        Call ShowEmp
    Else
        Call HideEmp
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub HideEmp()
     
    LVEmp.ListItems.Clear
    FraEmp.Enabled = False
    'chkOverTime.Value = vbUnchecked
'
'    lblEmp.Visible = False
'    cmbEmp.ID = ""
'    cmbEmp.Visible = False
'    chkOverTime.Visible = False
    
End Sub

Private Sub ShowEmp()
    
    LVEmp.ListItems.Clear
    FraEmp.Enabled = True
    'chkOverTime.Value = vbUnchecked
    
    If cmbEmp.Tag = "" Then
        cmbEmp.ClearVals
        cmbEmp.ListHeight = 2300
        cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name + ' {' + Designation + '} ' + '{' + DeptName + '}'", "VEmp", "EmpID", " WHERE Active=1"
    End If
    cmbEmp.Tag = "Added"
    
End Sub

Private Sub cmdSClear_Click()

    If Saved Then
        FormProcessID = cmbProcess.ID
        Unload Me
        Call ShowRcvVForm(FormProcessID)
    End If
    'Call ClearFields

End Sub

Private Sub cmdSClose_Click()

    'If Not AskForAutoApply Then Exit Sub

    If Saved Then
        Unload Me
    End If

End Sub

Private Sub cmdAdd_Click()
    
    If cmbEmp.MatchFound = False Then
        MsgBox "Please Select Employee.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To LVEmp.ListItems.count
        If cmbEmp.ID = LVEmp.ListItems(i).key Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        End If
    Next
     
    LVEmp.ListItems.add , cmbEmp.ID, cmbEmp.Text
    cmbEmp.ID = ""
    cmbEmp.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    con.Execute "Delete From PrintSession"
End Sub


Private Sub cmdCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
        Call LV_ItemCheck(LV.ListItems(i))
    Next
End Sub

Private Sub cmdClearAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdSave_Click()
    
    Dim lEntryID As Long
    
    If Saved(lEntryID) Then
        
        LV.ListItems.Clear
        cmbVend.ClearVals
        txtLotNo = ""
'        txtLotNo.SetFocus
        Call HideEmp
        
        FormProcessID = cmbProcess.ID
        'Unload Me
        'Call ShowRcvVForm(FormProcessID)
        'Call PrintRcvingSlips(lEntryID)
        'Call PrintRcvingSlips(Val(txtID.Tag))
        If chkPrint.Value = vbChecked Then
            Call PrintRcvingSlips(Val(txtID.Tag))
            DoEvents
        End If
        
'        Call cmbVend_matched
        
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    Dim lEntryID As Long
    If Saved(lEntryID) Then
'        If cmbProcess.ID <= 7 Then
'            If Not UserReqAuth Then Call PrintIssRcvRpts("Original", False)
'        End If
        'Call PrintRcvingSlips(lEntryID)
        If chkPrint.Value = vbChecked Then
            Call PrintRcvingSlips(Val(txtID.Tag))
            DoEvents
        End If
                
        If chkPrintPTC.Value = vbChecked Then
            Call PrintPTC(lLotNo)
            DoEvents
        End If
        
        Unload Me
        'con.Execute "Delete From PrintSession"
    End If
    
End Sub

Private Sub PrintRcvingSlips(lEntryID As Long)

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlipWithRcving.rpt")
    
    strSelection = "{VendIssued.EntryID}=" & lEntryID
     
    If chkPrintDirect.Value = vbChecked Then
        Call PrintReport(strSelection, rpt)
    Else
        Dim f As New frmPrevRpt
        f.ShowReport strSelection, rpt
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Pic.Visible = False
    ElseIf KeyCode = 13 Then
        If Val(txtQty) = 0 Then
                If InitLotProcess Then
                    LV.SelectedItem.SubItems(5) = Val(txtQty)
                Else
                    LV.SelectedItem.SubItems(4) = Val(txtQty)
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = ""
                'Reset The IssID's
                Dim TempID As Long
                TempID = 0
                
                For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                    TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1), 3)) - 1
                    If TempID > 0 Then
                        LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                         Format(ServerDate, "ddMMyy") & TempID
                    End If
                Next i
                If GotRcvID > 0 Then GotRcvID = GotRcvID - 1
                
                
                
                Pic.Visible = False
                Exit Sub
            End If
            
            If InitLotProcess Then
                If Val(txtQty) > Val(LV.SelectedItem.SubItems(3)) Then
                    MsgBox "Invalid Quantities!", vbInformation
                    Exit Sub
                End If
            Else
                If Val(txtQty) > Val(LV.SelectedItem.SubItems(3)) Then
                    MsgBox "Invalid Quantities!", vbInformation
                    Exit Sub
                End If
            End If
            
            If Not txtQty.Locked Then
                If cmbProcess.Tag = "Lot No" Then
                    If lLotNo = 0 Then
                        lLotNo = Val(con.Execute("Select Max(Cast (LotNo As Int)) From VendRcvdDetail").Fields(0).Value & "") + 1
                    Else
                        lLotNo = lLotNo + 1
                    End If
                    
                    If InitLotProcess Then
                        LV.SelectedItem.ListSubItems(4).Tag = lLotNo
                    Else
                        LV.SelectedItem.ListSubItems(4).Tag = lLotNo
                    End If
                End If
            End If
            
            If InitLotProcess Then
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.Value, "dd-MMM-yyyy")
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.Value, "dd-MMM-yyyy")
            End If
            
            '-----------------------
            
            'Dim RcvdID As String
            
            If cmbProcess.Tag = "Lot No" Then
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).Value & "") + 1  'and UPPER(Left(RecieptID,3))<>'HSR'
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                Format(DTPicker2, "ddMMyy") & GotRcvID
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).Value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                  Format(DTPicker2, "ddMMyy") & GotRcvID
            End If
                
            '-----------------------
            
            'Call AdjustQty
            Pic.Visible = False
            
            If LV.SelectedItem.Index < LV.ListItems.count Then
                'If Val(LV.SelectedItem.SubItems(3)) = 0 Then
                 LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                'End If
                LV.SetFocus
            End If
            
        ElseIf KeyAscii = 27 Then
            Pic.Visible = False
        End If
End Sub

Private Sub Form_Activate()
    'Check For User Access
'    If PermissionDenied Then
'        Me.Hide
'        MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
'        Unload Me
'        Exit Sub
'    End If


End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    mnuPop.Visible = False
    'Me.KeyPreview = True
    DTPicker1.Value = ServerDate

    cmbProcess.Enabled = True
    cmbProcess.ListHeight = 2600
    cmbVend.ListHeight = 2600
    Call cmbProcess.AddVals(con, "Description", "Processes", "ProcessID", "Order By SNo")
    
    lLotNo = 0
    Call SaveLV(LV)
    
    GotRcvID = 0
    
    If strCompany = "Sunlike" Then
        DTPicker1.Enabled = True
    Else
        If UserID = 1 Then DTPicker1.Enabled = True
    End If
    
    'Call AddToCombo(cmbCountedBy, "UserName", "CountedByUsers", " ORDER BY UserName")
    Call AddToCombo(cmbCountedBy, "Name", "Employees", " ORDER BY Name")
    
    If strCompany = "Banzai" Then
        txtTemperValue.Visible = True
        lblTemperValue.Visible = True
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Public Sub LockProc(ProcID As Long)
    Call Form_Resize
    cmbProcess.ID = ProcID
    cmbProcess.Enabled = False
    Me.Show
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    
    Set Col_OrdQty = Nothing
    Set frmVendRcvItems = Nothing
End Sub

Private Sub Form_Resize()
   FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub



Private Sub imgItem_Click()
    
    If imgItem.Picture = 0 Then Exit Sub
    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromPic imgItem.Picture
    
End Sub

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    If txtLotNo <> "" Then Exit Sub
    'If LV.SelectedItem.ListSubItems(7).Tag > 0 Then Exit Sub
    
    Dim i As Integer
    If iColNo = 10 Then
        With cmbCountedBy
            .Move LV.Left + 10 + LV.ColumnHeaders(iColNo).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(iColNo).Width
            
            For i = 0 To cmbCountedBy.ListCount - 1
                If cmbCountedBy.List(i) = LV.SelectedItem.SubItems(iColNo - 1) Then
                    cmbCountedBy.ListIndex = i
                    Exit For
                End If
            Next
            .Visible = True
            .SetFocus
        End With
        Exit Sub
    ElseIf iColNo = 7 Then
        
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                If Val(.ListSubItems(2).Tag) > Val(.SubItems(6)) Then
                    .Selected = True
                    Exit For
                End If
            End With
        Next
    End If
   
    If iColNo <> 8 Or LV.SelectedItem.ListSubItems(6).Tag <> "Kgs" Then iColNo = 7
    
    With Pic
        .Move LV.Left + 10 + LV.ColumnHeaders(iColNo).Left, LV.SelectedItem.Top + LV.Top
        .Visible = True
    End With
    
    With txtQty
        If Val(LV.SelectedItem.SubItems(iColNo - 1)) > 0 Then
            .Tag = "Simple"
        Else
            .Tag = ""
        End If
        .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.SubItems(4) = Item.SubItems(3)
    Else
        Item.SubItems(4) = 0
    End If
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LV_DblClick
    End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > LV.ColumnHeaders(i).Left And X < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    Me.PopUpMenu mnuPop
   ' Exit Sub
   ' If LV.ListItems.Count = 0 Then Exit Sub
    'If Val(LV.SelectedItem.SubItems(6)) = 0 Then Exit Sub
    'Dim iCol As Integer
    'iCol = GetColNo(LV, X)
    
    'If iCol = 8 Then
    '    Load frmVendRcvItems1
    '    LV.SelectedItem.ListSubItems(7).Text = frmVendRcvItems1.FillList(Col_OrdQty, LV.SelectedItem.SubItems(1))
    '    LV.SelectedItem.ListSubItems(7).Tag = LV.SelectedItem.ListSubItems(7).Text
    'End If
    
End Sub

Private Sub txtQtyOld_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
        
        If Val(txtQty) = 0 Then
            If InitLotProcess Then
                LV.SelectedItem.SubItems(5) = Val(txtQty)
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
            End If
            LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = ""
            'Reset The IssID's
            Dim TempID As Long
            TempID = 0
            
            For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1), 3)) - 1
                If TempID > 0 Then
                    LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                     Format(ServerDate, "ddMMyy") & "-" & Format(TempID, "000")
                End If
            Next i
            If GotRcvID > 0 Then GotRcvID = GotRcvID - 1
            
            
            
            Pic.Visible = False
            Exit Sub
        End If
        
        If InitLotProcess Then
            If Val(txtQty) > Val(LV.SelectedItem.SubItems(4)) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
        Else
            If Val(txtQty) > Val(LV.SelectedItem.SubItems(3)) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
        End If
        
        If Not txtQty.Locked Then
            If cmbProcess.Tag = "Lot No" Then
                If lLotNo = 0 Then
                    lLotNo = Val(con.Execute("Select Max(Cast(LotNo As int)) From VendRcvdDetail").Fields(0).Value & "") + 1
                Else
                    lLotNo = lLotNo + 1
                End If
                
                If InitLotProcess Then
                    LV.SelectedItem.ListSubItems(4).Tag = lLotNo
                Else
                    LV.SelectedItem.ListSubItems(4).Tag = lLotNo
                End If
            End If
        End If
        If InitLotProcess Then
            LV.SelectedItem.SubItems(5) = Val(txtQty)
        Else
            LV.SelectedItem.SubItems(4) = Val(txtQty)
        End If
        
        '-----------------------
        
        'Dim RcvdID As String
        
            If cmbProcess.Tag = "Lot No" Then
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).Value & "") + 1  'and UPPER(Left(RecieptID,3))<>'HSR'
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).Value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
            End If
            
        '-----------------------
        
        'Call AdjustQty
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, False)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub AdjustQty()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim lQty As Long, TotQty As Long, WasteQty As Long
    Dim TempQty As Long, TempWasteQty As Long
    Dim strDetail As String
    
    TempQty = 0
    TotQty = Val(LV.SelectedItem.SubItems(6))
    WasteQty = Val(LV.SelectedItem.SubItems(5))
    
    With rs
        .Open "Select OrderNo,ItemCode,TotQty From VTotIssdByOrder Where ItemCode='" & LV.SelectedItem.SubItems(1) & "'", con, adOpenForwardOnly
        
        Do Until .EOF
            Dim Obj As New RcvAdjOrderQty
            Obj.OrderNo = ![OrderNo] & ""
            Obj.ItemCode = ![ItemCode] & ""
            Obj.lOrdQty = ![TotQty] & ""
            Obj.lWastage = 0
            Obj.lQty = 0
            
            Col_OrdQty.add Obj, ![OrderNo] & "|" & ![ItemCode]
            Set Obj = Nothing
            .MoveNEXT
        Loop
        .MoveFirst
        
        Do While lQty < TotQty
            
            TempQty = TotQty - lQty
            TempWasteQty = 0
            If TempQty > ![TotQty] Then
                TempQty = ![TotQty]
            End If
            
            If WasteQty > 0 Then
                If TempQty > WasteQty Then
                    TempQty = TempQty - WasteQty
                    TempWasteQty = WasteQty
                    WasteQty = 0
                Else
                    TempWasteQty = WasteQty
                    TempQty = 0
                    WasteQty = WasteQty - TempQty
                End If
            End If
            
            
            Col_OrdQty(![OrderNo] & "|" & ![ItemCode]).lQty = TempQty
            Col_OrdQty(![OrderNo] & "|" & ![ItemCode]).lWastage = TempWasteQty
            
            lQty = lQty + TempQty + TempWasteQty
            
            strDetail = strDetail & ![OrderNo] & "=" & TempQty + TempWasteQty & ","
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    strDetail = Left(strDetail, Len(strDetail) - 1)
    LV.SelectedItem.ListSubItems(7) = strDetail
    LV.SelectedItem.ListSubItems(7).ToolTipText = strDetail
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Function OLDSaved() As Boolean

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Col As New Collection
    
    Dim cmd As New ADODB.Command
    
    Dim iRefID As Long
    
    Call StartTrans(con)
    
    con.Execute "Insert Into VendReceived(DT,RecieptNo,VendID) Values('" & DTPicker1 & "','" & txtRecNo & "'," & cmbVend.ID & ")"
    
    iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).Value & "")
    
    'Get The Next Process ID For Current Item...
    For i = 1 To LV.ListItems.count
        
            If Val(LV.ListItems(i).SubItems(6)) > 0 Then
                Dim iNextProcID As Integer
                'Dim cmd As New ADODB.Command
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_GetNextProcID"
                    .Parameters("@ItemCode").Value = LV.ListItems(i).SubItems(1)
                    .Parameters("@ProcID").Value = cmbProcess.ID
                    .Execute
                    iNextProcID = Val(.Parameters("@NextProcID").Value & "")
                End With
                
                For j = 1 To Col_OrdQty.count
                    With Col_OrdQty(j)
                        If .ItemCode = LV.ListItems(i).SubItems(1) And .lQty + .lWastage > 0 Then
                        
                            con.Execute "Insert Into VendRcvdDetail(RefID,OrderNo,ItemCode,NextProcessID,ProcessID,RcvdQty,Wastage) Values(" & _
                            iRefID & ",'" & .OrderNo & "','" & .ItemCode & "'," & iNextProcID & "," & cmbProcess.ID & "," & .lQty & "," & .lWastage & ")"
                            
                            '2. Now If Rcvd After Issued Then
                            'Update The RcvdQty In VendIssdDetails....
                            '-----------------------------------------------------------------
                            With cmd
                                .ActiveConnection = con
                                .CommandType = adCmdStoredProc
                                .CommandText = "SP_UpdateIssued"
                                .Parameters("@OrderNo").Value = Col_OrdQty(j).OrderNo
                                .Parameters("@ItemCode").Value = Col_OrdQty(j).ItemCode
                                .Parameters("@ProcessID").Value = cmbProcess.ID
                                .Parameters("@Qty").Value = Col_OrdQty(j).lQty + Col_OrdQty(j).lWastage
                                .Execute
                            End With
                            '-----------------------------------------------------------------
                                
                                                    
                            'con.Execute "Update VendIssdDetail Set RcvdQty=RcvdQty+" & .lQty + .lWastage & " Where OrderNo='" & .OrderNo & "' and ItemCode='" & .ItemCode & "'"
                        End If
                    End With
                Next j
                
                'For j = 1 To Col.Count
                '    Col.Remove 1
                'Next
                
                
                
                
                '---------------------------------------------------------------------------------------------------------------------------------------
                'Update RcvdQty In The VendIssdDetail Table....
                
                'With rs
                '    .Open "Select EntryID,IssQty-RcvdQty From VendIssdDetail Where ItemCode='" & LV.ListItems(i).SubItems(1) & "' and IssQty-RcvdQty>0 and ProcessID=" & cmbProcess.ID & " Order By EntryID", con, adOpenForwardOnly, adLockReadOnly
                    
                '    Do Until .EOF
                '        Dim Obj As New AdjIssQty
                '        Obj.EntryID = ![EntryID]
                '        Obj.lQty = .Fields(1)
                '        Col.Add Obj
                '        Set Obj = Nothing
                '        .MoveNext
                '    Loop
                '    .Close
                'End With
                
                'Dim TempQty As Long, TotQty As Long
                'Dim TempTot As Long
                'TotQty = Val(LV.ListItems(i).SubItems(6))
                'TempQty = 0
                'TempTotQty = 0
                
                'For j = 1 To Col.Count
                
                '    If TotQty - TempTot >= Col(j).lQty Then
                '        TempQty = Col(j).lQty
                '    Else
                '        TempQty = TotQty - TempTot
                '    End If
                '    TempTot = TempTot + TempQty
                '
                '    con.Execute "Update VendIssdDetail Set RcvdQty=" & TempQty & " Where EntryID=" & Col(j).EntryID
                '
                '    If TempTot >= TotQty Then Exit For
                '
                'Next j
                '---------------------------------------------------------------------------------------------------------------------------------------
            End If
        
    Next i
    
    Set rs = Nothing
    con.CommitTrans
    OLDSaved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function
    
Private Function Saved(Optional ByRef lEntryID As Long = 0) As Boolean

    On Error GoTo err
    If cmbVend.ID = iFactoryMakerID Then
        If LVEmp.ListItems.count = 0 Then
            MsgBox "Please Select atleast 1 Employee.", vbInformation
            Exit Function
        End If
    End If
    
'    If bByLotNo = False And strCompany = "IAA" Then
'        If getDBPassword("MasterPO_Receiving") = False Then Exit Function
'    End If
    
    Dim i As Integer
    Dim dForgingWeight As Double, d10Per As Double
     
'    For i = 1 To LV.ListItems.count
'        If Val(LV.ListItems(i).SubItems(6)) > 0 And LV.ListItems(i).ListSubItems(6).Tag = "Kgs" Then
'            dForgingWeight = Val(LV.ListItems(i).ListSubItems(5).Tag) * Val(LV.ListItems(i).SubItems(6)) / 1000 'Forging Weight X Rcvd Qty
'            '10% ...
'            d10Per = (dForgingWeight / 100) * 10
'            If Not (Val(LV.ListItems(i).SubItems(7)) > (dForgingWeight - d10Per) And Val(LV.ListItems(i).SubItems(7)) < dForgingWeight + d10Per) Then
'                MsgBox "Weight is not withing 10 Percent Range Limit." & vbNewLine & "It should be between " & dForgingWeight - d10Per & " AND " & dForgingWeight + d10Per & " Kgs." & vbNewLine & "Weight Entry should be in Kgs.", vbInformation
'                Exit Function
'            End If
'        End If
'        If (strCompany = "Dr-Frgz") And Val(LV.ListItems(i).ListSubItems(7).Tag) = 0 And txtMaterialCertNo = "" Then
'            MsgBox "Please Enter Material Cert. No.", vbInformation
'            Exit Function
'        End If
'    Next
    Dim iAuthRequired As Integer
    If GetSingleStringValue("Description", "Processes", " WHERE ProcessID=" & cmbProcess.ID & " AND AuthRequired=1") <> "" Then
        iAuthRequired = 1
    Else
        iAuthRequired = 0
    End If
    
    If LV.ListItems.count = 0 Then Exit Function
    Dim rs As New ADODB.Recordset
    Dim Col As New Collection
    
    Dim cmd As New ADODB.Command
    Dim Qty As Long
    
    Dim iRefID As Long
    Dim iIssRefID As Long
    
'    Dim strEmpID As String
'    If cmbEmp.Visible Then
'        strEmpID = cmbEmp.ID
'    Else
'        If cmbEmp.MatchFound Then
'            strEmpID = cmbEmp.ID
'        Else
'            strEmpID = cmbEmp.ID
'        End If
'    End If
        
    Call StartTrans(con)
    
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID,OverTime,UserName,MachineName,TemperValue) Values(" & cmbVend.ID & _
     ",'" & DTPicker1 & " " & ServerTime & "',''," & UserID & "," & cmbProcess.ID & "," & Val(txtLotNo.Tag) & "," & chkOverTime.Value & ",'" & CurrentUserName & "','" & strComputerName & "','" & txtTemperValue & "')"
    
    'iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).Value & "")
    iRefID = GetSingleLongValue("MAX(EntryID)", "VendReceived", " WHERE MachineName='" & strComputerName & "'")
    
    For i = 1 To LVEmp.ListItems.count
        con.Execute "INSERT INTO VendReceived_Employees(VR_RefID,EmpID) VALUES(" & iRefID & ",'" & LVEmp.ListItems(i).key & "')"
    Next
    
    'Get The Next Process ID For Current Item...
    lEntryID = iRefID
    Dim lVI_EntryID As Long, lMLS_EntryID As Long, strBatchNo As String, strMil_No As String
    Dim frmViewBookMark As frmAddBookmarkforReceiving
    
    Dim lQtyToUse As Long, lQtyUsed As Long, lLotQty As Long, lLotBuffer As Long
    Dim lEntryIDDetail As Long
    For i = 1 To LV.ListItems.count
        lLotQty = GetSingleLongValue("ItemMaxLotSize", "Items", " WHERE ItemID='" & LV.ListItems(i).SubItems(1) & "'")
        lLotBuffer = GetSingleLongValue("ItemLotSizeBuffer", "Items", " WHERE ItemID='" & LV.ListItems(i).SubItems(1) & "'")
        lQtyUsed = 0
        Qty = Val(LV.ListItems(i).SubItems(6))
        If lLotQty = 0 Or lLotBuffer = 0 Then
            MsgBox "Lot Qty or Lot's Buffer Qty is Zero.", vbInformation
            Exit Function
        End If
        While lQtyUsed < Qty
            
            If Val(LV.ListItems(i).ListSubItems(7).Tag) > 0 Then
                lQtyToUse = Qty
            ElseIf (Qty - lQtyUsed) <= lLotQty Then
                lQtyToUse = (Qty - lQtyUsed)
            Else
                If (Qty - lQtyUsed) <= (lLotQty + lLotBuffer) Then
                    lQtyToUse = (Qty - lQtyUsed)
                Else
                    lQtyToUse = lLotQty
                End If
            End If
           
            If lQtyToUse > 0 Then
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                
                    .CommandText = "SP_InsertVendRcvdNew"
                    .Parameters("@ItemCode").Value = LV.ListItems(i).SubItems(1)
                    .Parameters("@ProcessID").Value = Val(LV.ListItems(i).ListSubItems(4).Tag) 'cmbProcess.ID
                    .Parameters("@Qty").Value = lQtyToUse 'Val(LV.ListItems(i).SubItems(6) & "")
                    .Parameters("@WasteQty").Value = 0 'Val(LV.ListItems(i).SubItems(5) & "")
                    .Parameters("@iRefID").Value = iRefID
                    .Parameters("@VendID").Value = cmbVend.ID
                    If Val(LV.ListItems(i).ListSubItems(7).Tag) = 0 Then
                        'lLotNo = GetSingleLongValue("MAX(CAST(LotNo AS INT))", "VendRcvdDetail") + 1
                        'Identify the Receiving Type, If it's against forgint save forging Batch & Mil Nos.
                        '(3).tag has VendIssdDetail.EntryID, But we Need VendIssued.EntryID as it is linked with RMIssuance
                        'Eventually we need MaterialLocationwiseStatus
                        lVI_EntryID = GetSingleLongValue("RefID", "VendIssdDetail", " WHERE EntryID=" & Val(LV.ListItems(i).ListSubItems(3).Tag))
                        lMLS_EntryID = GetSingleLongValue("MLS_RefID", "RMID_MLS_Details RMIDMLS INNER JOIN RawMaterialIssuanceDetail RMID ON RMID.EntryID=RMIDMLS.RMID_RefID INNER JOIN RawMaterialIssuance RMI ON RMI.IssNo=RMID.IssNo", " WHERE VI_RefID=" & lVI_EntryID)
                        strBatchNo = GetSingleStringValue("BatchNo", "MaterialLocationwiseStatus", " WHERE EntryID=" & lMLS_EntryID)
                        strMil_No = GetSingleStringValue("Mill_Certificate_No", "MaterialLocationwiseStatus", " WHERE EntryID=" & lMLS_EntryID)
                        If strMil_No = "" And txtMaterialCertNo <> "" Then
                            strMil_No = txtMaterialCertNo
                        End If
                        lLotNo = getNextLotNo
                        con.Execute "INSERT INTO Lots_List(LotNo,ItemID,Lot_Type,Reference_LotNo,Batch_No,Mill_Certificate_NO) VALUES('" & lLotNo & "','" & LV.ListItems(i).SubItems(1) & "',0,'','" & strBatchNo & "','" & strMil_No & "')"
                    Else
                        lLotNo = (LV.ListItems(i).ListSubItems(7).Tag)
                    End If
                    .Parameters("@LotNo").Value = lLotNo 'Val(LV.ListItems(i).ListSubItems(7).Tag)
                    .Parameters("@RcvDT").Value = Date 'LV.ListItems(i).SubItems(7)
                    .Parameters("@RecID").Value = "RCV-" & Format(ServerDate, "ddMMyy") & (GetSingleLongValue("MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-10) As Int))", "VVendRcvdItems", " WHERE CONVERT(VARCHAR(50),DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' and UPPER(LEFT(RecieptID,3))='RCV'") + 1) 'LV.ListItems(i).SubItems(10)
                    .Parameters("@OrderNo").Value = LV.ListItems(i).ListSubItems(1).Tag
                    .Parameters("@DeliveryRefID").Value = txtLotNo.Tag 'txtID.Tag
                    .Parameters("@CountedBy").Value = LV.ListItems(i).SubItems(9)
                    .Parameters("@Issue_RefID").Value = Val(LV.ListItems(i).ListSubItems(3).Tag)
                    .Parameters("@Weight").Value = Val(LV.ListItems(i).ListSubItems(7))
                    .Parameters("@ReqAuth").Value = iAuthRequired 'UserReqAuth
                    .Parameters("@UserID").Value = UserID
                    
                    .Parameters("@ReWorkLot").Value = LV.ListItems(i).ListSubItems(8).Tag
                    .Parameters("@RepairType").Value = Val(LV.ListItems(i).ListSubItems(9).Tag)     'RepairType
                    .Execute
                    
                    lEntryIDDetail = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetail")
                    If (strCompany = "Dr-Frgz") And Val(LV.ListItems(i).ListSubItems(7).Tag) = 0 Then
                        'If Lot No. Already Exists, Update Only
                        Dim iRecAffect_MD As Integer
                        con.Execute "UPDATE VendRcvdDetail_MoreDetails SET MaterialCertNo='" & txtMaterialCertNo & "',ManualPTCNo=NULL WHERE LotNo='" & lLotNo & "'", iRecAffect_MD
                        If iRecAffect_MD = 0 Then
                            con.Execute "INSERT INTO VendRcvdDetail_MoreDetails(LotNo,MaterialCertNo) VALUES('" & _
                             lLotNo & "','" & txtMaterialCertNo & "')"
                        End If
    
                        
                    End If
                     
                    'Show Bookmark If Exists....
                    Set frmViewBookMark = New frmAddBookmarkforReceiving
                    Load frmViewBookMark
                    
                    Call frmViewBookMark.ViewBookmark(Val(LV.ListItems(i).ListSubItems(3).Tag), True)
                    
                End With
        
                '            If UserReqAuth = 0 And cmbProcess.ID <> 4 Then
                '                Dim DetailRefID As Long
                '                DetailRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendRcvdDetail").Fields(0).Value & "")
                '                con.Execute "INSERT INTO VendRcvAuthDetail(UserID,RcvID,DT) Values(" & _
                '                UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
                '            End If
                con.Execute "INSERT INTO PrintSession(RecieptNo) VALUES('" & LV.ListItems(i).SubItems(10) & "')"
                lQtyUsed = lQtyUsed + lQtyToUse
            End If
        Wend
    Next i
    
    Set rs = Nothing
    
    If LV.ListItems(1).ListSubItems(7).Tag <> "" Then 'Receiving against Lot
        Call Process_Multiple_Issue_Receive(Val(LV.ListItems(1).ListSubItems(3).Tag))
    End If
    con.CommitTrans
    
    If Val(LV.ListItems(1).ListSubItems(7).Tag) = 0 Then    'Receiving against Master PO #...
        If MsgBox("Do you want to Close master PO?", vbQuestion + vbYesNo) = vbYes Then
            con.Execute "UPDATE VendIssued SET Closed=1 WHERE RecieptID='M-ISU-" & txtID & "'"
        End If
    End If
    
    If iAuthRequired Then
'        Dim lEntryID_ForInsp As Long
'        lEntryID_ForInsp = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetail")
        Dim f As New frmAuthorizeIssRcvAdv
        Load f
        'f.ShowMe (lEntryID_ForInsp)
        f.ShowMe (iRefID)
    End If
    
    Saved = True
    
    Exit Function
err:
    Call RollBackTransactions
    MsgBox err.Description
        
End Function
    

Private Sub LVEmp_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVEmp.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            LVEmp.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub mnuShowPicture_Click()

    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromDBField "ItemPic", "Items", " WHERE ItemID='" & LV.SelectedItem.SubItems(1) & "'"
        
End Sub

Private Sub txtCountedBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCountedBy) = "" Then
            MsgBox "Invlaid Name.", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.SubItems(iColNo - 1) = txtCountedBy
        txtCountedBy.Visible = False
    ElseIf KeyAscii = 27 Then
        txtCountedBy.Visible = False
    End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call GetEntry
    End If
    
    
End Sub

Private Sub GetEntry(Optional p_bByLotNo As Boolean = False)
    
    Dim strIssID As String
    Dim lEntryID As Long, lVendID As Long, lProcessID As Long, strItemID As String
    Dim iMaxRcvings As Integer, iRcvd As Integer
    'strIssID = txtID1.Text & txtID.Text
    Dim bNotAuthorized As Boolean
    Dim rs As New ADODB.Recordset
    bByLotNo = p_bByLotNo
    Dim lDetailEntryID As Long
    
'    lblMatCertNo.Visible = False
'    txtMaterialCertNo.Visible = False
    
    Dim strIssEmpID As String
    With rs
    
        If bByLotNo Then
            .Open "SELECT EntryID,VendID,ProcessID,ItemID,Authorized,MaximumRcvingsAgainstPO,IssEmpID FROM VendIssued WHERE EntryID IN(SELECT MAX(RefID) FROM VendIssdDetail WHERE LotNo='" & (txtLotNo) & "')", con, adOpenForwardOnly, adLockReadOnly
        Else
            Dim bClosed As Boolean
            bClosed = GetSingleBooleanValue("Closed", "VendIssued", " WHERE RecieptID='" & strIssID & "'")
            If bClosed Then
                MsgBox "Master PO Closed.", vbInformation
                Exit Sub
            End If
            .Open "SELECT EntryID,VendID,ProcessID,ItemID,Authorized,MaximumRcvingsAgainstPO,IssEmpID FROM VendIssued WHERE RecieptID='" & strIssID & "'", con, adOpenForwardOnly, adLockReadOnly
            If (strCompany = "Dr-Frgz") Then
                lblMatCertNo.Visible = True
                txtMaterialCertNo.Visible = True
            End If
        End If
        
        If .EOF = False Then
            If !Authorized Then
                lEntryID = Val(!EntryID & "")
                lVendID = Val(!VendID & "")
                lProcessID = Val(!ProcessID & "")
                strItemID = !ITemID & ""
                iMaxRcvings = Val(!MaximumRcvingsAgainstPO & "")
            Else
                MsgBox "Can't Receive, Order Not Authorized !!!", vbCritical
                .Close
                Exit Sub
            End If
            strIssEmpID = !issempid & ""
        End If
        .Close
        
        If bByLotNo Then
            .Open "SELECT EntryID FROM VendIssdDetail WHERE EntryID IN(SELECT MAX(EntryID) FROM VendIssdDetail WHERE LotNo='" & (txtLotNo) & "')", con, adOpenForwardOnly, adLockReadOnly
            If .EOF Then
                lDetailEntryID = 0
            Else
                lDetailEntryID = Val(!EntryID & "")
            End If
            .Close
            If GetSingleLongValue("COUNT(Issue_RefID)", "VendRcvdDetail", "WHERE Issue_RefID=" & lDetailEntryID & " AND LotNo='" & (txtLotNo) & "'") > 0 Then
                MsgBox "Already Received.", vbInformation
                Exit Sub
            End If
        End If
        
        .Open "SELECT COUNT(Issuance_RefID) FROM VendReceived WHERE Issuance_RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        iRcvd = Val(.Fields(0).Value & "")
        .Close
       
        cmbProcess.ID = lProcessID
        cmbVend.ID = lVendID
        'txtID.Tag = lEntryID
        txtLotNo.Tag = lEntryID
    End With
    Set rs = Nothing
    
    If Not (strCompany = "Weldon Instruments" Or strCompany = "QEL") Then
        If iRcvd >= iMaxRcvings Then
            MsgBox "Order Has been Received " & iRcvd & " times, While Allowed Rcvings are " & iMaxRcvings, vbInformation
            Exit Sub
        End If
    End If
    
    cmbVend.ClearVals
    Call cmbVend.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID=" & lVendID)
    cmbVend.ID = lVendID
    
    If strIssEmpID <> "" Then
        cmbEmp.ID = strIssEmpID
    End If
    
End Sub

Private Sub txtLotNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call GetEntry(True)
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    Dim dRate As Double, i As Integer
    If KeyAscii = 13 Then
        If iColNo = 7 Then
            'If Val(txtQty) > Val(LV.SelectedItem.ListSubItems(2).Tag) Then
            If Val(txtQty) > Val(LV.Tag) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
            If cmbVend.ID <> iFactoryMakerID Then
                dRate = GetSingleDoubleValue("Rate", "VendAssItems", " WHERE VendID=" & cmbVend.ID & " AND ItemID='" & LV.SelectedItem.SubItems(1) & "' AND ProcessID=" & cmbProcess.ID) 'Val(con.Execute("SELECT Rate FROM VendAssItems WHERE VendID=" & cmbVend.ID & " AND ItemID='" & LV.SelectedItem.SubItems(1) & "' AND ProcessID=" & cmbProcess.ID).Fields(0).Value & "")
                If dRate = 0 Then
                    MsgBox "Rate Is Zero in Vender Profile.", vbInformation
                    Exit Sub
                End If
            End If
        Else    'Weight Entry
            Dim lTotalRcvdQty As Long
            For i = 1 To LV.ListItems.count
                lTotalRcvdQty = lTotalRcvdQty + Val(LV.ListItems(i).SubItems(6))
            Next
            
            Dim dForgingWeight As Double, d10Per As Double
            dForgingWeight = Val(LV.SelectedItem.ListSubItems(5).Tag)
            dForgingWeight = dForgingWeight * lTotalRcvdQty / 1000
            'dForgingWeight = dForgingWeight * Val(LV.SelectedItem.SubItems(6)) / 1000 'Divide by 1000 to convert to KGS
            '10% ...
            
            d10Per = (dForgingWeight / 100) * 10
            
            If Not (Val(txtQty) > dForgingWeight - d10Per And Val(txtQty) < dForgingWeight + d10Per) Then
                MsgBox "Weight is not withing 10 Percent Range Limit." & vbNewLine & "It should be between " & dForgingWeight - d10Per & " AND " & dForgingWeight + d10Per & " Kgs." & vbNewLine & "Weight Entry should be in Kgs.", vbInformation
                Exit Sub
            End If
            
        End If

        If Val(txtQty) = 0 Then
            LV.SelectedItem.SubItems(6) = Val(txtQty)
               
            LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = ""
            'Reset The IssID's
            Dim TempID As Long
            TempID = 0
                
            For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1), 3)) - 1
                If TempID > 0 Then
                    LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                    Format(ServerDate, "ddMMyy") & TempID
                End If
            Next i
            If GotRcvID > 0 Then GotRcvID = GotRcvID - 1
                
            Pic.Visible = False
            Exit Sub
        
        End If
          
        If txtQty.Tag = "Simple" Then
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtQty)
            Pic.Visible = False
            Exit Sub
        End If
        Dim dQty As Double, dQtytoUse As Double, iStartIndex As Integer
        dQty = 0
        iStartIndex = LV.SelectedItem.Index
        If iColNo = 7 Or iColNo = 8 Then    'RcvdQty or RcvdWeigth
        
            While dQty < Val(txtQty)
                If iColNo = 7 Then
                    dQtytoUse = Val(LV.SelectedItem.ListSubItems(10).Tag)  'Val(LV.SelectedItem.ListSubItems(2).Tag)
                Else
                    dQtytoUse = (Val(txtQty) / lTotalRcvdQty) * Val(LV.SelectedItem.ListSubItems(10).Tag)
                End If
                If LV.SelectedItem.Index = LV.ListItems.count Then
                    dQtytoUse = Val(txtQty) - dQty
                Else
                    If dQtytoUse > Val(txtQty) - dQty Then dQtytoUse = Val(txtQty) - dQty
                End If
                
                dQtytoUse = Round(dQtytoUse, 2)
                LV.SelectedItem.SubItems(iColNo - 1) = dQtytoUse
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).Value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(10) = "RCV-" & Format(DTPicker1, "ddMMyy") & GotRcvID
                If LV.SelectedItem.Index < LV.ListItems.count Then
                    LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                    LV.SetFocus
                End If
                dQty = dQty + dQtytoUse
            Wend
            
        Else
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtQty)
        End If
        Pic.Visible = False
        DTPicker1.Enabled = False
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, iColNo <> 7)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Process_Multiple_Issue_Receive(p_lVID_EntryID As Long)
    
    Dim myDT As Date, lVendID As Long
    
    Dim lCount_MAI As Long
    Dim lMakerIssNo As Long, strMakerIssNo As String, strItemID As String, dRate As Double, lQty As Long, strOrderNo As String, bReWorkLot As Boolean, lRepair_RefID As Long
    Dim strLotNo As String
    Dim lProcessID As Long
    
    Dim lVI_EntryID As Long, lVID_EntryID As Long
    Dim lVRD_EntryID_Last As Long
    
    Dim lVR_EntryID As Long, lSNo_ProcessID As Long, lProcessID_Next As Long
    Dim lVRD_IssdQty As Long
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT VendIssued.VendID,VendIssued.DT,VID_Multi.ProcessID,VID_Multi.Rate,VendIssdDetail.ItemCode,VendIssdDetail.IssQty,VendIssdDetail.OrderNo,VendIssdDetail.LotNo,VendIssdDetail.ReWorkLot,VendIssdDetail.Repair_RefID FROM VID_Multiple_Process_Issuance VID_Multi INNER JOIN VendIssdDetail ON VID_Multi.VID_RefID=VendIssdDetail.EntryID INNER JOIN VendIssued ON VendIssued.EntryID=VendIssdDetail.RefID WHERE VID_RefID=" & p_lVID_EntryID & " ORDER BY VID_Multi.EntryID", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
                myDT = !DT
                lVendID = Val(!VendID & "")
                lProcessID = Val(!ProcessID & "")
                strItemID = !ItemCode & ""
                strOrderNo = !OrderNo & ""
                dRate = Val(!Rate & "")
                lQty = Val(!IssQty & "")
                strLotNo = !LotNo & ""
                bReWorkLot = !ReWorklot
                lRepair_RefID = Val(!Repair_RefID)
                If lMakerIssNo = 0 Then
                    lMakerIssNo = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-12) As Int)) FROM VendIssued WHERE Convert(varchar,DT,6)='" & Format(myDT, "dd MMM yy") & "'").Fields(0).Value & "")
                End If
                lMakerIssNo = lMakerIssNo + 1
                IssID = "ISU-" & Format(myDT, "ddMMyy") & GotIssID
                strMakerIssNo = "M-ISU-" & Format(myDT, "ddMMyy") & lMakerIssNo
                
                con.Execute "INSERT INTO VendIssued(VendID,DT,RecieptID,UserID,ProcessID,ItemID,UserName,MachineName,SpecialInstructions" & _
                 ",ExcessQtyPercentage,MaximumRcvingsAgainstPO,MasterPONo,Authorized,IssEmpID,SteelType_RefID) VALUES(" & lVendID & ",'" & _
                 myDT & "','" & strMakerIssNo & "'," & UserID & "," & lProcessID & ",'" & strItemID & "','" & CurrentUserName & "','" & _
                 strComputerName & "','Auto',0,0,'',1,'',0)"

                'lRefID = Val(con.Execute("SELECT Max(EntryID) FROM VendIssued").Fields(0).Value & "")
                lVI_EntryID = GetSingleLongValue("MAX(EntryID)", "VendIssued", " WHERE MachineName='" & strComputerName & "'")
                lVRD_EntryID_Last = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetail", " WHERE LotNo='" & strLotNo & "'")
                con.Execute "INSERT INTO VendIssdDetail(RefID,RecieptID,ItemCode,Rate,IssQty,RcvdQty,ReqAuth,OrderNo,RcvProcessID,ReturnDT,Priority" & _
                    ",Rcvd_RefID,LotNo,ReWorkLot,Repair_RefID) VALUES(" & lVI_EntryID & ",'" & strMakerIssNo & "','" & strItemID & "'," & _
                     dRate & "," & lQty & "," & lQty & ",0,'" & strOrderNo & "'," & lProcessID & ",'" & myDT & "',0," & _
                     lVRD_EntryID_Last & ",'" & strLotNo & "'," & Abs(bReWorkLot) & "," & lRepair_RefID & ")"
                     
                lVID_EntryID = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail")
                
                lCount_MAI = GetSingleLongValue("COUNT(*)", "VendAssItems", " WHERE VendID=" & lVendID & " AND ProcessID=" & lProcessID & " AND ItemID='" & strItemID & "'")
                If lCount_MAI = 0 Then
                    con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID,Rate,Unit,Remarks) VALUES(" & lVendID & "," & lProcessID & ",'" & strItemID & "',1,'Pcs','Auto-Multi')"
                End If
                
                
                ''''Receiving
                con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID,OverTime,UserName,MachineName,TemperValue) VALUES(" & lVendID & _
                 ",'" & myDT & "',''," & UserID & "," & lProcessID & "," & lVI_EntryID & ",0,'" & CurrentUserName & "','" & strComputerName & "','')"
    
            
                lVR_EntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived", " WHERE MachineName='" & strComputerName & "'")
                
                lSNo_ProcessID = GetSingleLongValue("SNo", "ItemProcesses", " WHERE ProcessID=" & lProcessID & " AND ItemID='" & strItemID & "'")
                lProcessID_Next = GetSingleLongValue("TOP 1 ProcessID", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND SNo>" & lSNo_ProcessID & " ORDER BY SNo")
                If .RecordCount = .AbsolutePosition Then
                    lVRD_IssdQty = 0
                Else
                    lVRD_IssdQty = lQty
                End If
                con.Execute "INSERT INTO VendRcvdDetail(RefID,RecieptID,ItemCode,NextProcessID,RcvdQty,IssQty,Wastage,Rate,LotNo,ReqAuth,OrderNo" & _
                    ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,ReWorkLot,Repair_RefID) VALUES(" & lVR_EntryID & ",'','" & _
                    strItemID & "'," & lProcessID_Next & "," & lQty & "," & lVRD_IssdQty & ",0," & dRate & ",'" & strLotNo & "',0,'" & strOrderNo & "',''," & _
                    lVID_EntryID & "," & lProcessID & ",0," & Abs(bReWorkLot) & "," & lRepair_RefID & ")"
                
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


