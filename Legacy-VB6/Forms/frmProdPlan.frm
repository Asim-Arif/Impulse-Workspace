VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProdPlan 
   ClientHeight    =   11130
   ClientLeft      =   60
   ClientTop       =   60
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
   LinkTopic       =   "Form37"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11130
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   10095
      Left            =   60
      TabIndex        =   3
      Top             =   570
      Width           =   11475
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   2325
         TabIndex        =   8
         Top             =   435
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   52822019
         CurrentDate     =   37631
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   4860
         Left            =   105
         TabIndex        =   7
         Top             =   735
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   8573
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S.#"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Order Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Total Plannable"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "InHouse"
            Object.Width           =   2540
         EndProperty
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3930
         Left            =   90
         TabIndex        =   10
         Top             =   5640
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   6932
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "In House Availability"
         TabPicture(0)   =   "frmProdPlan.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtRate"
         Tab(0).Control(1)=   "LVInHouse"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Transferred"
         TabPicture(1)   =   "frmProdPlan.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "LVTrans"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "txtUpdate1"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Production Planning"
         TabPicture(2)   =   "frmProdPlan.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "LVPlan"
         Tab(2).Control(1)=   "Pic1"
         Tab(2).Control(2)=   "txtRate1"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Planned"
         TabPicture(3)   =   "frmProdPlan.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "LVIssued"
         Tab(3).Control(1)=   "txtUpdate"
         Tab(3).ControlCount=   2
         Begin VB.TextBox txtUpdate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3330
            TabIndex        =   26
            Top             =   1650
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox txtRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -70050
            TabIndex        =   18
            Top             =   645
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox txtRate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72855
            TabIndex        =   17
            Top             =   1590
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.PictureBox Pic1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   -73065
            ScaleHeight     =   255
            ScaleWidth      =   6930
            TabIndex        =   12
            Top             =   870
            Visible         =   0   'False
            Width           =   6930
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   285
               Left            =   5580
               TabIndex        =   13
               Top             =   0
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   503
               _Version        =   393216
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   52822019
               CurrentDate     =   37679
            End
            Begin MSForms.ComboBox cmbVender 
               Height          =   255
               Left            =   2580
               TabIndex        =   16
               Top             =   0
               Width           =   3000
               VariousPropertyBits=   746604571
               BackColor       =   15199215
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5292;450"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               BorderColor     =   14277081
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtQty 
               Height          =   255
               Left            =   1590
               TabIndex        =   15
               Top             =   0
               Width           =   1005
               VariousPropertyBits=   679495707
               BackColor       =   15199215
               BorderStyle     =   1
               Size            =   "1764;450"
               BorderColor     =   14277081
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.ComboBox cmbPlan 
               Height          =   255
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   1600
               VariousPropertyBits=   746604571
               BackColor       =   15199215
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2822;450"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               BorderColor     =   14277081
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.TextBox txtUpdate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72060
            TabIndex        =   11
            Top             =   1875
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSComctlLib.ListView LVInHouse 
            Height          =   3465
            Left            =   -74910
            TabIndex        =   19
            Top             =   390
            Width           =   11100
            _ExtentX        =   19579
            _ExtentY        =   6112
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Production Phase"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Available Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Qty To Transfer"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView LVIssued 
            Height          =   3465
            Left            =   -74925
            TabIndex        =   20
            Top             =   375
            Width           =   11100
            _ExtentX        =   19579
            _ExtentY        =   6112
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
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Production Phase"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Vender/Supevisor"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item Code"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Description"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Issued Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Reciept ID"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Date"
               Object.Width           =   1940
            EndProperty
         End
         Begin MSComctlLib.ListView LVPlan 
            Height          =   3465
            Left            =   -74940
            TabIndex        =   21
            Top             =   375
            Width           =   11100
            _ExtentX        =   19579
            _ExtentY        =   6112
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Production Phase"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Type"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Vender/Supervisor"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Date"
               Object.Width           =   2355
            EndProperty
         End
         Begin MSComctlLib.ListView LVTrans 
            Height          =   3465
            Left            =   30
            TabIndex        =   22
            Top             =   375
            Width           =   11100
            _ExtentX        =   19579
            _ExtentY        =   6112
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Production Phase"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Code"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Transffered Qty"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   6240
         TabIndex        =   25
         Top             =   9660
         Width           =   1665
         Caption         =   "Save & New   "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   110
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   9675
         TabIndex        =   24
         Top             =   9660
         Width           =   1665
         Caption         =   "Cancel            "
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
      Begin MSForms.CommandButton cmdSaveClose 
         Height          =   360
         Left            =   7950
         TabIndex        =   23
         Top             =   9660
         Width           =   1665
         Caption         =   "Save & Close  "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbOrders 
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   435
         Width           =   2220
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   6244673
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3916;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox TBox 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   6
         Top             =   405
         Visible         =   0   'False
         Width           =   1635
         VariousPropertyBits=   679495711
         MaxLength       =   2
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "2884;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbItems 
         Height          =   285
         Left            =   5955
         TabIndex        =   5
         Top             =   405
         Visible         =   0   'False
         Width           =   1860
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   6244673
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   240
         Left            =   90
         TabIndex        =   4
         Top             =   195
         Width           =   2220
         BackColor       =   11517387
         Caption         =   "  Order No                   Items"
         Size            =   "3916;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Planning"
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
         TabIndex        =   1
         Top             =   120
         Width           =   11700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Planning"
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
         Index           =   1
         Left            =   75
         TabIndex        =   2
         Top             =   135
         Width           =   11700
      End
   End
End
Attribute VB_Name = "frmProdPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim SelItem As ListItem
Dim MyCol As New Collection
Dim OrderDT As Date
Dim PONo As Long
Dim ReqNo As Long
Public Function CheckForPlanning(TempOrd As String) As Boolean

    On Error GoTo err
    
    cmbOrders.Text = TempOrd
    
    If cmbOrders.MatchFound Then
        Call AddPlanning(TempOrd)
        CheckForPlanning = True
    Else
        CheckForPlanning = False
    End If
    
    Exit Function
    
err:
    MsgBox err.Description
    
End Function
Public Sub AddPlanning(TempOrd As String)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    Dim itm As ListItem
    
    strOrderNo = TempOrd
    Tbox(1) = strOrderNo
    cmbOrders = Tbox(1)
    For i = 1 To MyCol.Count
        MyCol.Remove 1
    Next i
    
    With rs
        .Open "Select CompItemCode,Qty,FreeStock,ItemName,RemQty,TotalFreeStock From VFItemsQtyToPlan Where OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        LV1.ListItems.Clear
        LVIssued.ListItems.Clear
        LVInHouse.ListItems.Clear
        LVTrans.ListItems.Clear
        LVPlan.ListItems.Clear
        
        Do Until .EOF
            Set itm = LV1.ListItems.Add(, ![CompItemCode] & "'", LV1.ListItems.Count + 1)
            itm.ListSubItems.Add , , ![CompItemCode] & ""
            itm.ListSubItems.Add , , ![ItemName] & ""
            
            itm.ListSubItems.Add , , Val(![Qty] & "") + Val(![FreeStock] & "")
            itm.ListSubItems.Add , , ![RemQty] & ""
            
            itm.ListSubItems.Add , , IIf(Val(![TotalFreeStock] & "") > 0, "InHouse", "No")
            'ITM.ListSubItems.Add , , Val(![Qty] & "") + Val(![FreeStock] & "")
            'ITM.ListSubItems.Add , , Val(![Qty] & "") + Val(![FreeStock] & "")
            'ITM.ListSubItems.Add
            '-----------------------------------------------------------------
            
            Dim Obj As New IssueItems
            
            Obj.LVIndex = LV1.ListItems.Count + 1
            
            Obj.IssueableQty = Val(![Qty] & "") + Val(![FreeStock] & "")
            MyCol.Add Obj, ![CompItemCode]
            
            Set Obj = Nothing
            
            .MoveNext
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
    OrderDT = con.Execute("Select DT From FCustomerOrders Where OrderNo='" & strOrderNo & "'").Fields(0).value & ""
    
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub EditOrder(TempOrd As String)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    strOrderNo = TempOrd
    
    With rs
        .Open "Select * From VFSteelPurchOrd Where OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        cmbOrders.Locked = True
        cmbOrders = ![OrderNo] & ""
        cmbVendor.ID = ![VendID]
        cmbLab.ID = ![LabID]
        Tbox(0) = ![ExtraPer] & ""
        Tbox(1) = ![PurchOrdNo] & ""
        Tbox(1).Tag = ![DT] & ""
        .Close
    End With
    Set rs = Nothing
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()
    On Error GoTo err
    Exit Sub
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VFOrderItemsProcs Where OrderNo+ItemID='" & Tbox(1) & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        Do Until .EOF
            Set itm = LV1.ListItems.Add(, ![ProcessID] & "'", LV1.ListItems.Count + 1)
            itm.SubItems(1) = ![Description] & ""
            itm.SubItems(2) = ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub cmbProcess_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode = 13 Then
        If cmbProcess.MatchFound = False Then
            MsgBox "Please Select Process From The List.", vbInformation
            Exit Sub
        End If
        
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select VendID,VenderName,EntryID,NickName from VVenders Where PhaseID=" & cmbProcess.List(cmbProcess.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
            cmbVend.Clear
            Do Until .EOF
                cmbVend.AddItem ![VenderName] & "(" & ![NickName] & ")"
                cmbVend.List(cmbVend.ListCount - 1, 1) = ![VendID]
                cmbVend.List(cmbVend.ListCount - 1, 2) = ![EntryID]
                .MoveNext
            Loop
            .Close
        End With
        Set rs = Nothing
        cmbPlan.SetFocus
    ElseIf KeyCode = 27 Then
        Pic1.Visible = False
    End If
    
End Sub


Private Sub cmbVend_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode = 13 Then
        If cmbProcess.MatchFound = False Then
            MsgBox "Please Select Process From The List."
            cmbProcess.SetFocus
            Exit Sub
        End If
        If cmbPlan.MatchFound = False Then
            MsgBox "Please Select Plan From The List."
            cmbPlan.SetFocus
            Exit Sub
        End If
        If cmbVend.MatchFound = False Then
            MsgBox "Please Select Vendor From The List."
            cmbVend.SetFocus
            Exit Sub
        End If
        
        SelItem.ListSubItems(2).Text = cmbProcess.Text
        SelItem.ListSubItems(2).Tag = cmbProcess.List(cmbProcess.ListIndex, 1)
        
        SelItem.SubItems(3) = cmbPlan.Text
        
        SelItem.ListSubItems(4).Text = cmbVend.Text
        SelItem.ListSubItems(4).Tag = cmbVend.List(cmbVend.ListIndex, 2)
        'SelItem.ListSubItems(4).Text cmbVend.Text
        
        Pic1.Visible = False
    ElseIf KeyCode = 27 Then
        Pic1.Visible = False
    End If
    
End Sub







Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Load frmPrevRpt
    Dim rptPurchOrdForSteel As craxddrt.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptPurchOrdForSteel.rpt") '
    rptPurchOrdForSteel.FormulaFields(5).Text = Val(Tbox(0))
    'rptPurchOrdForSteel.FormulaFields(2).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "{VFOrderSteelWeight.OrderNo}='" & cmbOrders.Text & "'", rptPurchOrdForSteel
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbOrders_Change()
    Call cmbOrders_Click
End Sub

Private Sub cmbOrders_Click()
    If cmbOrders.MatchFound Then
        Call AddPlanning(cmbOrders.Text)
    End If
End Sub

Private Sub cmbPlan_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    On Error GoTo err
    
    Select Case KeyCode.value
        Case 13
            If cmbPlan.ListIndex = 1 Then
                cmbVender.Clear
                cmbVender.AddItem con.Execute("Select Supervisor From Processes Where ProcessID=" & Val(LVPlan.SelectedItem.key)).Fields(0).value & ""
                cmbVender.ListIndex = 0
            ElseIf cmbPlan.ListIndex = 0 Then
                Dim rs As New ADODB.Recordset
                With rs
                    If Tbox(1) = "HICO Order" Then
                        .Open "Select VendID,Rate,VenderName  From VVendAssItems Where ItemID='" & LV1.SelectedItem.SubItems(1) & "' and PhaseID=" & Val(LVPlan.SelectedItem.key) & " and VendID Not In" & cmbVender.Tag & " Or PhaseID=-1", con, adOpenForwardOnly, adLockReadOnly
                    Else
                        .Open "Select VendID,Rate,VenderName  From VVendAssItems Where ItemID='" & LV1.SelectedItem.SubItems(1) & "' and PhaseID=" & Val(LVPlan.SelectedItem.key) & " and VendID Not In" & cmbVender.Tag, con, adOpenForwardOnly, adLockReadOnly
                    End If
                    cmbVender.Clear
                    Do Until .EOF
                        cmbVender.AddItem .Fields(2) & ""
                        cmbVender.List(.AbsolutePosition - 1, 1) = ![VendID] & ""
                        cmbVender.List(.AbsolutePosition - 1, 2) = ![Rate] & ""
                        .MoveNext
                    Loop
                    .Close
                End With
                Set rs = Nothing
            End If
            With txtQty
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
        Case 27
            Pic1.Visible = False
    End Select
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbVender_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    On Error GoTo err
    
    If KeyCode = 13 Then
        DTPicker2.SetFocus
    ElseIf KeyCode = 27 Then
        Pic1.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSaveClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub
Private Function CheckForRepeat() As Boolean

    
    
End Function
Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)

    If cmbVender.MatchFound = False Then Exit Sub
    
    'Check To see Whether Selected Vender Is Inactivate...
    If cmbPlan.ListIndex = 0 Then
        
    End If
    
    Select Case KeyCode
        Case 13
            
            For i = 1 To LVIssued.ListItems.Count
                If LV1.SelectedItem.key & "|" & cmbVender.List(cmbVender.ListIndex, 1) = LVIssued.ListItems(i).key Then
                    MsgBox "Selected Item Is Already Issued To " & cmbVender.Text & vbNewLine & _
                    "You Can Edit Or Delete It From Issued Tab.", vbInformation
                    Exit Sub
                End If
            Next i
            
            Dim IssID As String
            For i = 1 To LVIssued.ListItems.Count
                If Val(cmbVender.List(cmbVender.ListIndex, 1) & "") = Val(LVIssued.ListItems(i).ListSubItems(2).Tag & "") Then
                    IssID = LVIssued.ListItems(i).ListSubItems(6)
                    GoTo GotID:
                End If
            Next i
            
            If cmbPlan.ListIndex = 1 Then   'In House Production, Requsition Will Rise.
                If ReqNo = 0 Then
                    ReqNo = Val(con.Execute("Select Max(right(ReqNo,3)) From Requsitions Where Convert(varchar,PlanDT,6)='" & Format(ServerDate, "dd MMM yy") & "'").Fields(0).value & "")
                    ReqNo = ReqNo + 1
                Else
                    ReqNo = ReqNo + 1
                End If
                
                IssID = "HSRQ-" & _
                 Format(ServerDate, "ddMMyy") & "-" & Format(ReqNo, "000")
                 
            Else                            'Direct Purchase, Purchase Order.
                If PONo = 0 Then
         
                    PONo = Val(con.Execute("Select Max(right(PONo,3)) From VendPOrders Where Convert(varchar,PlanDT,6)='" & Format(ServerDate, "dd MMM yy") & "'").Fields(0).value & "")
                    PONo = PONo + 1
                Else
                    PONo = PONo + 1
                End If
                
                IssID = "HSPO-" & _
                 Format(ServerDate, "ddMMyy") & "-" & Format(PONo, "000")
                
            End If
GotID:
        
            Dim itm As ListItem, SelItm As ListItem, SelITM1 As ListItem
            Set SelItm = LVPlan.SelectedItem
            Set SelITM1 = LV1.SelectedItem
            Set itm = LVIssued.ListItems.Add(, LV1.SelectedItem.key & "|" & cmbVender.List(cmbVender.ListIndex, 1), LVIssued.ListItems.Count + 1)
            itm.Tag = SelITM1.SubItems(1)
            itm.ListSubItems.Add , SelItm.key, SelItm.SubItems(1)
            'ITM.ListSubItems.Add , cmbVender.List(cmbVender.ListIndex, 1) & "'", cmbVender.Text
            itm.ListSubItems.Add , , cmbVender.Text
            itm.ListSubItems(2).Tag = cmbVender.List(cmbVender.ListIndex, 1) 'VenderID
            
            itm.ListSubItems.Add , , SelITM1.SubItems(1) 'Item Code
            itm.ListSubItems.Add , , SelITM1.SubItems(2) 'Item Name
            itm.ListSubItems.Add(, , Val(txtQty)).Tag = cmbVender.List(cmbVender.ListIndex, 2) 'Rate
            itm.ListSubItems.Add , , IssID
            itm.ListSubItems.Add , , DTPicker2.value
            'Add Key Of MyCol Collection To The Tag Of This ListSubItem
            'Which Will Be Used For Updation Or deletion Of This Item's Qty.
            itm.ListSubItems(7).Tag = LV1.SelectedItem.Index
            Pic1.Visible = False
            
            MyCol(LV1.SelectedItem.Index).IssueableQty = MyCol(LV1.SelectedItem.Index).IssueableQty - Val(txtQty)
            LV1.SelectedItem.SubItems(4) = Val(LV1.SelectedItem.SubItems(4)) - Val(txtQty)
            'txtRate.Visible = False
            
            If LV1.SelectedItem.Index < LV1.ListItems.Count Then
            
                If Val(LV1.SelectedItem.SubItems(4)) = 0 Then
                    LV1.ListItems(LV1.SelectedItem.Index + 1).Selected = True
                End If
                
            End If
            Call LV1_ItemClick(LV1.SelectedItem)
            LV1.SetFocus
        Case 27
            Pic1.Visible = False
    End Select
End Sub

Private Sub Form_Load()
         
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        '.Open "Select Distinct OrderNo From VFItemsQtyToPlan Where RemQty>0", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select OrderNo From FCustomerOrders Where OrderNo Not In(Select OrderNo From VendIssdDetail)", con, adOpenForwardOnly, adLockReadOnly
        cmbOrders.Clear
        Do Until .EOF
            cmbOrders.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    cmbPlan.AddItem "Direct Purchase"
    cmbPlan.AddItem "InHouse Production"
    SSTab1.Tab = 0
    'cmbVend.AddItem "Abc"
    DTPicker1.value = ServerDate
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Left = (Me.ScaleWidth - fTitle.Width) / 2
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    'Me.Top = (MainForm.Height - Me.Height) / 10
    '28894
    
End Sub
Private Function Saved1() As Boolean

    On Error GoTo err
    
    
    Saved1 = False
    
    'For i = 1 To LV1.ListItems.Count
    '    If LV1.ListItems(i).SubItems(6) <> "0" Then
    '        MsgBox "Please Complete Entries For All The Items.", vbInformation
    '        Exit Function
    '    End If
    'Next i
    
    Me.MousePointer = vbHourglass
    Call StartTrans(con)
    
    For i = 1 To LVIssued.ListItems.Count
        With LVIssued.ListItems(i)
            con.Execute "Insert Into FOrdProdPlan(OrderNo,ItemCode,ProcessID,InHouse,VendID,Qty) Values('" & _
             Tbox(1) & "','" & .SubItems(3) & "'," & _
             Val(.ListSubItems(1).key) & ",1," & _
             Val(.ListSubItems(2).key) & "," & Val(.SubItems(5)) & ")"
        End With
    Next i
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    con.CommitTrans
    
    strOrderNo = ""
    Saved1 = True
    Me.MousePointer = vbDefault
    Exit Function
    
err:
    Me.MousePointer = vbDefault
    MsgBox err.Description
    
End Function



Private Sub LV1_DblClick()
    
    Exit Sub
    If LV1.ListItems.Count = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset
    
    Set SelItem = LV1.SelectedItem
    
    With rs
        .Open "Select * From VFOrderItemsProcs Where OrderNo='" & strOrderNo & "' and CompItemCode='" & SelItem.SubItems(1) & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbProcess.Clear
        Do Until .EOF
            cmbProcess.AddItem ![Description] & ""
            cmbProcess.List(cmbProcess.ListCount - 1, 1) = ![ProcessID] & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    'If cmbProcess.ListCount > 0 Then cmbProcess.ListIndex = 0
    
    
    Pic1.Move LV1.ColumnHeaders(3).Left + LV1.Left, LV1.SelectedItem.Top + LV1.Top
    
    If LV1.SelectedItem.SubItems(2) = "" Then
        cmbProcess.ListIndex = 0
    Else
        cmbProcess.Text = LV1.SelectedItem.SubItems(2) & ""
    End If
    
    'Combo1.Move LV1.ColumnHeaders(3).Left + LV1.Left, LV1.SelectedItem.Top + LV1.Top, LV1.ColumnHeaders(3).Width
    Pic1.Visible = True
    cmbProcess.SetFocus
    
End Sub

Private Sub LV1_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo err
    
    If Left(Tbox(1), 10) <> "HICO Order" Then
        If LV1.SelectedItem.SubItems(5) = "InHouse" Then
            Call FillFreeStock
        Else
            LVInHouse.ListItems.Clear
        End If
    End If
    
    Call FillProcesses
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub FillProcesses()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From VFOrderItemsProcs Where OrderNo='" & strOrderNo & "' and CompItemCode='" & LV1.SelectedItem.key, con, adOpenForwardOnly, adLockReadOnly
        LVPlan.ListItems.Clear
        Do Until .EOF
            Set itm = LVPlan.ListItems.Add(, ![ProcessID] & "'", LVPlan.ListItems.Count + 1)
            itm.ListSubItems.Add , , ![Description] & ""
            itm.ListSubItems.Add , , ""
            itm.ListSubItems.Add , , ""
            itm.ListSubItems.Add , , ""
            .MoveNext
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub FillFreeStock()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    Dim ITM1 As ListItem
    Dim lFreeStock As Long
    
    With rs
        .Open "Select * From VFFreeStockNew Where ItemCode='" & LV1.SelectedItem.key, con, adOpenForwardOnly, adLockReadOnly
        LVInHouse.ListItems.Clear
        
        Do Until .EOF
        
            Set itm = LVInHouse.ListItems.Add(, ![ProcessID] & "'", LVInHouse.ListItems.Count + 1)
            itm.Tag = Val(![InStock] & "")
            'If IsNull(![NextProcessID]) Then
            '    itm.ListSubItems.Add , , "Finished Pcs Available."
            'Else
            itm.ListSubItems.Add , , ![ProcDescription] & ""
            itm.ListSubItems(1).Tag = Val(![OutStock] & "")
            'End If
            
            lFreeStock = Val(![FreeStock] & "")
            
            For i = 1 To LVTrans.ListItems.Count
                Set ITM1 = LVTrans.ListItems(i)
                'If ITM1.SubItems(2) = LV1.SelectedItem.key And Val(ITM1.ListSubItems(1).key) = ![ProcessID] Then
                If ITM1.SubItems(2) = Left(LV1.SelectedItem.key, Len(LV1.SelectedItem.key) - 1) And Val(ITM1.ListSubItems(1).key) = ![ProcessID] Then
                    lFreeStock = lFreeStock - Val(LVTrans.ListItems(i).SubItems(4) & "")
                End If
            Next i
            
            itm.ListSubItems.Add , , lFreeStock
            itm.ListSubItems.Add , , "0"
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV1_KeyDown(KeyCode As Integer, Shift As Integer)

    'Ctrl+Enter =  Keycode=13 And Shift=2
    If KeyCode = 13 Then
        If Shift = 2 Then
            SSTab1.Tab = 0
            LVInHouse.SetFocus
        Else
            SSTab1.Tab = 2
            LVPlan.SetFocus
        End If
    End If
    
End Sub

Private Sub LVInHouse_DblClick()

    If LVInHouse.ListItems.Count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub

    
    With txtRate
        .Move LVInHouse.Left + LVInHouse.ColumnHeaders(4).Left, LVInHouse.Top + LVInHouse.SelectedItem.Top, LVInHouse.ColumnHeaders(4).Width
        '.Text = MyCol(LV.SelectedItem.SubItems(2) & "|" & LV.SelectedItem.SubItems(3)).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        '.Text = MyCol(LV1.SelectedItem.Key).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Text = LVInHouse.SelectedItem.SubItems(2)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End Sub


Private Sub LVInHouse_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LVInHouse_DblClick
    ElseIf KeyAscii = 8 Then
        LV1.SetFocus
    End If
End Sub

Private Sub LVIssued_DblClick()
    On Error GoTo err
    
    If LVIssued.ListItems.Count = 0 Then Exit Sub
    With txtUpdate
        .Text = Val(LVIssued.SelectedItem.SubItems(5))
        .Tag = Val(LVIssued.SelectedItem.SubItems(5))
        .Move LVIssued.Left + LVIssued.ColumnHeaders(6).Left + 10, LVIssued.SelectedItem.Top + LVIssued.Top, LVIssued.ColumnHeaders(6).Width
        
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LVIssued_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If txtUpdate.Visible Then txtUpdate.Visible = False
End Sub

Private Sub LVIssued_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    Dim IsPO As Boolean 'To Remember Whether The PO Entry is Being deleted Or Req.
    If LVIssued.ListItems.Count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        
        If MsgBox("Are You Sure to Delete This Entry!.", vbQuestion + vbYesNo) = vbYes Then
            Dim lKey As Long
            lKey = LVIssued.SelectedItem.ListSubItems(7).Tag
            MyCol(lKey).IssueableQty = MyCol(lKey).IssueableQty + LVIssued.SelectedItem.ListSubItems(5)
            
            'First Get The Item Index Of The LVIssued's Selected Item For LV1.
            Dim lIndex As Long
            For i = 1 To LV1.ListItems.Count
                If LV1.ListItems(i).SubItems(1) = LVIssued.SelectedItem.SubItems(3) Then
                    lIndex = i
                    Exit For
                End If
            Next i
            
            
            LV1.ListItems(lIndex).SubItems(4) = MyCol(lKey).IssueableQty 'Val(LV1.ListItems(lIndex).SubItems(4)) + LVIssued.SelectedItem.ListSubItems(5)
            
            If Left(LVIssued.SelectedItem.SubItems(6), 4) = "HSPO" Then
                IsPO = True
            Else
                IsPO = False
            End If
            
            LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
            
            'Reset The IssID's
            Dim TempID As Long
            TempID = 0
            If LVIssued.ListItems.Count > 0 Then
                For i = LVIssued.SelectedItem.Index To LVIssued.ListItems.Count
                
                    If Left(LVIssued.ListItems(i).SubItems(6), 4) = "HSRQ" And IsPO = False Then
                        TempID = Val(Right(LVIssued.ListItems(i).SubItems(6), 3)) - 1
                        If TempID > 0 Then
                            LVIssued.ListItems(i).SubItems(6) = "HSRQ-" & _
                             Format(ServerDate, "ddMMyy") & "-" & Format(TempID, "000")
                        End If
                    ElseIf Left(LVIssued.ListItems(i).SubItems(6), 4) = "HSPO" And IsPO Then
                        TempID = Val(Right(LVIssued.ListItems(i).SubItems(6), 3)) - 1
                        If TempID > 0 Then
                            LVIssued.ListItems(i).SubItems(6) = "HSPO-" & _
                             Format(ServerDate, "ddMMyy") & "-" & Format(TempID, "000")
                        End If
                    End If
                Next i
            Else
                PONo = 0
                ReqNo = 0
            End If
            
            If PONo > 0 And IsPO Then PONo = PONo - 1
            If ReqNo > 0 And IsPO = False Then ReqNo = ReqNo - 1
            
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LVPlan_DblClick()

    On Error GoTo err
    
    
    If Val(LV1.SelectedItem.SubItems(4)) = 0 Then
        'MsgBox "You Must Plan For All The Items!", vbInformation
        Exit Sub
    End If
    
    Dim rs As New ADODB.Recordset
    'Dim ITM As ListItem
    
    Dim VendIDList As String
    VendIDList = ""
    
    For i = 1 To LVIssued.ListItems.Count
        With LVIssued.ListItems(i)
            If LV1.SelectedItem.key = .Tag Then
                VendIDList = VendIDList & Val(.ListSubItems(1).key) & ","
            End If
        End With
    Next i
    
    If VendIDList <> "" Then
        VendIDList = "(" & Left(VendIDList, Len(VendIDList) - 1) & ")"
    Else
        VendIDList = "(0)"
    End If
    cmbVender.Tag = VendIDList
    
    
    cmbPlan.Text = ""
    txtQty = ""
    cmbVender = ""
    
    Pic1.Move LVPlan.Left + LVPlan.ColumnHeaders(3).Left + 10, LVPlan.SelectedItem.Top + LVPlan.Top
    txtQty = Val(LV1.SelectedItem.SubItems(4))
    Pic1.Visible = True
    cmbPlan.SetFocus
    
    DTPicker2.value = OrderDT
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LVPlan_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Pic1.Visible = False
End Sub

Private Sub LVPlan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LVPlan_DblClick
    ElseIf KeyAscii = 8 Then
        LV1.SetFocus
    End If
End Sub

Private Sub LVTrans_DblClick()

    On Error GoTo err
    
    If LVTrans.ListItems.Count = 0 Then Exit Sub
    With txtUpdate1
        .Text = Val(LVTrans.SelectedItem.SubItems(4))
        .Tag = Val(LVTrans.SelectedItem.SubItems(4))
        .Move LVTrans.Left + LVTrans.ColumnHeaders(5).Left + 10, LVTrans.SelectedItem.Top + LVTrans.Top, LVTrans.ColumnHeaders(5).Width
        
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LVTrans_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If LVTrans.ListItems.Count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        
        If MsgBox("Are You Sure to Delete This Entry!.", vbQuestion + vbYesNo) = vbYes Then
            Dim lKey As Long
            lKey = LVTrans.SelectedItem.ListSubItems(4).Tag
            MyCol(lKey).IssueableQty = MyCol(lKey).IssueableQty + LVTrans.SelectedItem.ListSubItems(4)
            
            'First Get The Item Index Of The LVIssued's Selected Item For LV1.
            Dim lIndex As Long
            For i = 1 To LV1.ListItems.Count
                If LV1.ListItems(i).SubItems(1) = LVTrans.SelectedItem.SubItems(2) Then
                    lIndex = i
                    Exit For
                End If
            Next i
            
            LV1.ListItems(lIndex).SubItems(4) = MyCol(lKey).IssueableQty 'Val(LV1.ListItems(lIndex).SubItems(4)) + LVIssued.SelectedItem.ListSubItems(5)
            
            LVTrans.ListItems.Remove LVTrans.SelectedItem.Index
            
            Call LV1_ItemClick(LV1.SelectedItem)
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtQty_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    Select Case KeyCode.value
        Case 13
        
            If Val(txtQty) > Val(LV1.SelectedItem.SubItems(4)) Then
                MsgBox "Invalid Qty.", vbInformation
                KeyCode.value = 0
                txtQty.SelStart = 0
                txtQty.SelLength = Len(txtQty)
                txtQty.SetFocus
                Exit Sub
            End If
            
            If cmbPlan.ListIndex = 0 Then
                cmbVender.SetFocus
            ElseIf cmbPlan.ListIndex = 1 Then
                DTPicker2.SetFocus
            End If
        Case 27
            Pic1.Visible = False
    End Select
    
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    Select Case KeyCode
    
        Case 13
            
            
            If Val(txtRate) > Val(LV1.SelectedItem.SubItems(4)) Then
                MsgBox "Invalid Qty.", vbInformation
                KeyCode = 0
                txtRate.SelStart = 0
                txtRate.SelLength = Len(txtQty)
                txtRate.SetFocus
                Exit Sub
            End If
            
            Dim itm As ListItem, SelItm As ListItem, SelITM1 As ListItem
            Set SelItm = LVInHouse.SelectedItem
            Set SelITM1 = LV1.SelectedItem
            Set itm = LVTrans.ListItems.Add(, LV1.SelectedItem.key & "|" & Val(LVInHouse.SelectedItem.key), LVTrans.ListItems.Count + 1)
            itm.Tag = SelITM1.SubItems(1)
            itm.ListSubItems.Add , SelItm.key, SelItm.SubItems(1)
            itm.ListSubItems.Add , , SelITM1.SubItems(1)
            itm.ListSubItems.Add , , SelITM1.SubItems(2)
            itm.ListSubItems.Add , , Val(txtRate)
            
            itm.ListSubItems(1).Tag = SelItm.Tag
            itm.ListSubItems(2).Tag = SelItm.ListSubItems(1).Tag
            
            itm.ListSubItems(4).Tag = LV1.SelectedItem.Index    'To Keep Track
            itm.ListSubItems(3).Tag = LVInHouse.SelectedItem.Index           'To Keep Track For InHouse
            
            txtRate.Visible = False
                    
            MyCol(LV1.SelectedItem.Index).IssueableQty = MyCol(LV1.SelectedItem.Index).IssueableQty - Val(txtRate)
            LV1.SelectedItem.SubItems(4) = Val(LV1.SelectedItem.SubItems(4)) - Val(txtRate)
            'LVInHouse.SelectedItem.SubItems(4) = Val(LVInHouse.SelectedItem.SubItems(4)) - Val(txtRate)
            'txtRate.Visible = False
            Call LV1_ItemClick(LV1.SelectedItem)
            
        Case 27
            txtRate.Visible = False
            
    End Select
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub
Private Function Saved() As Boolean

    On Error GoTo err
    
    For i = 1 To LV1.ListItems.Count
        If Val(LV1.ListItems(i).SubItems(4)) > 0 Then
            MsgBox "You Must Plan For All The Items!", vbInformation
            Exit Function
        End If
    Next i
    
    Dim SaveIt As Boolean
    
    Dim InHOuse As Boolean
    
    Dim VendID As New Collection
    Dim VIssID As New Collection
    Dim VOrdID As New Collection
    
    Dim TempIndex As Integer
    TempIndex = 1
    
    Dim lRefID As Long
    Call StartTrans(con)
    
    For i = 1 To LVIssued.ListItems.Count
        With LVIssued.ListItems(i)
            
            SaveIt = True
            For j = 1 To VendID.Count
                'If Val(LVIssued.ListItems(i).ListSubItems(1).Key) = VendID(j) Then
                If Val(LVIssued.ListItems(i).ListSubItems(2).Tag & "") = VendID(j) Then
                    'It Means That It's Been Saved Already.
                    SaveIt = False
                    Exit For
                End If
            Next j

            If SaveIt Then
            
                'VendID.Add Val(LVIssued.ListItems(i).ListSubItems(1).Key)
                VendID.Add Val(LVIssued.ListItems(i).ListSubItems(2).Tag & "")
                
                    con.Execute "Insert Into VendIssued(DT,VendID) Values('" & _
                     DTPicker1 & "'," & Val(LVIssued.ListItems(i).ListSubItems(2).Tag & "") & ")"
                 
                    lRefID = Val(con.Execute("Select Max(EntryID) From VendIssued").Fields(0).value & "")
        
                    VIssID.Add lRefID
                
            Else
                
                    'lRefID = VOrdID(j)
                    lRefID = VIssID(j)
            End If
            
            
                '1. Insert Into VendIssuedItems
                con.Execute "Insert Into VendIssdDetail(RefID,RecieptID,OrderNo,ItemCode" & _
                 ",Rate,ProcessID,IssQty,FromPlan) Values('" & lRefID & _
                 "','" & .SubItems(6) & "','" & Tbox(1) & "','" & .SubItems(3) & "'," & _
                 Val(.ListSubItems(5).Tag & "") & "," & Val(.ListSubItems(1).key) & "," & _
                 Val(.SubItems(5)) & ",1)"
                                 
                
                'Get The ID OF The New Entry Using The Same Variable...
                lRefID = Val(con.Execute("Select Max(EntryID) From VendIssdDetail").Fields(0).value & "")
                '2. Insert Into PO Or Req.
                
                If Left(.SubItems(6), 4) = "HSPO" Then
                    con.Execute "Insert Into VendPOrders(PONo,RefID,PlanDT,RcvingDT) Values('" & _
                     .SubItems(6) & "'," & lRefID & _
                     ",'" & ServerDate & "','" & .SubItems(7) & "')"
                Else
                    con.Execute "Insert Into Requsitions(ReqNo,RefID,PlanDT,RcvingDT) Values('" & _
                     .SubItems(6) & "'," & lRefID & ",'" & _
                     ServerDate & "','" & .SubItems(7) & "')"
                End If
        End With
        
    Next i
                       
   
    'Now Save Transferred Items...
    Dim cmd As New ADODB.Command
    Dim lTotInStock As Long
    Dim lTotOutStock As Long
    Dim lInStock As Long
    Dim lOutStock As Long
    Dim lQty As Long
    
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    For i = 1 To LVTrans.ListItems.Count
    
        
        
        lQty = Val(LVTrans.ListItems(i).SubItems(4) & "")
        lTotInStock = Val(LVTrans.ListItems(i).ListSubItems(1).Tag & "")
        lTotOutStock = Val(LVTrans.ListItems(i).ListSubItems(2).Tag & "")
        
        
        If lQty <= lTotInStock Then
            lInStock = lQty
            lOutStock = 0
        Else
            lInStock = lTotInStock
            lOutStock = lQty - lTotInStock
        End If
            
        With cmd
            If lInStock > 0 Then
        
                .CommandText = "SP_InsertTransferred"
                .Parameters("@OrderNo").value = Tbox(1)
                .Parameters("@ItemCode").value = LVTrans.ListItems(i).SubItems(2) & ""
                .Parameters("@ProcessID").value = Val(LVTrans.ListItems(i).ListSubItems(1).key)
                        
            
                
                .Parameters("@Qty").value = lInStock 'Val(LVTrans.ListItems(i).SubItems(4) & "")
                .Parameters("@DT").value = Format(ServerDate, "dd-MMM-yyyy")
                .Execute
                
            End If
            If lOutStock > 0 Then
                'Do For Out Stock Tooooooooo.
                .CommandText = "SP_InsertTransferredIssd"
                .Parameters("@OrderNo").value = Tbox(1)
                .Parameters("@ItemCode").value = LVTrans.ListItems(i).SubItems(2) & ""
                .Parameters("@ProcessID").value = Val(LVTrans.ListItems(i).ListSubItems(1).key)
                
                .Parameters("@Qty").value = lOutStock 'Val(LVTrans.ListItems(i).SubItems(4) & "")
                .Parameters("@DT").value = Format(ServerDate, "dd-MMM-yyyy")
                .Execute
            End If
        End With
    Next i
    
    con.CommitTrans
    
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function


Private Sub txtUpdate_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    
    
    If KeyAscii = 13 Then
        
        Dim lKey As Long
        lKey = LVIssued.SelectedItem.ListSubItems(7).Tag
        
        If Val(txtUpdate) = 0 Then
            MsgBox "Invalid Qty!", vbInformation
            Exit Sub
        ElseIf Val(txtUpdate) > Val(txtUpdate.Tag) + MyCol(lKey).IssueableQty Then
            MsgBox "Invalid Qty!", vbInformation
            Exit Sub
        End If
        
        MyCol(lKey).IssueableQty = MyCol(lKey).IssueableQty + Val(txtUpdate.Tag) - Val(txtUpdate)
        
        'First Get The Item Index Of The LVIssued's Selected Item For LV1.
        Dim lIndex As Long
        For i = 1 To LV1.ListItems.Count
            If LV1.ListItems(i).SubItems(1) = LVIssued.SelectedItem.SubItems(3) Then
                lIndex = i
                Exit For
            End If
        Next i
        
        LV1.ListItems(lIndex).SubItems(4) = MyCol(lKey).IssueableQty
        LVIssued.SelectedItem.SubItems(5) = Val(txtUpdate)
        
        
        txtUpdate.Visible = False
        
    ElseIf KeyAscii = 27 Then
        txtUpdate.Visible = False
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub txtUpdate1_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    
    
    If KeyAscii = 13 Then
        
        Dim lKey As Long
        lKey = LVTrans.SelectedItem.ListSubItems(4).Tag
        
        If Val(txtUpdate1) = 0 Then
            MsgBox "Invalid Qty!", vbInformation
            Exit Sub
        ElseIf Val(txtUpdate1) > Val(txtUpdate1.Tag) + MyCol(lKey).IssueableQty Then
            MsgBox "Invalid Qty!", vbInformation
            Exit Sub
        End If
        
        MyCol(lKey).IssueableQty = MyCol(lKey).IssueableQty + Val(txtUpdate1.Tag) - Val(txtUpdate1)
        
        'First Get The Item Index Of The lvtrans's Selected Item For LV1.
        Dim lIndex As Long
        For i = 1 To LV1.ListItems.Count
            If LV1.ListItems(i).SubItems(1) = LVTrans.SelectedItem.SubItems(2) Then
                lIndex = i
                Exit For
            End If
        Next i
        
        LV1.ListItems(lIndex).SubItems(4) = MyCol(lKey).IssueableQty
        LVTrans.SelectedItem.SubItems(4) = Val(txtUpdate1)
        
        'If LVTrans.SelectedItem.SubItems(2) = Left(LV1.SelectedItem.key, Len(LV1.SelectedItem.key) - 1) And Val(ITM1.ListSubItems(1).key) = ![ProcessID] Then
            'lFreeStock = lFreeStock - Val(LVTrans.ListItems(i).SubItems(4) & "")
        'End If
        
        'Update The Qnty In LvInHouse
        If LVTrans.SelectedItem.SubItems(2) = Left(LV1.SelectedItem.key, Len(LV1.SelectedItem.key) - 1) Then
            LVInHouse.ListItems(LVTrans.SelectedItem.ListSubItems(3).Tag).SubItems(2) = Val(LVInHouse.ListItems(LVTrans.SelectedItem.ListSubItems(3).Tag).SubItems(2)) + Val(txtUpdate1.Tag) - Val(txtUpdate1.Text)
        End If
        
        txtUpdate1.Visible = False
        Call LV1_ItemClick(LV1.SelectedItem)
    ElseIf KeyAscii = 27 Then
        txtUpdate1.Visible = False
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub
