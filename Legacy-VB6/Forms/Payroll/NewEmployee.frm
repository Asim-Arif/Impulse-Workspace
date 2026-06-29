VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form NewEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee's Info..."
   ClientHeight    =   7965
   ClientLeft      =   -18000
   ClientTop       =   2160
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10845
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   8265
      TabIndex        =   93
      Top             =   -1335
      Visible         =   0   'False
      Width           =   3915
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin VB.Frame fCat 
      Caption         =   "Department's Info :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   225
      TabIndex        =   94
      Top             =   150
      Width           =   8070
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   96
         Top             =   330
         Width           =   1650
      End
      Begin MSForms.ComboBox cmbDept 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   285
         Width           =   3090
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5450;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbDeptID 
         Height          =   300
         Left            =   6375
         TabIndex        =   69
         Top             =   285
         Width           =   1575
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2778;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Department ID :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   4965
         TabIndex        =   95
         Top             =   345
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab EmpTab 
      Height          =   6240
      Left            =   195
      TabIndex        =   72
      Top             =   990
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   11007
      _Version        =   393216
      Style           =   1
      Tabs            =   18
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "NewEmployee.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCNICPDF"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CD2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FEmpInfo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkPayFullSalary"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtFilePath"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBrowseBarcode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "General Info"
      TabPicture(1)   =   "NewEmployee.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "FHistory"
      Tab(1).Control(2)=   "FRefInfo"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Employee Info"
      TabPicture(2)   =   "NewEmployee.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FGInfo"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Job Duties"
      TabPicture(3)   =   "NewEmployee.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Qualification"
      TabPicture(4)   =   "NewEmployee.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Settings"
      TabPicture(5)   =   "NewEmployee.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame4"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Hidden"
      TabPicture(6)   =   "NewEmployee.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Stxts(9)"
      Tab(6).Control(1)=   "chkMultiple"
      Tab(6).Control(2)=   "fraHoliday(3)"
      Tab(6).Control(3)=   "fraHoliday(2)"
      Tab(6).Control(4)=   "fraHoliday(1)"
      Tab(6).Control(5)=   "fraHoliday(0)"
      Tab(6).Control(6)=   "chkLateHrs"
      Tab(6).Control(7)=   "chkAbsents"
      Tab(6).Control(8)=   "Frame3"
      Tab(6).Control(9)=   "DTLunch"
      Tab(6).Control(10)=   "label1(37)"
      Tab(6).Control(11)=   "label1(38)"
      Tab(6).ControlCount=   12
      TabCaption(7)   =   "Guarantor List"
      TabPicture(7)   =   "NewEmployee.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame12"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Employment History"
      TabPicture(8)   =   "NewEmployee.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame9"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Assets"
      TabPicture(9)   =   "NewEmployee.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame7"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Medical Info."
      TabPicture(10)  =   "NewEmployee.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "FOptions(4)"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Family Info."
      TabPicture(11)  =   "NewEmployee.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "FOptions(1)"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "Successions"
      TabPicture(12)  =   "NewEmployee.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Frame11"
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "Reviews"
      TabPicture(13)  =   "NewEmployee.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "FOptions(0)"
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "Notices"
      TabPicture(14)  =   "NewEmployee.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "FOptions(2)"
      Tab(14).ControlCount=   1
      TabCaption(15)  =   "Salary Change History"
      TabPicture(15)  =   "NewEmployee.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Frame13"
      Tab(15).ControlCount=   1
      TabCaption(16)  =   "Targets"
      TabPicture(16)  =   "NewEmployee.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Frame18"
      Tab(16).ControlCount=   1
      TabCaption(17)  =   "Costing Details"
      TabPicture(17)  =   "NewEmployee.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Frame19"
      Tab(17).ControlCount=   1
      Begin VB.CommandButton cmdBrowseBarcode 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8820
         TabIndex        =   340
         Top             =   5565
         Width           =   405
      End
      Begin VB.TextBox txtFilePath 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3135
         TabIndex        =   339
         Top             =   5580
         Width           =   5625
      End
      Begin VB.CheckBox chkPayFullSalary 
         Caption         =   "Pay Full Salary"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   1215
         TabIndex        =   338
         Top             =   5370
         Width           =   1815
      End
      Begin VB.Frame Frame19 
         Height          =   5430
         Left            =   -74895
         TabIndex        =   331
         Top             =   675
         Width           =   10065
         Begin ComboList.Usercontrol1 cmbProcessCD 
            Height          =   285
            Left            =   90
            TabIndex        =   332
            Top             =   375
            Width           =   4680
            _ExtentX        =   8255
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
         Begin ComboList.Usercontrol1 cmbItemGroupCD 
            Height          =   285
            Left            =   4785
            TabIndex        =   333
            Top             =   375
            Width           =   4410
            _ExtentX        =   7779
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
         Begin VB.CommandButton cmdAddCosting 
            Caption         =   "Add"
            Height          =   450
            Left            =   9210
            TabIndex        =   337
            Top             =   210
            Width           =   750
         End
         Begin MSComctlLib.ListView LVCD 
            Height          =   4665
            Left            =   90
            TabIndex        =   334
            Top             =   690
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   8229
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Extra"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Process"
               Object.Width           =   8255
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Group"
               Object.Width           =   7779
            EndProperty
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Process"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   114
            Left            =   90
            TabIndex        =   336
            Top             =   165
            Width           =   555
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Item Group"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   113
            Left            =   4785
            TabIndex        =   335
            Top             =   180
            Width           =   810
         End
      End
      Begin VB.Frame Frame18 
         Height          =   5535
         Left            =   -74910
         TabIndex        =   319
         Top             =   615
         Width           =   10095
         Begin VB.CommandButton cmdDeleteAll 
            Caption         =   "Delete All"
            Height          =   330
            Left            =   1215
            TabIndex        =   342
            Top             =   5160
            Width           =   1095
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "&Import"
            Height          =   330
            Left            =   90
            TabIndex        =   330
            Top             =   5160
            Width           =   1095
         End
         Begin VB.CheckBox chkDTOverTime 
            Caption         =   "Over Time"
            Height          =   195
            Left            =   7290
            TabIndex        =   329
            Top             =   210
            Width           =   1140
         End
         Begin VB.CheckBox chkGroupofItems 
            Caption         =   "Group of Items"
            Height          =   195
            Left            =   3360
            TabIndex        =   328
            Top             =   210
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.TextBox txtDTQtyEdit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   327
            Top             =   255
            Visible         =   0   'False
            Width           =   690
         End
         Begin ComboList.Usercontrol1 cmbDTProc 
            Height          =   285
            Left            =   105
            TabIndex        =   320
            Top             =   420
            Width           =   2805
            _ExtentX        =   4948
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
         Begin ComboList.Usercontrol1 cmbDTItem 
            Height          =   285
            Left            =   2940
            TabIndex        =   322
            Top             =   420
            Width           =   6330
            _ExtentX        =   11165
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
         Begin MSComctlLib.ListView LVDTItems 
            Height          =   4410
            Left            =   105
            TabIndex        =   321
            Top             =   735
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   7779
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   4948
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Group"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Quantity"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Over Time"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Group"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   112
            Left            =   2940
            TabIndex        =   326
            Top             =   225
            Width           =   435
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Qty"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   111
            Left            =   9285
            TabIndex        =   325
            Top             =   210
            Width           =   270
         End
         Begin MSForms.TextBox txtDTQty 
            Height          =   285
            Left            =   9300
            TabIndex        =   323
            Top             =   420
            Width           =   690
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "1217;503"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Process"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   110
            Left            =   105
            TabIndex        =   324
            Top             =   210
            Width           =   555
         End
      End
      Begin VB.Frame Frame13 
         Height          =   5280
         Left            =   -73950
         TabIndex        =   252
         Top             =   720
         Width           =   7260
         Begin MSComctlLib.ListView LVSalHistory 
            Height          =   4950
            Left            =   90
            TabIndex        =   253
            Top             =   210
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   8731
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   6376
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Salary"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin VB.Frame FOptions 
         Height          =   4710
         Index           =   2
         Left            =   -74430
         TabIndex        =   296
         Top             =   1290
         Width           =   9150
         Begin VB.CommandButton cmdAddWarning 
            Caption         =   "Add Notice"
            Height          =   315
            Left            =   7500
            TabIndex        =   297
            Top             =   255
            Width           =   1545
         End
         Begin MSComctlLib.ListView LVNotices 
            Height          =   4035
            Left            =   135
            TabIndex        =   298
            Top             =   600
            Width           =   8910
            _ExtentX        =   15716
            _ExtentY        =   7117
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr No."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Notice By"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Notice"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.Frame FOptions 
         Height          =   5385
         Index           =   0
         Left            =   -74895
         TabIndex        =   293
         Top             =   1005
         Width           =   10050
         Begin VB.CommandButton cmdAddReview 
            Caption         =   "Add Review"
            Height          =   315
            Left            =   8340
            TabIndex        =   294
            Top             =   255
            Width           =   1545
         End
         Begin MSComctlLib.ListView LVReview 
            Height          =   4710
            Left            =   210
            TabIndex        =   295
            Top             =   585
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   8308
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr No."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Type"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Rating"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Reviewed By"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Review"
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin VB.Frame Frame11 
         Height          =   5280
         Left            =   -73530
         TabIndex        =   286
         Top             =   1080
         Width           =   7260
         Begin VB.CommandButton cmdAddSuccession 
            Caption         =   "Add"
            Height          =   345
            Left            =   6405
            TabIndex        =   289
            Top             =   570
            Width           =   705
         End
         Begin VB.TextBox txtSuccession 
            Height          =   285
            Index           =   0
            Left            =   3225
            Locked          =   -1  'True
            TabIndex        =   287
            Top             =   615
            Width           =   3120
         End
         Begin ComboList.Usercontrol1 cmbSuccessionEmpID 
            Height          =   285
            Left            =   90
            TabIndex        =   288
            Top             =   615
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   556
            BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ListIndex       =   -1
         End
         Begin MSComctlLib.ListView LVSuccessions 
            Height          =   4200
            Left            =   90
            TabIndex        =   290
            Top             =   930
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   7408
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Employee"
               Object.Width           =   4128
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Designation"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Employee"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   102
            Left            =   105
            TabIndex        =   292
            Top             =   405
            Width           =   690
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   100
            Left            =   3255
            TabIndex        =   291
            Top             =   405
            Width           =   840
         End
      End
      Begin VB.Frame FOptions 
         Height          =   5115
         Index           =   1
         Left            =   -74385
         TabIndex        =   273
         Top             =   1245
         Width           =   9150
         Begin VB.Frame Frame10 
            Caption         =   "Dependant Family Members :"
            Height          =   2460
            Left            =   75
            TabIndex        =   278
            Top             =   120
            Width           =   9000
            Begin VB.CommandButton cmdAddChildren 
               Caption         =   "Add"
               Height          =   255
               Left            =   7980
               TabIndex        =   279
               Top             =   180
               Width           =   930
            End
            Begin MSComctlLib.ListView LVChildren 
               Height          =   1860
               Left            =   90
               TabIndex        =   280
               Top             =   510
               Width           =   8805
               _ExtentX        =   15531
               _ExtentY        =   3281
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   10
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Name"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Relation"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "N.I.C. No"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Date Of Birth"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Age"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Marital Status"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Education Level"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Occupation"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "Class Name"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Institute"
                  Object.Width           =   2646
               EndProperty
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Non Dependant Family Members :"
            Height          =   2460
            Left            =   75
            TabIndex        =   274
            Top             =   2580
            Width           =   9000
            Begin VB.CommandButton cmdAddBro 
               Caption         =   "Add"
               Height          =   255
               Left            =   7980
               TabIndex        =   275
               Top             =   195
               Visible         =   0   'False
               Width           =   930
            End
            Begin MSComCtl2.DTPicker DTEdit 
               Height          =   330
               Left            =   4650
               TabIndex        =   276
               Top             =   720
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   582
               _Version        =   393216
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   151977987
               CurrentDate     =   36891
            End
            Begin MSComctlLib.ListView LVBro 
               Height          =   1860
               Left            =   90
               TabIndex        =   277
               Top             =   540
               Width           =   8805
               _ExtentX        =   15531
               _ExtentY        =   3281
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   10
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Name"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Relation"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "N.I.C. No"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Date Of Birth"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Age"
                  Object.Width           =   4939
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Marital Status"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Education"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Occupation"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "Class Name"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Institute"
                  Object.Width           =   0
               EndProperty
            End
         End
      End
      Begin VB.Frame FOptions 
         Height          =   5115
         Index           =   4
         Left            =   -74460
         TabIndex        =   254
         Top             =   1215
         Width           =   9150
         Begin VB.Frame Frame16 
            Caption         =   "Medical History"
            Height          =   1560
            Left            =   135
            TabIndex        =   258
            Top             =   225
            Width           =   8835
            Begin VB.ComboBox cmbHBsAg 
               Height          =   315
               ItemData        =   "NewEmployee.frx":01F8
               Left            =   150
               List            =   "NewEmployee.frx":0205
               Style           =   2  'Dropdown List
               TabIndex        =   263
               Top             =   465
               Width           =   1575
            End
            Begin VB.ComboBox cmbAntiHcv 
               Height          =   315
               ItemData        =   "NewEmployee.frx":022B
               Left            =   1755
               List            =   "NewEmployee.frx":0238
               Style           =   2  'Dropdown List
               TabIndex        =   262
               Top             =   465
               Width           =   1575
            End
            Begin VB.ComboBox cmbBloodSugar 
               Height          =   315
               ItemData        =   "NewEmployee.frx":025E
               Left            =   3360
               List            =   "NewEmployee.frx":026B
               Style           =   2  'Dropdown List
               TabIndex        =   261
               Top             =   465
               Width           =   1575
            End
            Begin VB.ComboBox cmbBloodPressure 
               Height          =   315
               ItemData        =   "NewEmployee.frx":0291
               Left            =   4965
               List            =   "NewEmployee.frx":029E
               Style           =   2  'Dropdown List
               TabIndex        =   260
               Top             =   465
               Width           =   1575
            End
            Begin VB.ComboBox cmbAllergy 
               Height          =   315
               ItemData        =   "NewEmployee.frx":02C4
               Left            =   150
               List            =   "NewEmployee.frx":02CE
               Style           =   2  'Dropdown List
               TabIndex        =   259
               Top             =   1035
               Width           =   1575
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "HBsAg"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   96
               Left            =   165
               TabIndex        =   272
               Top             =   240
               Width           =   465
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "AntiHcv"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   89
               Left            =   1770
               TabIndex        =   271
               Top             =   240
               Width           =   555
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Blood Sugar"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   88
               Left            =   3375
               TabIndex        =   270
               Top             =   240
               Width           =   855
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Blood Pressure"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   87
               Left            =   4965
               TabIndex        =   269
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Blood Group"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   86
               Left            =   6585
               TabIndex        =   268
               Top             =   240
               Width           =   870
            End
            Begin MSForms.ComboBox cmbBloodGroup 
               Height          =   315
               Left            =   6570
               TabIndex        =   267
               Tag             =   "BloodGroup"
               Top             =   465
               Width           =   2190
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "3863;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Allergy"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   85
               Left            =   165
               TabIndex        =   266
               Top             =   810
               Width           =   495
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   38
               Left            =   1755
               TabIndex        =   265
               Top             =   1035
               Visible         =   0   'False
               Width           =   7005
               VariousPropertyBits=   679495707
               Size            =   "12356;556"
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Allergy Description"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   101
               Left            =   1740
               TabIndex        =   264
               Top             =   810
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "Medical Check-Ups :"
            Height          =   3120
            Left            =   135
            TabIndex        =   255
            Top             =   1875
            Width           =   8835
            Begin VB.CommandButton cmdAddMC 
               Caption         =   "Add"
               Height          =   255
               Left            =   7785
               TabIndex        =   256
               Top             =   180
               Width           =   930
            End
            Begin MSComctlLib.ListView LVMC 
               Height          =   2565
               Left            =   75
               TabIndex        =   257
               Top             =   480
               Width           =   8670
               _ExtentX        =   15293
               _ExtentY        =   4524
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
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
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Sr No"
                  Object.Width           =   1058
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Medical Test"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Lab"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Date"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Status"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Treatment"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Remarks"
                  Object.Width           =   2469
               EndProperty
            End
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Assets :"
         Height          =   4785
         Left            =   -74430
         TabIndex        =   238
         Top             =   1410
         Width           =   9000
         Begin VB.TextBox txtAsset 
            Height          =   315
            Left            =   765
            TabIndex        =   245
            Top             =   405
            Width           =   3720
         End
         Begin VB.TextBox txtAssetSNo 
            Height          =   315
            Left            =   60
            TabIndex        =   244
            Text            =   "1"
            Top             =   405
            Width           =   675
         End
         Begin VB.TextBox txtAssetRemarks 
            Height          =   315
            Left            =   5910
            TabIndex        =   243
            Top             =   390
            Width           =   2985
         End
         Begin VB.PictureBox PicEditAssets 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   60
            ScaleHeight     =   255
            ScaleWidth      =   8085
            TabIndex        =   239
            Top             =   1050
            Visible         =   0   'False
            Width           =   8115
            Begin VB.TextBox txtEditAssets 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   5115
               TabIndex        =   242
               Top             =   -15
               Width           =   2985
            End
            Begin VB.TextBox txtEditAssets 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   -15
               TabIndex        =   241
               Top             =   -15
               Width           =   3720
            End
            Begin VB.TextBox txtEditAssets 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   3705
               TabIndex        =   240
               Top             =   -15
               Width           =   1410
            End
         End
         Begin MSComctlLib.ListView LVAssets 
            Height          =   3960
            Left            =   60
            TabIndex        =   246
            Top             =   735
            Width           =   8850
            _ExtentX        =   15610
            _ExtentY        =   6985
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
               Text            =   "Sr. No"
               Object.Width           =   1191
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Asset Description"
               Object.Width           =   6562
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   2487
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Remarks"
               Object.Width           =   5265
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTAsset 
            Height          =   315
            Left            =   4500
            TabIndex        =   247
            Top             =   405
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   152109059
            CurrentDate     =   36891
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "S # :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   84
            Left            =   75
            TabIndex        =   251
            Top             =   195
            Width           =   360
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Asset Description"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   83
            Left            =   780
            TabIndex        =   250
            Top             =   195
            Width           =   1245
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Remarks"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   82
            Left            =   5925
            TabIndex        =   249
            Top             =   195
            Width           =   615
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   80
            Left            =   4515
            TabIndex        =   248
            Top             =   195
            Width           =   345
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4680
         Left            =   -74835
         TabIndex        =   223
         Top             =   1320
         Width           =   10050
         Begin VB.TextBox txtHistory 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   6195
            TabIndex        =   228
            Top             =   510
            Width           =   2985
         End
         Begin VB.CommandButton cmdAddHistory 
            Caption         =   "Add"
            Height          =   345
            Left            =   9240
            TabIndex        =   227
            Top             =   465
            Width           =   705
         End
         Begin VB.TextBox txtHistory 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   2565
            TabIndex        =   226
            Top             =   510
            Width           =   1665
         End
         Begin VB.TextBox txtHistory 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   4260
            TabIndex        =   225
            Top             =   510
            Width           =   945
         End
         Begin VB.TextBox txtHistory 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   5235
            TabIndex        =   224
            Top             =   510
            Width           =   945
         End
         Begin MSComCtl2.DTPicker DTHistory 
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   229
            Top             =   510
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   152109059
            CurrentDate     =   39827
         End
         Begin MSComctlLib.ListView LVHistory 
            Height          =   3630
            Left            =   90
            TabIndex        =   230
            Top             =   930
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   6403
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "From"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "To"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Company"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Designation"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Industry"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Responsibilities"
               Object.Width           =   4586
            EndProperty
         End
         Begin MSComCtl2.DTPicker DTHistory 
            Height          =   315
            Index           =   1
            Left            =   1335
            TabIndex        =   231
            Top             =   510
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   152109059
            CurrentDate     =   39827
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Responsibilities"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   90
            Left            =   6210
            TabIndex        =   237
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "From Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   91
            Left            =   90
            TabIndex        =   236
            Top             =   300
            Width           =   750
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "To Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   92
            Left            =   1335
            TabIndex        =   235
            Top             =   300
            Width           =   570
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Company"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   93
            Left            =   2580
            TabIndex        =   234
            Top             =   300
            Width           =   675
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   94
            Left            =   4275
            TabIndex        =   233
            Top             =   300
            Width           =   840
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Industry"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   95
            Left            =   5250
            TabIndex        =   232
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5085
         Left            =   -74355
         TabIndex        =   218
         Top             =   1020
         Width           =   9135
         Begin VB.Frame Frame15 
            Caption         =   "Address :"
            Height          =   2415
            Left            =   105
            TabIndex        =   317
            Top             =   2550
            Width           =   4365
            Begin MSForms.TextBox TBox 
               Height          =   2085
               Index           =   49
               Left            =   60
               TabIndex        =   318
               Top             =   240
               Width           =   4215
               VariousPropertyBits=   -1467987941
               ScrollBars      =   2
               Size            =   "7435;3678"
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Comments :"
            Height          =   2415
            Left            =   4650
            TabIndex        =   315
            Top             =   2550
            Width           =   4365
            Begin MSForms.TextBox TBox 
               Height          =   2085
               Index           =   48
               Left            =   60
               TabIndex        =   316
               Top             =   240
               Width           =   4215
               VariousPropertyBits=   -1467987941
               ScrollBars      =   2
               Size            =   "7435;3678"
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin ComboList.Usercontrol1 cmbGuarantors 
            Height          =   285
            Left            =   105
            TabIndex        =   300
            Top             =   210
            Width           =   8910
            _ExtentX        =   15716
            _ExtentY        =   556
            BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextBackColor   =   12640511
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ListIndex       =   -1
            Appearance      =   0
         End
         Begin VB.CommandButton cmdAddTraining 
            Caption         =   "Add"
            Height          =   360
            Index           =   0
            Left            =   8115
            TabIndex        =   219
            Top             =   150
            Visible         =   0   'False
            Width           =   885
         End
         Begin MSComctlLib.ListView LVGuarantors 
            Height          =   300
            Left            =   105
            TabIndex        =   220
            Top             =   4155
            Visible         =   0   'False
            Width           =   8910
            _ExtentX        =   15716
            _ExtentY        =   529
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
               Text            =   "Sr. No"
               Object.Width           =   1191
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Name"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Address"
               Object.Width           =   4233
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Phone"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Known By"
               Object.Width           =   2822
            EndProperty
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Name :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   109
            Left            =   135
            TabIndex        =   314
            Top             =   660
            Width           =   510
         End
         Begin MSForms.TextBox TBox 
            Height          =   390
            Index           =   47
            Left            =   105
            TabIndex        =   313
            Top             =   885
            Width           =   4440
            VariousPropertyBits=   679495707
            Size            =   "7832;688"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox TBox 
            Height          =   390
            Index           =   46
            Left            =   4575
            TabIndex        =   312
            Top             =   885
            Width           =   4440
            VariousPropertyBits=   679495707
            Size            =   "7832;688"
            FontName        =   "Tahoma"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Father Name :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   108
            Left            =   4575
            TabIndex        =   311
            Top             =   675
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   45
            Left            =   6810
            TabIndex        =   310
            Top             =   1560
            Width           =   2205
            VariousPropertyBits=   679495711
            Size            =   "3889;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Cell :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   107
            Left            =   6750
            TabIndex        =   309
            Top             =   1350
            Width           =   360
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   44
            Left            =   4575
            TabIndex        =   308
            Top             =   1560
            Width           =   2205
            VariousPropertyBits=   679495711
            Size            =   "3889;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   106
            Left            =   4575
            TabIndex        =   307
            Top             =   1350
            Width           =   555
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   43
            Left            =   105
            TabIndex        =   306
            Top             =   1560
            Width           =   2205
            VariousPropertyBits=   679495711
            Size            =   "3889;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "N.I.C. # : (Nadra Format)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   105
            Left            =   165
            TabIndex        =   305
            Top             =   1320
            Width           =   1875
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   42
            Left            =   2340
            TabIndex        =   304
            Top             =   1560
            Width           =   2205
            VariousPropertyBits=   679495711
            Size            =   "3889;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Cast"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   104
            Left            =   2340
            TabIndex        =   303
            Top             =   1335
            Width           =   330
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   41
            Left            =   105
            TabIndex        =   302
            Top             =   2130
            Width           =   8910
            VariousPropertyBits=   679495711
            Size            =   "15716;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Known By "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   103
            Left            =   180
            TabIndex        =   301
            Top             =   1905
            Width           =   750
         End
      End
      Begin VB.TextBox Stxts 
         Height          =   285
         Index           =   9
         Left            =   -73575
         TabIndex        =   195
         Text            =   "08:00"
         Top             =   4740
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CheckBox chkMultiple 
         Caption         =   "Allow Multiple Entries"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74925
         TabIndex        =   194
         Top             =   5295
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Frame fraHoliday 
         Height          =   705
         Index           =   3
         Left            =   -69645
         TabIndex        =   190
         Top             =   3090
         Visible         =   0   'False
         Width           =   2355
         Begin VB.OptionButton optPublic 
            Caption         =   "Yes"
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   960
            TabIndex        =   192
            Top             =   360
            Width           =   675
         End
         Begin VB.OptionButton optPublic 
            Caption         =   "No"
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   1680
            TabIndex        =   191
            Top             =   360
            Width           =   555
         End
         Begin VB.Label label1 
            Caption         =   "Deduct Absent After Public Holiday"
            ForeColor       =   &H00800000&
            Height          =   405
            Index           =   40
            Left            =   90
            TabIndex        =   193
            Top             =   120
            Width           =   2205
         End
      End
      Begin VB.Frame fraHoliday 
         Height          =   705
         Index           =   2
         Left            =   -69645
         TabIndex        =   186
         Top             =   2505
         Visible         =   0   'False
         Width           =   2355
         Begin VB.OptionButton optPublic 
            Caption         =   "No"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   188
            Top             =   390
            Width           =   615
         End
         Begin VB.OptionButton optPublic 
            Caption         =   "Yes"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   187
            Top             =   390
            Width           =   675
         End
         Begin VB.Label label1 
            Caption         =   "Deduct Absent Before Public Holiday"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   39
            Left            =   90
            TabIndex        =   189
            Top             =   150
            Width           =   2115
         End
      End
      Begin VB.Frame fraHoliday 
         Height          =   675
         Index           =   1
         Left            =   -69645
         TabIndex        =   182
         Top             =   1965
         Visible         =   0   'False
         Width           =   2355
         Begin VB.OptionButton optGeneral 
            Caption         =   "No"
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   1680
            TabIndex        =   184
            Top             =   330
            Width           =   615
         End
         Begin VB.OptionButton optGeneral 
            Caption         =   "Yes"
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   960
            TabIndex        =   183
            Top             =   330
            Width           =   675
         End
         Begin VB.Label label1 
            Caption         =   "Deduct Absent After Sunday"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   34
            Left            =   90
            TabIndex        =   185
            Top             =   120
            Width           =   2205
         End
      End
      Begin VB.Frame fraHoliday 
         Height          =   675
         Index           =   0
         Left            =   -69645
         TabIndex        =   178
         Top             =   1410
         Visible         =   0   'False
         Width           =   2355
         Begin VB.OptionButton optGeneral 
            Caption         =   "Yes"
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   180
            Top             =   330
            Width           =   675
         End
         Begin VB.OptionButton optGeneral 
            Caption         =   "No"
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   179
            Top             =   330
            Width           =   615
         End
         Begin VB.Label label1 
            Caption         =   "Deduct Absent Before Sunday"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   41
            Left            =   60
            TabIndex        =   181
            Top             =   150
            Width           =   2265
         End
      End
      Begin VB.CheckBox chkLateHrs 
         Caption         =   "Deduct Late Hours"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -69555
         TabIndex        =   177
         Top             =   4320
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CheckBox chkAbsents 
         Caption         =   "Deduct Absents"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -69555
         TabIndex        =   176
         Top             =   3990
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Frame Frame3 
         Height          =   3405
         Left            =   -74940
         TabIndex        =   151
         Top             =   1275
         Visible         =   0   'False
         Width           =   5085
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   1
            Left            =   3540
            TabIndex        =   161
            Text            =   "0"
            Top             =   420
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   2
            Left            =   3540
            TabIndex        =   160
            Text            =   "0"
            Top             =   1080
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   3
            Left            =   3540
            TabIndex        =   159
            Text            =   "0"
            Top             =   1410
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   158
            Top             =   1740
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   5
            Left            =   3540
            TabIndex        =   157
            Text            =   "0"
            Top             =   2070
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   6
            Left            =   3540
            TabIndex        =   156
            Text            =   "0"
            Top             =   2400
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   7
            Left            =   3540
            TabIndex        =   155
            Text            =   "0"
            Top             =   2700
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   8
            Left            =   3540
            TabIndex        =   154
            Text            =   "0"
            Top             =   3030
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   10
            Left            =   3540
            TabIndex        =   153
            Text            =   "0"
            Top             =   1740
            Width           =   1185
         End
         Begin VB.TextBox Stxts 
            Height          =   315
            Index           =   0
            Left            =   3540
            TabIndex        =   152
            Text            =   "0"
            Top             =   120
            Width           =   1185
         End
         Begin MSComCtl2.DTPicker dtLHours 
            Height          =   315
            Left            =   3540
            TabIndex        =   162
            Top             =   750
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   152109059
            UpDown          =   -1  'True
            CurrentDate     =   37683
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            Caption         =   "How Many Abs.(Abs. Without Pay) Allowed :"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   43
            Left            =   240
            TabIndex        =   175
            Top             =   210
            Width           =   3225
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Over Time Hrs Per Day  :"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   44
            Left            =   885
            TabIndex        =   174
            Top             =   1500
            Width           =   2580
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "How Many Times Late Time Allowed :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   45
            Left            =   825
            TabIndex        =   173
            Top             =   825
            Width           =   2640
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "How  Much Late Time Allowed :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   46
            Left            =   1245
            TabIndex        =   172
            Top             =   540
            Width           =   2220
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Social Security Amount :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   47
            Left            =   1740
            TabIndex        =   171
            Top             =   1830
            Width           =   1740
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Advance Salary Payable  :"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   48
            Left            =   900
            TabIndex        =   170
            Top             =   2445
            Width           =   2580
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Short Term Loan :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   49
            Left            =   2190
            TabIndex        =   169
            Top             =   2130
            Width           =   1290
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "After How Much Time Long Term Loan is App. :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   50
            Left            =   120
            TabIndex        =   168
            Top             =   2745
            Width           =   3360
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Long Term Loan Limit :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   51
            Left            =   1875
            TabIndex        =   167
            Top             =   3075
            Width           =   1605
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "How Many Times Late Time Allowed :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   52
            Left            =   810
            TabIndex        =   166
            Top             =   1140
            Width           =   2640
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   42
            Left            =   4815
            TabIndex        =   165
            Top             =   2205
            Width           =   165
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   53
            Left            =   4815
            TabIndex        =   164
            Top             =   2535
            Width           =   165
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mon"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   54
            Left            =   4740
            TabIndex        =   163
            Top             =   2835
            Width           =   300
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Emergency Contact Info"
         Height          =   1455
         Left            =   -73650
         TabIndex        =   126
         Top             =   3045
         Width           =   7635
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   63
            Left            =   4275
            TabIndex        =   131
            Top             =   255
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   840
            Index           =   31
            Left            =   4230
            TabIndex        =   26
            Top             =   480
            Width           =   3255
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "5741;1482"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone 2"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   62
            Left            =   2985
            TabIndex        =   130
            Top             =   810
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   30
            Left            =   2955
            TabIndex        =   25
            Top             =   1020
            Width           =   1200
            VariousPropertyBits=   679495707
            Size            =   "2117;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone 1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   61
            Left            =   2955
            TabIndex        =   129
            Top             =   270
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   29
            Left            =   2970
            TabIndex        =   24
            Top             =   480
            Width           =   1200
            VariousPropertyBits=   679495707
            Size            =   "2117;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Father's Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   60
            Left            =   105
            TabIndex        =   128
            Top             =   810
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   28
            Left            =   90
            TabIndex        =   23
            Top             =   1020
            Width           =   2835
            VariousPropertyBits=   679495707
            Size            =   "5001;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   300
            Index           =   27
            Left            =   90
            TabIndex        =   22
            Top             =   495
            Width           =   2835
            VariousPropertyBits=   679495707
            Size            =   "5001;529"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Emergency Contact Person"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   59
            Left            =   135
            TabIndex        =   127
            Top             =   270
            Width           =   1950
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Timings"
         Height          =   4530
         Left            =   -74160
         TabIndex        =   118
         Top             =   1500
         Width           =   7935
         Begin VB.Frame Frame6 
            Caption         =   "Payroll Catagory"
            Height          =   2910
            Left            =   75
            TabIndex        =   199
            Top             =   1515
            Width           =   7800
            Begin VB.ComboBox cmbCatagory 
               Height          =   315
               Left            =   1110
               Style           =   2  'Dropdown List
               TabIndex        =   201
               Top             =   270
               Width           =   1335
            End
            Begin MSComctlLib.ListView LVCat 
               Height          =   2160
               Left            =   105
               TabIndex        =   202
               Top             =   645
               Width           =   7590
               _ExtentX        =   13388
               _ExtentY        =   3810
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Setting"
                  Object.Width           =   4233
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Value"
                  Object.Width           =   4233
               EndProperty
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Catagory # :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   72
               Left            =   120
               TabIndex        =   200
               Top             =   315
               Width           =   945
            End
         End
         Begin MSComCtl2.DTPicker TExit 
            Height          =   285
            Left            =   4995
            TabIndex        =   60
            Top             =   450
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   152240131
            UpDown          =   -1  'True
            CurrentDate     =   37958.7083333333
         End
         Begin MSComCtl2.DTPicker TEnter 
            Height          =   285
            Left            =   1575
            TabIndex        =   59
            Top             =   450
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   152240131
            UpDown          =   -1  'True
            CurrentDate     =   37958.3333333333
         End
         Begin MSComCtl2.DTPicker LunchStart 
            Height          =   285
            Left            =   2715
            TabIndex        =   143
            Top             =   450
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   152240131
            UpDown          =   -1  'True
            CurrentDate     =   37683.5208333333
         End
         Begin MSComCtl2.DTPicker LunchEnd 
            Height          =   285
            Left            =   3855
            TabIndex        =   144
            Top             =   450
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   152240131
            UpDown          =   -1  'True
            CurrentDate     =   37683.5625
         End
         Begin MSComCtl2.DTPicker FridayBreakStart 
            Height          =   285
            Left            =   1575
            TabIndex        =   208
            Top             =   1065
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   152240131
            UpDown          =   -1  'True
            CurrentDate     =   37958.7083333333
         End
         Begin MSComCtl2.DTPicker FridayBreakEnd 
            Height          =   285
            Left            =   3090
            TabIndex        =   209
            Top             =   1065
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   152240131
            UpDown          =   -1  'True
            CurrentDate     =   37958.7083333333
         End
         Begin MSComCtl2.DTPicker FridayCheckOut 
            Height          =   285
            Left            =   4605
            TabIndex        =   210
            Top             =   1065
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   152240131
            UpDown          =   -1  'True
            CurrentDate     =   37958.7083333333
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Friday Check Out :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   76
            Left            =   4605
            TabIndex        =   207
            Top             =   870
            Width           =   1350
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Friday Break :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   75
            Left            =   3090
            TabIndex        =   206
            Top             =   870
            Width           =   1005
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Friday Break :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   74
            Left            =   1575
            TabIndex        =   205
            Top             =   870
            Width           =   1005
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Lunch Start :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   69
            Left            =   2715
            TabIndex        =   146
            Top             =   240
            Width           =   930
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Lunch End :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   68
            Left            =   3855
            TabIndex        =   145
            Top             =   240
            Width           =   840
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Check Out :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   36
            Left            =   4995
            TabIndex        =   120
            Top             =   225
            Width           =   855
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Check In :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   35
            Left            =   1575
            TabIndex        =   119
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Define Employee Duties Below"
         Height          =   4365
         Left            =   -74010
         TabIndex        =   116
         Top             =   1635
         Width           =   7965
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   6600
            TabIndex        =   55
            Top             =   510
            Width           =   675
         End
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   5910
            MaxLength       =   1
            TabIndex        =   54
            Top             =   510
            Width           =   675
         End
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   2880
            TabIndex        =   52
            Top             =   510
            Width           =   1755
         End
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   570
            TabIndex        =   51
            Top             =   510
            Width           =   2295
         End
         Begin VB.TextBox txts 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "1"
            Top             =   510
            Width           =   465
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   315
            Left            =   4650
            TabIndex        =   53
            Top             =   510
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   152240131
            UpDown          =   -1  'True
            CurrentDate     =   37683
         End
         Begin MSComctlLib.ListView LVQ 
            Height          =   3345
            Left            =   90
            TabIndex        =   57
            Top             =   930
            Width           =   7740
            _ExtentX        =   13653
            _ExtentY        =   5900
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr No"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Diploma/Degree/Certificate"
               Object.Width           =   4128
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Institute"
               Object.Width           =   3120
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Passing Year"
               Object.Width           =   2194
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Grade"
               Object.Width           =   1270
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Division"
               Object.Width           =   1270
            EndProperty
         End
         Begin MSForms.CommandButton cmdPost 
            Height          =   315
            Left            =   7260
            TabIndex        =   56
            Top             =   510
            Width           =   630
            PicturePosition =   327683
            Size            =   "1111;556"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label label1 
            Caption         =   "Sr No:  Diploma/Degree/Certificate :     Name Of Institute :       Passing Year :      Grade :   Division :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   30
            Left            =   90
            TabIndex        =   117
            Top             =   270
            Width           =   7305
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Define Employee Duties Below"
         Height          =   4365
         Left            =   -73950
         TabIndex        =   115
         Top             =   1620
         Width           =   7905
         Begin VB.TextBox txtSrNo 
            BackColor       =   &H00E7EBEF&
            Height          =   315
            Left            =   90
            TabIndex        =   47
            Text            =   "1"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtDuty 
            BackColor       =   &H00E7EBEF&
            Height          =   315
            Left            =   750
            TabIndex        =   48
            Top             =   240
            Width           =   7095
         End
         Begin MSComctlLib.ListView LV 
            Height          =   3705
            Left            =   90
            TabIndex        =   49
            Top             =   570
            Width           =   7740
            _ExtentX        =   13653
            _ExtentY        =   6535
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr. No"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Description"
               Object.Width           =   12347
            EndProperty
         End
      End
      Begin VB.Frame FGInfo 
         Height          =   5295
         Left            =   -73350
         TabIndex        =   104
         Top             =   795
         Width           =   7095
         Begin ComboList.Usercontrol1 cmbRelEmp 
            Height          =   285
            Left            =   3825
            TabIndex        =   285
            Top             =   225
            Width           =   3090
            _ExtentX        =   5450
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
            TextForeColor   =   12582912
            ListIndex       =   -1
         End
         Begin VB.ComboBox cmbDesignation 
            Height          =   315
            ItemData        =   "NewEmployee.frx":02E5
            Left            =   1680
            List            =   "NewEmployee.frx":02E7
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   945
            Width           =   2130
         End
         Begin VB.CheckBox chkBankPymt 
            Alignment       =   1  'Right Justify
            Caption         =   "Salary is Paid Through Bank"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4620
            TabIndex        =   36
            Top             =   1950
            Width           =   2310
         End
         Begin VB.ComboBox cmbEmpType 
            Height          =   315
            ItemData        =   "NewEmployee.frx":02E9
            Left            =   165
            List            =   "NewEmployee.frx":02EB
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   945
            Width           =   1515
         End
         Begin MSComCtl2.DTPicker JoinDate 
            Height          =   315
            Left            =   165
            TabIndex        =   32
            Top             =   1575
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   149356547
            CurrentDate     =   37383
         End
         Begin MSComCtl2.DTPicker RetireDate 
            Height          =   315
            Left            =   1425
            TabIndex        =   33
            Top             =   1575
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   149356547
            CurrentDate     =   37383
         End
         Begin MSComCtl2.DTPicker DTInsurable 
            Height          =   315
            Left            =   5475
            TabIndex        =   135
            Top             =   945
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   149356547
            CurrentDate     =   37383
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Emp.:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   99
            Left            =   3375
            TabIndex        =   284
            Top             =   285
            Width           =   420
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   39
            Left            =   2175
            TabIndex        =   283
            Top             =   225
            Width           =   1140
            VariousPropertyBits=   679495707
            Size            =   "2011;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Relation with Comp. Emp. :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   98
            Left            =   210
            TabIndex        =   282
            Top             =   285
            Width           =   1950
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Food Allowance"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   79
            Left            =   5535
            TabIndex        =   216
            Top             =   4635
            Width           =   1125
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   37
            Left            =   5535
            TabIndex        =   215
            Top             =   4830
            Width           =   1395
            VariousPropertyBits=   679495707
            Size            =   "2461;556"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Social Security Amt"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   78
            Left            =   1650
            TabIndex        =   214
            Top             =   4620
            Visible         =   0   'False
            Width           =   1365
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   36
            Left            =   1635
            TabIndex        =   213
            Top             =   4830
            Visible         =   0   'False
            Width           =   1395
            VariousPropertyBits=   679495707
            Size            =   "2461;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "EOBI Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   77
            Left            =   180
            TabIndex        =   212
            Top             =   4620
            Visible         =   0   'False
            Width           =   960
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   35
            Left            =   180
            TabIndex        =   211
            Top             =   4830
            Visible         =   0   'False
            Width           =   1395
            VariousPropertyBits=   679495707
            Size            =   "2461;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Fix. Allowance"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   73
            Left            =   5760
            TabIndex        =   204
            Top             =   1365
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   34
            Left            =   5775
            TabIndex        =   203
            Top             =   1575
            Width           =   1140
            VariousPropertyBits=   679495707
            Size            =   "2011;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Main Contractor"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   67
            Left            =   5280
            TabIndex        =   142
            Top             =   4020
            Width           =   1155
         End
         Begin MSForms.ComboBox cmbMainContractor 
            Height          =   300
            Left            =   5250
            TabIndex        =   141
            Top             =   4230
            Width           =   1755
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3096;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPerPiece 
            Height          =   345
            Left            =   5655
            TabIndex        =   140
            Top             =   3585
            Width           =   1395
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2461;609"
            Value           =   "0"
            Caption         =   "Per Piece Rate"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   300
            Index           =   33
            Left            =   3570
            TabIndex        =   139
            Top             =   1920
            Visible         =   0   'False
            Width           =   1005
            VariousPropertyBits=   679495707
            Size            =   "1773;529"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Account # :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   66
            Left            =   2655
            TabIndex        =   138
            Top             =   1980
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSForms.CheckBox chkArmy 
            Height          =   540
            Left            =   5655
            TabIndex        =   137
            Top             =   2715
            Width           =   1275
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2249;952"
            Value           =   "0"
            Caption         =   "Retired Army Personnel"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Insurable Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   65
            Left            =   5490
            TabIndex        =   136
            Top             =   750
            Width           =   1065
         End
         Begin MSForms.CheckBox chkLunch 
            Height          =   345
            Left            =   3705
            TabIndex        =   134
            Top             =   4215
            Width           =   1590
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2805;609"
            Value           =   "0"
            Caption         =   "Lunch Deduction"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkUnionFund 
            Height          =   345
            Left            =   165
            TabIndex        =   133
            Top             =   4215
            Visible         =   0   'False
            Width           =   2790
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "4921;609"
            Value           =   "0"
            Caption         =   "Union Member (Union Fund Ded.)"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Consider Over Time If Work Done is Greater Than :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   64
            Left            =   180
            TabIndex        =   132
            Top             =   3915
            Visible         =   0   'False
            Width           =   3675
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   32
            Left            =   3870
            TabIndex        =   45
            Top             =   3855
            Visible         =   0   'False
            Width           =   1005
            VariousPropertyBits=   679495707
            Size            =   "1773;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkOTByQTy 
            Height          =   315
            Left            =   150
            TabIndex        =   44
            Top             =   3585
            Width           =   3450
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "6085;556"
            Value           =   "0"
            Caption         =   "Over Time is Paid According To Work Done"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Social Security #"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   58
            Left            =   3750
            TabIndex        =   125
            Top             =   3045
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "EOBI #"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   57
            Left            =   1845
            TabIndex        =   124
            Top             =   3045
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Tax Amount"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   56
            Left            =   195
            TabIndex        =   123
            Top             =   3060
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Production Process"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   55
            Left            =   3825
            TabIndex        =   122
            Top             =   720
            Width           =   1365
         End
         Begin MSForms.ComboBox cmbProcess 
            Height          =   315
            Left            =   3825
            TabIndex        =   31
            Top             =   945
            Width           =   1635
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2884;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Join Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   165
            TabIndex        =   114
            Top             =   1350
            Width           =   675
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Retirement Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   1425
            TabIndex        =   113
            Top             =   1350
            Width           =   1185
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Salary (Ext)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   2895
            TabIndex        =   112
            Top             =   4260
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   15
            Left            =   2865
            TabIndex        =   71
            Top             =   4275
            Visible         =   0   'False
            Width           =   1080
            VariousPropertyBits=   679495707
            Size            =   "1905;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Remarks"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   195
            TabIndex        =   111
            Top             =   1965
            Width           =   615
         End
         Begin MSForms.TextBox TBox 
            Height          =   465
            Index           =   16
            Left            =   165
            TabIndex        =   37
            Top             =   2250
            Width           =   6765
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "11933;820"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkTaxable 
            Height          =   300
            Left            =   165
            TabIndex        =   38
            Top             =   2775
            Width           =   915
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1614;529"
            Value           =   "0"
            Caption         =   "Taxable"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkSocial 
            Height          =   300
            Left            =   3705
            TabIndex        =   42
            Top             =   2790
            Width           =   1950
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3440;529"
            Value           =   "0"
            Caption         =   "Social Security Holder"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkOldAge 
            Height          =   300
            Left            =   1815
            TabIndex        =   40
            Top             =   2775
            Width           =   1905
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3360;529"
            Value           =   "0"
            Caption         =   "Old Age Benifit Holder"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   17
            Left            =   4605
            TabIndex        =   35
            Top             =   1575
            Width           =   1140
            VariousPropertyBits=   679495707
            Size            =   "2011;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Salary"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   4635
            TabIndex        =   110
            Top             =   1350
            Width           =   450
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   1710
            TabIndex        =   109
            Top             =   735
            Width           =   840
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   2925
            TabIndex        =   108
            Top             =   1350
            Width           =   465
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   24
            Left            =   3720
            TabIndex        =   43
            Top             =   3255
            Visible         =   0   'False
            Width           =   1875
            VariousPropertyBits=   679495707
            Size            =   "3307;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   5
            Left            =   165
            TabIndex        =   39
            Top             =   3255
            Visible         =   0   'False
            Width           =   1635
            VariousPropertyBits=   679495707
            Size            =   "2884;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbStatus 
            Height          =   315
            Left            =   2895
            TabIndex        =   34
            Top             =   1575
            Width           =   1695
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2990;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   2
            Left            =   1830
            TabIndex        =   41
            Top             =   3255
            Visible         =   0   'False
            Width           =   1875
            VariousPropertyBits=   679495707
            Size            =   "3307;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   22
            Left            =   4740
            TabIndex        =   68
            Top             =   3210
            Visible         =   0   'False
            Width           =   855
            VariousPropertyBits=   679495707
            Size            =   "1508;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblSal 
            AutoSize        =   -1  'True
            Caption         =   "Salary :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4155
            TabIndex        =   107
            Top             =   3255
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lblOldSal 
            AutoSize        =   -1  'True
            Caption         =   "Salary :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1680
            TabIndex        =   106
            Top             =   3255
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   25
            Left            =   2910
            TabIndex        =   67
            Top             =   3210
            Visible         =   0   'False
            Width           =   795
            VariousPropertyBits=   679495707
            Size            =   "1402;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Employee Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   105
            Top             =   720
            Width           =   1095
         End
         Begin MSForms.CheckBox chkBonus 
            Height          =   315
            Left            =   3705
            TabIndex        =   46
            Top             =   3585
            Width           =   1335
            VariousPropertyBits=   1015023643
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2355;556"
            Value           =   "0"
            Caption         =   "Yearly Bonus"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame FHistory 
         Caption         =   "Previous Job History"
         Height          =   1380
         Left            =   -73650
         TabIndex        =   100
         Top             =   4560
         Width           =   7620
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   12
            Left            =   165
            TabIndex        =   27
            Top             =   405
            Width           =   3165
            VariousPropertyBits=   679495707
            Size            =   "5583;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Last Firm"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   180
            TabIndex        =   103
            Top             =   195
            Width           =   645
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   13
            Left            =   165
            TabIndex        =   28
            Top             =   945
            Width           =   3165
            VariousPropertyBits=   679495707
            Size            =   "5583;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Last Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   195
            TabIndex        =   102
            Top             =   720
            Width           =   1185
         End
         Begin MSForms.TextBox TBox 
            Height          =   870
            Index           =   14
            Left            =   3375
            TabIndex        =   29
            Top             =   405
            Width           =   4140
            VariousPropertyBits=   679495707
            Size            =   "7302;1535"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Reason To Discontinue"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   3420
            TabIndex        =   101
            Top             =   180
            Width           =   1635
         End
      End
      Begin VB.Frame FRefInfo 
         Caption         =   "Referal Information"
         Height          =   1455
         Left            =   -73620
         TabIndex        =   87
         Top             =   1575
         Width           =   7635
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Refered By (Name)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   135
            TabIndex        =   92
            Top             =   270
            Width           =   1380
         End
         Begin MSForms.TextBox TBox 
            Height          =   300
            Index           =   10
            Left            =   90
            TabIndex        =   17
            Top             =   495
            Width           =   2835
            VariousPropertyBits=   679495707
            Size            =   "5001;529"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   18
            Left            =   90
            TabIndex        =   18
            Top             =   1020
            Width           =   2835
            VariousPropertyBits=   679495707
            Size            =   "5001;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Father's Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   105
            TabIndex        =   91
            Top             =   810
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   19
            Left            =   2970
            TabIndex        =   19
            Top             =   480
            Width           =   1200
            VariousPropertyBits=   679495707
            Size            =   "2117;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone 1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   2955
            TabIndex        =   90
            Top             =   270
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   20
            Left            =   2955
            TabIndex        =   20
            Top             =   1020
            Width           =   1200
            VariousPropertyBits=   679495707
            Size            =   "2117;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone 2"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   2985
            TabIndex        =   89
            Top             =   810
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   840
            Index           =   21
            Left            =   4230
            TabIndex        =   21
            Top             =   480
            Width           =   3255
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "5741;1482"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   28
            Left            =   4275
            TabIndex        =   88
            Top             =   255
            Width           =   585
         End
      End
      Begin VB.Frame FEmpInfo 
         Height          =   4455
         Left            =   195
         TabIndex        =   75
         Top             =   840
         Width           =   9870
         Begin VB.ComboBox cmbProbationDays 
            Height          =   315
            ItemData        =   "NewEmployee.frx":02ED
            Left            =   1425
            List            =   "NewEmployee.frx":0306
            TabIndex        =   149
            Text            =   "Combo1"
            Top             =   3960
            Width           =   1050
         End
         Begin VB.ComboBox cmbBloodGroup1 
            Height          =   315
            Left            =   9015
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   255
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSComCtl2.DTPicker DOB 
            Height          =   315
            Left            =   1515
            TabIndex        =   6
            Top             =   1350
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   312475651
            CurrentDate     =   36891
            MaxDate         =   36891
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Maslak"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   116
            Left            =   6360
            TabIndex        =   346
            Top             =   1140
            Width           =   480
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   51
            Left            =   6330
            TabIndex        =   345
            Top             =   1350
            Width           =   1875
            VariousPropertyBits=   679495707
            Size            =   "3307;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Caste"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   115
            Left            =   6360
            TabIndex        =   344
            Top             =   420
            Width           =   420
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   50
            Left            =   6345
            TabIndex        =   343
            Top             =   630
            Width           =   1875
            VariousPropertyBits=   679495707
            Size            =   "3307;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   40
            Left            =   90
            TabIndex        =   299
            Top             =   630
            Width           =   1395
            VariousPropertyBits=   679495711
            Size            =   "2461;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Years"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   97
            Left            =   5805
            TabIndex        =   281
            Top             =   3960
            Width           =   405
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   9
            Left            =   4455
            TabIndex        =   222
            Tag             =   "S/O"
            Top             =   3915
            Width           =   1230
            VariousPropertyBits=   679495707
            Size            =   "2170;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Bond Period :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   81
            Left            =   3495
            TabIndex        =   221
            Top             =   3975
            Width           =   960
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Days"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   71
            Left            =   2505
            TabIndex        =   150
            Top             =   4005
            Width           =   360
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Probation Period :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   70
            Left            =   105
            TabIndex        =   148
            Top             =   3990
            Width           =   1290
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   26
            Left            =   2550
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   915
            VariousPropertyBits=   679495707
            Size            =   "1614;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Blood Group"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   9015
            TabIndex        =   121
            Top             =   30
            Visible         =   0   'False
            Width           =   870
         End
         Begin MSForms.ComboBox cmbMaritalStatus 
            Height          =   315
            Left            =   3930
            TabIndex        =   8
            Top             =   1350
            Width           =   2385
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "4207;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Thumb Exp/Signature"
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   9
            Left            =   8250
            TabIndex        =   99
            Top             =   2400
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Thumb Exp. /Signature"
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   8430
            TabIndex        =   74
            Top             =   2940
            Width           =   1215
         End
         Begin MSForms.CommandButton cmdCopy 
            Height          =   675
            Left            =   3870
            TabIndex        =   15
            Top             =   3045
            Width           =   555
            Caption         =   "Copy"
            PicturePosition =   65543
            Size            =   "979;1191"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.OptionButton SO 
            Height          =   345
            Index           =   0
            Left            =   3915
            TabIndex        =   2
            Top             =   270
            Width           =   675
            VariousPropertyBits=   1015023635
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "1191;609"
            Value           =   "1"
            Caption         =   "S/O"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton SO 
            Height          =   345
            Index           =   1
            Left            =   4620
            TabIndex        =   3
            Top             =   270
            Width           =   690
            VariousPropertyBits=   1015023635
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "1217;609"
            Value           =   "0"
            Caption         =   "D/O"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton SO 
            Height          =   345
            Index           =   2
            Left            =   5355
            TabIndex        =   4
            Top             =   270
            Width           =   735
            VariousPropertyBits=   1015023635
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "1296;609"
            Value           =   "0"
            Caption         =   "W/O"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbReligion 
            Height          =   315
            Left            =   2475
            TabIndex        =   11
            Top             =   2085
            Width           =   1965
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3466;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   990
            Index           =   8
            Left            =   90
            TabIndex        =   14
            Top             =   2880
            Width           =   3765
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "6641;1746"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Temporary Address"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   33
            Left            =   4440
            TabIndex        =   98
            Top             =   2640
            Width           =   1410
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Picture."
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   8490
            TabIndex        =   73
            Top             =   1305
            Width           =   1095
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Picture"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   8280
            TabIndex        =   97
            Top             =   420
            Width           =   495
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1740
            Left            =   8265
            Stretch         =   -1  'True
            Top             =   630
            Width           =   1545
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   1
            Left            =   3930
            TabIndex        =   5
            Tag             =   "S/O"
            Top             =   630
            Width           =   2385
            VariousPropertyBits=   679495707
            Size            =   "4207;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   1515
            TabIndex        =   1
            Top             =   630
            Width           =   2385
            VariousPropertyBits=   679495707
            Size            =   "4207;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1515
            TabIndex        =   86
            Top             =   420
            Width           =   405
         End
         Begin VB.Label lblEmpID 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1965
            TabIndex        =   65
            Top             =   75
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Employee ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   85
            Top             =   420
            Width           =   900
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Date Of Birth"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   1515
            TabIndex        =   84
            Top             =   1125
            Width           =   945
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Religion"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   2505
            TabIndex        =   83
            Top             =   1875
            Width           =   555
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Marital Status "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   3915
            TabIndex        =   82
            Top             =   1125
            Width           =   1035
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   11
            Left            =   3930
            TabIndex        =   70
            Top             =   1350
            Visible         =   0   'False
            Width           =   2385
            VariousPropertyBits=   679495707
            Size            =   "4207;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Permanent Address"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   105
            TabIndex        =   81
            Top             =   2640
            Width           =   1410
         End
         Begin MSForms.TextBox TBox 
            Height          =   990
            Index           =   23
            Left            =   4440
            TabIndex        =   16
            Top             =   2880
            Width           =   3765
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "6641;1746"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "CNIC"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   80
            Top             =   1875
            Width           =   375
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   90
            TabIndex        =   10
            Top             =   2085
            Width           =   2355
            VariousPropertyBits=   679495707
            Size            =   "4154;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbSex 
            Height          =   315
            Left            =   90
            TabIndex        =   66
            Top             =   1350
            Width           =   1395
            VariousPropertyBits=   679495711
            DisplayStyle    =   3
            Size            =   "2461;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   79
            Top             =   1125
            Width           =   270
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   3015
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1350
            Width           =   885
            VariousPropertyBits=   679495711
            Size            =   "1561;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Age"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   3045
            TabIndex        =   78
            Top             =   1125
            Width           =   285
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone 2"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   6360
            TabIndex        =   77
            Top             =   1875
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   7
            Left            =   6330
            TabIndex        =   13
            Top             =   2085
            Width           =   1875
            VariousPropertyBits=   679495707
            Size            =   "3307;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone 1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   4440
            TabIndex        =   76
            Top             =   1875
            Width           =   585
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   4440
            TabIndex        =   12
            Top             =   2085
            Width           =   1875
            VariousPropertyBits=   679495707
            Size            =   "3307;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1260
            Left            =   8250
            Stretch         =   -1  'True
            Top             =   2610
            Width           =   1575
         End
      End
      Begin MSComCtl2.DTPicker DTLunch 
         Height          =   285
         Left            =   -73680
         TabIndex        =   196
         Top             =   5565
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   312475651
         UpDown          =   -1  'True
         CurrentDate     =   37683.0416666667
      End
      Begin MSComDlg.CommonDialog CD2 
         Left            =   1995
         Top             =   5550
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "PDF Files Only|*.pdf"
      End
      Begin VB.Label lblCNICPDF 
         AutoSize        =   -1  'True
         Caption         =   "CNIC PDF File"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3150
         TabIndex        =   341
         Top             =   5295
         Width           =   1470
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lunch Duration :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   37
         Left            =   -74985
         TabIndex        =   198
         Top             =   5625
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Working Duration"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   38
         Left            =   -74940
         TabIndex        =   197
         Top             =   4830
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin MSForms.CheckBox chkExempt 
      Height          =   315
      Left            =   210
      TabIndex        =   58
      Top             =   7290
      Width           =   2865
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "5054;556"
      Value           =   "0"
      Caption         =   "Exempt From Dept. Settings"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox chkShowInExternal 
      Height          =   285
      Left            =   210
      TabIndex        =   61
      Top             =   7620
      Visible         =   0   'False
      Width           =   2955
      VariousPropertyBits=   1015023643
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "5212;503"
      Value           =   "1"
      Caption         =   "Show In External Salary Sheet"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   480
      Left            =   9000
      TabIndex        =   64
      Top             =   7380
      Width           =   1620
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2857;847"
      Accelerator     =   67
      FontName        =   "Segoe UI"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   480
      Left            =   5640
      TabIndex        =   62
      Top             =   7380
      Width           =   1620
      Caption         =   "Save & New   "
      PicturePosition =   327683
      Size            =   "2857;847"
      Accelerator     =   78
      FontName        =   "Segoe UI"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   480
      Left            =   7320
      TabIndex        =   63
      Top             =   7380
      Width           =   1620
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2857;847"
      Accelerator     =   83
      FontName        =   "Segoe UI"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Job"
      End
   End
   Begin VB.Menu mnupopup1 
      Caption         =   "popup1"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteQual 
         Caption         =   "Delete This Entry"
      End
   End
End
Attribute VB_Name = "NewEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelOption As Integer
Dim LastPoint As Integer
Public add As Boolean, edit As Boolean
Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Dim dontShowList As Boolean
Public SqlSettings As String
Dim PicFileName As String
Dim ThumbFileName As String
Public TableName As String

Dim DeleteJob As Boolean, DeleteCV As Boolean, CVID As String

Dim bTempEmp As Boolean
Dim strDelListEmpDT As String, strDelListEmpCD As String
Dim strCVID As String
Dim strFileName As String
Private Sub chkParent_Click()

    If chkParent = vbChecked Then
        TBox(1) = "0.00"
        TBox(1).Enabled = False
    Else
        TBox(1).Enabled = True
    End If
   
End Sub

Private Sub FillCmbs()

    cmbSex.AddItem "Male"
    cmbSex.AddItem "Female"
    
    cmbSex.ListIndex = 0
    
    Dim DeptTable As String
    If TableName = "Employees" Then
        DeptTable = "Departments"
    Else
        DeptTable = "ContDepartments"
    End If
    
    Call AddToCombo(cmbDeptID, "deptID", DeptTable, "Order By DeptID")
    Call AddToCombo(cmbDept, "Name", DeptTable, " Order By DeptID")
    
    Call AddToCombo(cmbMainContractor, "EmpID", "VEmp", " Where TempDept=1 AND Active=1")
    Call AddToCombo(cmbBloodGroup, "BloodGroup", "BloodGroups", " Order By BloodGroup")
    Call AddToCombo(cmbDesignation, "Designation", "Designations", " ORDER BY Designation")
    
    cmbBloodGroup.Text = "Unknown"
    
    Call AddToCombo(cmbCatagory, "EntryID", "EmployeeAttendanceCatagories", "")
    
    'Call AddToCombo(cmbSaleAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '31%' AND Not Parent Order By ACCTitle")
    'Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '41%' AND Not Parent Order By ACCTitle")
    'Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
        
    cmbRelEmp.ListHeight = 3000
    cmbRelEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", " ORDER BY EmpID"
        
        
    cmbSuccessionEmpID.ListHeight = 2400
    cmbSuccessionEmpID.AddVals con, "'{ ' + EmpID + ' } ' + Name", "Employees", "EmpID", " WHERE Active=1"
    
    cmbHBsAg.ListIndex = 2
    cmbAntiHcv.ListIndex = 2
    cmbBloodSugar.ListIndex = 2
    cmbBloodPressure.ListIndex = 2
    cmbAllergy.ListIndex = 0
        
    'Daily Targets Employee
    cmbDTProc.ListHeight = 3000
    cmbDTProc.AddVals con, "Description", "Processes", "ProcessID", " ORDER BY SNO"
    cmbDTItem.ListHeight = 3000
    'cmbDTItem.DropDownOffLine = True
    'cmbDTItem.AddVals con, "ItemID + '{' + Itemname + '}'", "Items", "ItemID", " WHERE InActive=0"
    cmbDTItem.AddVals con, "Description", "ItemGroups", "ID", " "
    
        
    cmbProcessCD.ListHeight = 3000
    cmbProcessCD.AddVals con, "Description", "Processes", "ProcessID", " ORDER BY SNO"
    
    cmbItemGroupCD.ListHeight = 3000
    cmbItemGroupCD.AddVals con, "Description", "ItemGroups", "ID"
    
End Sub

Private Sub chkBankPymt_Click()

    If chkBankPymt.Value = vbChecked Then
        TBox(33).Text = TBox(33).Tag
        TBox(33).Visible = True
        Label1(66).Visible = True
    Else
        TBox(33).Tag = TBox(33).Text
        TBox(33).Visible = False
        Label1(66).Visible = False
    End If
    
End Sub

Private Sub chkExempt_Click()

    If chkExempt.Value = True Then
        'Call EditSetting
        Call LoadDeptSettings
        EmpTab.TabVisible(5) = True
        'EmpTab.TabsPerRow = 6
    Else
        EmpTab.TabVisible(5) = False
        'EmpTab.TabsPerRow = 5
    End If

End Sub

Private Sub LoadDeptSettings()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM GeneralSettings WHERE DeptID='" & cmbDeptID.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            FirstEntry = False
            TEnter.Value = CDate(IIf(IsNull(!EnterTime), "09:00", !EnterTime))
            TExit.Value = CDate(IIf(IsNull(!ExitTime), "17:00", !ExitTime))
            LunchStart.Value = CDate(IIf(IsNull(!LunchInTime), "12:30", !LunchInTime))
            LunchEnd.Value = CDate(IIf(IsNull(!LunchOutTime), "13:30", !LunchOutTime))
            DTLunch.Value = CDate(IIf(IsNull(!LunchTime), "01:00", MinsToTime(Val(!LunchTime & "") * 60)))
            
            FridayBreakStart.Value = CDate(IIf(IsNull(!FridayBreakStartTime), "12:15", !FridayBreakStartTime))
            FridayBreakEnd.Value = CDate(IIf(IsNull(!FridayBreakEndTime), "14:00", !FridayBreakEndTime))
            FridayCheckOut.Value = CDate(IIf(IsNull(!FridayCheckOut), "18:00", !FridayCheckOut))
            
            Stxts(0) = ![AbsAllowed] & ""
            Stxts(1) = ![TotalAbsAllowed] & ""
            
            dtLHours = CDate(IIf(IsNull(!TimeAllowed), "09:00", MinsToTime(!TimeAllowed * 60)))
            Stxts(2) = ![NTimes] & ""
            Stxts(3) = ![OverTime] & ""
            
            Stxts(4) = ![DinnerAmt] & ""
            Stxts(5) = ![ShortTermPer] & ""
            Stxts(6) = ![AdvSalPer] & ""
            
            Stxts(7) = ![longtermtime] & ""
            Stxts(8) = ![LongTermLimit] & ""
            Stxts(10) = Val(!SocialAmt & "")
            
            chkAbsents = Abs(![DeductAbsents])
            chkLateHrs = Abs(![DeductLHrs])
            
            optGeneral(0).Value = IIf((!BGHoliday = True), True, False)
            optGeneral(1).Value = IIf((!BGHoliday = False), True, False)
            optGeneral(2).Value = IIf((!AGHoliday = True), True, False)
            optGeneral(3).Value = IIf((!AGHoliday = False), True, False)
            
            optPublic(0).Value = IIf((!BPHoliday = True), True, False)
            optPublic(1).Value = IIf((!BPHoliday = False), True, False)
            optPublic(2).Value = IIf((!APHoliday = True), True, False)
            optPublic(3).Value = IIf((!APHoliday = False), True, False)
            Stxts(9) = CalcNet
            'chkMultiple = IIf(!AllowMultiple & "" = "", 0, Abs(!AllowMultiple))
            If Val(!EmpAttCat_RefID & "") <> 0 Then
                cmbCatagory.Text = Val(!EmpAttCat_RefID & "")
            End If
        Else
            FirstEntry = True
        End If
        .Close
    End With
    
    Set rs = Nothing
End Sub

Private Sub chkGroupofItems_Click()
    If chkGroupofItems.Value = vbChecked Then
        'cmbDTItem.ID = ""
        cmbDTItem.Enabled = False
    Else
        cmbDTItem.Enabled = True
    End If
End Sub

Private Sub chkOldAge_Click()
    
    If chkOldAge.Value Then
        TBox(2).Visible = True
        'TBox(25).Visible = True
        Label1(57).Visible = True
        lblOldSal.Visible = True
    Else
        TBox(2).Visible = False
        'TBox(25).Visible = False
        Label1(57).Visible = False
        lblOldSal.Visible = False
        
        TBox(2) = ""
        TBox(25) = ""
    End If
    
End Sub

Private Sub chkOTByQTy_Click()
    Label1(64).Visible = chkOTByQTy.Value
    TBox(32).Visible = chkOTByQTy.Value
End Sub

Private Sub chkPerPiece_Click()
    If chkPerPiece.Value Then
        Label1(24).Caption = "Rate"
    Else
        Label1(24).Caption = "Salary"
    End If
End Sub

Private Sub chkSocial_Click()
    'If Not Loaded Then Exit Sub
    If chkSocial.Value Then
        TBox(24).Visible = True
        Label1(58).Visible = True
        lblSal.Visible = True
        'TBox(24).SetFocus
    Else
        TBox(24).Visible = False
        'TBox(22).Visible = False
        Label1(58).Visible = False
        lblSal.Visible = False
        TBox(24) = ""
        'TBox(22) = ""
    End If
    
End Sub

Private Sub chkSpecial_Click()
    'If Not Loaded Then Exit Sub
    'If Tbox(40) = "" Then chkSpecial.value = False
    'If chkSpecial Then
    '    If SqlSettings = "" Then
            
    '        Load frmSettings
            'frmSettings.Tbox(40) = Tbox(40)
   '         frmSettings.IsNew = True
            
  '          Dim rs As New ADODB.Recordset
  '          rs.Open "Select ID from Generalsettings Where ID='" & Tbox(40) & "'", con, adOpenForwardOnly, adLockReadOnly
  '          If rs.RecordCount > 0 Then
  '              frmSettings.Add = False
  '              Call frmSettings.Editrec
  '          Else
   '             frmSettings.Add = True
    '        End If
   '         rs.Close
    '        Set rs = Nothing
            
     '       frmSettings.Show 1
     '   Else
     '       frmSettings.Show 1
     '   End If
    'End If
End Sub

Private Sub chkTaxable_Click()
    'If Not Loaded Then Exit Sub
    If chkTaxable.Value Then
        TBox(5).Visible = True
        Label1(56).Visible = True
    Else
        TBox(5).Visible = False
        TBox(5) = ""
        Label1(56).Visible = True
    End If
End Sub

Private Sub cmbCatagory_Change()
    Call cmbCatagory_Click
End Sub

Private Sub cmbCatagory_Click()
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    LVCat.ListItems.Clear
    With rs
        .Open "Select * From EmployeeAttendanceCatagories Where EntryID=" & cmbCatagory.Text, con, adOpenForwardOnly, adLockReadOnly
        
        Set ITM = LVCat.ListItems.add(, , "Lunch Scanning")
        ITM.SubItems(1) = IIf(!LunchTimeScanning, "Required", "Not Required")
        
        Set ITM = LVCat.ListItems.add(, , "Time Deduction")
        ITM.SubItems(1) = IIf(!LateDeduction, "Yes", "No")
         
        Set ITM = LVCat.ListItems.add(, , "Time Deduction After")
        ITM.SubItems(1) = !LateDeductionAfter & ""
        
        Set ITM = LVCat.ListItems.add(, , "Gate Pass Required")
        ITM.SubItems(1) = IIf(!GatePassRequired, "Yes", "No")
        
        Set ITM = LVCat.ListItems.add(, , "Leaves Recorded")
        ITM.SubItems(1) = IIf(!LeavesRecorded, "Yes", "No")
        
        Set ITM = LVCat.ListItems.add(, , "Over Time Paid")
        ITM.SubItems(1) = IIf(!OverTimePaid, "Yes", "No")
        
        Set ITM = LVCat.ListItems.add(, , "Sandwich Leaves Deduction")
        ITM.ListSubItems.add(, , IIf(!DeductionSpecialAbsents, "Yes", "No")).Tag = Abs(!DeductionSpecialAbsents)
        
        Set ITM = LVCat.ListItems.add(, , "Over Time Rate")
        ITM.ListSubItems.add(, , Val(!OTRate & "")).Tag = Val(!OTRate & "")
        
        Set ITM = LVCat.ListItems.add(, , "Deduction Percentage")
        ITM.ListSubItems.add(, , Val(!DeductionPercentage & "")).Tag = Val(!DeductionPercentage & "")
        
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbdept_DropButtonClick()
    List1.Visible = False
End Sub

Private Sub cmbdept_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode = vbKeyDown Then
       If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
       KeyCode = 0
       dontShowList = True
    ElseIf KeyCode = vbKeyUp Then
       If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
       KeyCode = 0
       dontShowList = True
    ElseIf KeyCode = 13 And List1.Visible Then
       cmbDept = List1
       List1.Visible = False
       cmbDept.SetFocus
    End If

End Sub

Private Sub cmbdept_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If dontShowList Then dontShowList = False Else ShowList List1, cmbDept, cmbDept ', , fCat.Top + 110
End Sub

Private Sub cmbdept_LostFocus()
    List1.Visible = False
End Sub


Private Sub cmbdeptID_Validate(cancel As Boolean)
    If Not cmbDeptID.MatchFound Then cmbDeptID.ListIndex = cmbDept.ListIndex
End Sub
Private Sub cmbdept_Change()
    If cmbDept.MatchFound Then Call cmbdept_Click 'Else If cmbdept <> "" Then cmbdept.DropDown
End Sub
Private Sub cmbdept_Click()
    If cmbDeptID.ListIndex <> cmbDept.ListIndex Then cmbDeptID.ListIndex = cmbDept.ListIndex
End Sub
Private Sub cmbdept_Validate(cancel As Boolean)
   If Not cmbDept.MatchFound Then cmbDept.ListIndex = cmbDeptID.ListIndex
End Sub
Private Sub cmbsaleAcc_Validate(cancel As Boolean)
    If Not cmbSaleAcc.MatchFound Then cmbSaleAcc = ""
End Sub

Private Sub cmbPchAcc_Validate(cancel As Boolean)
    If Not cmbPchAcc.MatchFound Then cmbPchAcc = ""
End Sub

Private Sub cmbDeptID_Change()
    If cmbDeptID.MatchFound Then Call cmbDeptID_Click Else If Not cmbDeptID = "" Then cmbDeptID.DropDown
End Sub

Private Sub cmbDeptID_Click()

    If cmbDept.ListIndex <> cmbDeptID.ListIndex Then cmbDept.ListIndex = cmbDeptID.ListIndex
    
    If add Then
        If TableName = "Employees" Then
            TBox(40) = GetNextEmpID(cmbDeptID, bTempEmp)
        Else
    '        Tbox(40) = GetNextContID(cmbDeptID)
        End If
    End If
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select TempDept From Departments Where DeptID='" & cmbDeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        If !TempDept Then
            chkPerPiece.Visible = True
            Label1(67).Visible = True
            cmbMainContractor.Visible = True
        Else
            chkPerPiece.Value = vbUnchecked
            chkPerPiece.Visible = False
            Label1(67).Visible = False
            cmbMainContractor.Visible = False
            cmbMainContractor.Text = ""
        End If
        .Close
    End With
    Set rs = Nothing
    
End Sub



Private Sub cmbEmpType_Click()
    cmbProcess.Enabled = CBool(cmbEmpType.ListIndex = 2)
    If cmbEmpType.ListIndex = 2 Then
        cmbStatus.ListIndex = 1
    Else
        cmbStatus.ListIndex = 0
    End If
End Sub

Private Sub cmbGuarantors_matched()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM GuarantorList_Imported WHERE intGuarantorID=" & cmbGuarantors.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            TBox(47) = !strFullName & ""
            TBox(46) = !strFatherName & ""
            TBox(43) = !strNICNo & ""
            TBox(42) = !strCast & ""
            TBox(44) = !strPhoneNo & ""
            TBox(45) = !strCellNo & ""
            TBox(41) = !strKnownBy & ""
            TBox(49) = !strAddress & ""
            TBox(48) = !strComments & ""
        End If
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmbMaritalStatus_Click()
TBox(11) = cmbMaritalStatus
End Sub


Private Sub cmdAddBro_Click()

    Dim f As New frmAddBroSis
    Load f
    Dim strName As String, Relation As String, NICNo As String, Birth As Date, strAge As String
    Dim MaritalStatus As String, Education As String, Occupation As String
    
    Dim myVal As Integer
    myVal = f.ShowMe(strName, Relation, NICNo, MaritalStatus, Birth, strAge, Education, Occupation)
    If myVal = 1 Or myVal = 2 Then
        'Code To Add It In The List...
        Dim ITM As ListItem
        
        Set ITM = LVBro.ListItems.add(, , strName)
        ITM.SubItems(1) = Relation
        ITM.SubItems(2) = NICNo
        ITM.SubItems(3) = Birth
        ITM.SubItems(4) = strAge
        ITM.SubItems(5) = MaritalStatus
        ITM.SubItems(6) = Education
        ITM.SubItems(7) = Occupation
    End If
    
    If myVal = 2 Then Call cmdAddBro_Click
    
End Sub

Private Sub cmdAddChildren_Click()
    
  
'    Dim strName As String, Relation As String, NICNo As String, Birth As Date, strAge As String
'    Dim MaritalStatus As String, Education As String, ClassName As String, Institute As String, Occupation As String
'    Dim Dependent As Boolean
'    Dim myVal As Integer
    LvChildIndex = 0
    Call ShowChildForm
    
End Sub
Private Sub ShowChildForm()

    Dim f As New frmAddChild
    Dim LV As ListView
    Load f
    Dim LVR As ListView
    Dim strName As String, Relation As String, NICNo As String, MaritalStatus As String
    Dim Birth As Date, strAge As String, Education As String, Occupation As String, Dependent As Boolean, ClassName As String, Institute As String
    myVal = f.ShowMe(strName, Relation, NICNo, MaritalStatus, Birth, strAge, Education, Occupation, Dependent, ClassName, Institute)
    
    If Relation = "Daughter" Then
        If MaritalStatus = "Single" Then
            Set LV = LVChildren
        Else
            Set LV = LVBro
        End If
    ElseIf Relation = "Son" Then
        If DateDiffInYears(Birth) < 18 Then
            Set LV = LVChildren
            Set LVR = LVBro
        Else
            Set LV = LVBro
            Set LVR = LVChildren
        End If
    ElseIf Relation = "Wife" Or Relation = "Husband" Then
        Set LV = LVChildren
    Else
        If Dependent Then
            Set LV = LVChildren
        Else
            Set LV = LVBro
        End If
    End If
            
    If myVal = 1 Or myVal = 2 Then
        'Code To Add It In The List...
        Dim ITM As ListItem
        If LvChildIndex <> 0 Then
            LVR.ListItems.Remove (LvChildIndex)
        End If
        
        Set ITM = LV.ListItems.add(, , strName)
        ITM.SubItems(1) = Relation
        ITM.SubItems(2) = NICNo
        ITM.SubItems(3) = Format(Birth, "dd-MMM-yyyy")
        ITM.SubItems(4) = strAge
        ITM.SubItems(5) = MaritalStatus
        ITM.SubItems(6) = Education
        ITM.SubItems(7) = Occupation
        ITM.SubItems(8) = ClassName
        ITM.SubItems(9) = Institute

    End If
    
    If myVal = 2 Then Call cmdAddChildren_Click

End Sub

Private Sub cmdAddCosting_Click()

    If cmbProcessCD.MatchFound = False Then
        MsgBox "Please Select Process.", vbInformation
        Exit Sub
    ElseIf cmbItemGroupCD.MatchFound = False Then
        MsgBox "Please Select Item Group.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To LVCD.ListItems.count
        With LVCD.ListItems(i)
            If Val(cmbProcessCD.ID) = Val(.ListSubItems(1).Tag) And Val(cmbItemGroupCD.ID) = Val(.ListSubItems(2).Tag) Then
                MsgBox "Already Added.", vbInformation
                Exit Sub
            End If
        End With
    Next
    
    Dim ITM As ListItem
    Set ITM = LVCD.ListItems.add(, , "")
    ITM.ListSubItems.add(, , cmbProcessCD.Text).Tag = cmbProcessCD.ID
    ITM.ListSubItems.add(, , cmbItemGroupCD.Text).Tag = cmbItemGroupCD.ID
    
End Sub

Private Sub cmdAddHistory_Click()

    On Error GoTo err
    
    Dim ITM As ListItem
    Set ITM = LVHistory.ListItems.add(, , Format(DTHistory(0), "dd-MMM-yyyy"))
    
    ITM.SubItems(1) = Format(DTHistory(1), "dd-MMM-yyyy")
    ITM.SubItems(2) = txtHistory(0)
    ITM.SubItems(3) = txtHistory(1)
    ITM.SubItems(4) = txtHistory(2)
    ITM.SubItems(5) = txtHistory(3)
         
    txtHistory(0) = ""
    txtHistory(1) = ""
    txtHistory(2) = ""
    txtHistory(3) = ""
    DTHistory(0).SetFocus
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdAddMC_Click()

    Dim strMedicalTest As String, strLab As String, myDT As Date, strStatus As String, strTreatment As String, strRemarks As String
    Dim f As New frmAddMedicalCheckup
    If f.ShowMe(strMedicalTest, strLab, myDT, strStatus, strTreatment, strRemarks) = False Then Exit Sub
   
    Set ITM = LVMC.ListItems.add(, , LVMC.ListItems.count + 1)
    ITM.ListSubItems.add , , strMedicalTest
    ITM.ListSubItems.add , , strLab
    ITM.ListSubItems.add , , Format(myDT, "dd-MMM-yy")
    ITM.ListSubItems.add , , strStatus
    ITM.ListSubItems.add , , strTreatment
    ITM.ListSubItems.add , , strRemarks
    
End Sub


Private Sub cmdAddReview_Click()

    Dim DT As Date, strReviewBy As String, strReview As String, strType As String, strRating As String
    Dim f As New frmAddReview
    
    Load f
    f.ShowMe DT, strReviewBy, strReview, strType, strRating
    
    Dim ITM As ListItem
    Set ITM = LVReview.ListItems.add(, , LVReview.ListItems.count + 1)
    
    ITM.SubItems(1) = Format(DT, "dd-MMM-yyyy")
    ITM.SubItems(2) = strType
    ITM.SubItems(3) = strRating
    ITM.SubItems(4) = strReviewBy
    ITM.SubItems(5) = strReview
    
End Sub

Private Sub cmdAddSuccession_Click()

    On Error GoTo err
    
    If cmbSuccessionEmpID.MatchFound = False Then Exit Sub
    Dim ITM As ListItem
    Set ITM = LVSuccessions.ListItems.add(, cmbSuccessionEmpID.ID, LVSuccessions.ListItems.count + 1)
    
    ITM.SubItems(1) = cmbSuccessionEmpID.Text
    ITM.SubItems(2) = txtSuccession(0).Text
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdAddWarning_Click()

    Dim DT As Date, strWarnedBy As String, strWarning As String
    Dim f As New frmAddNotice
    
    Load f
    f.ShowMe DT, strWarnedBy, strWarning
    
    Dim ITM As ListItem
    Set ITM = LVNotices.ListItems.add(, , LVNotices.ListItems.count + 1)
    
    ITM.SubItems(1) = Format(DT, "dd-MMM-yyyy")
    ITM.SubItems(2) = strWarnedBy
    ITM.SubItems(3) = strWarning

End Sub


Private Sub cmdDeleteAll_Click()

    Dim i As Integer
    For i = 1 To LVDTItems.ListItems.count
        If Val(LVDTItems.ListItems(i).Tag) > 0 Then
            strDelListEmpDT = strDelListEmpDT & Val(LVDTItems.ListItems(i).Tag) & ","
        End If
        'LVDTItems.ListItems.Remove LVDTItems.SelectedItem.Index
    Next
    LVDTItems.ListItems.Clear
    
End Sub

Private Sub cmdImport_Click()

    Dim f As New frmGetSingleComboValue
    Load f
    f.cmbCombo.AddVals con, "'{' + EmpID + '} ' + Name + ' {' + Designation '}'", "Employees", "EmpID"
    Dim strEmpID As String
    If f.getComboValue(strEmpID, "Employee", "Employee") = False Then Exit Sub
    
    Call LoadDailyTargets(strEmpID, True)

End Sub

Private Sub EmpTab_Click(PreviousTab As Integer)
    If EmpTab.Tab = 16 Then
        Me.KeyPreview = False
    Else
        Me.KeyPreview = True
    End If
End Sub

Private Sub LVBro_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVBro.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to delete ?", vbQuestion + vbYesNo) = vbYes Then
            LVBro.ListItems.Remove LVBro.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LVCD_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        If LVCD.ListItems.count = 0 Then Exit Sub
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LVCD.SelectedItem.Tag) > 0 Then
                strDelListEmpCD = strDelListEmpCD & Val(LVCD.SelectedItem.Tag) & ","
            End If
            LVCD.ListItems.Remove LVCD.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LVChildren_DblClick()

    With LVChildren.SelectedItem
        
        strName = .Text
        Relation = .ListSubItems(1)
        NICNo = .ListSubItems(2)
        Birth = .ListSubItems(3)
        strAge = .ListSubItems(4)
        MaritalStatus = .ListSubItems(5) & ""
        Education = .ListSubItems(6)
        Occupation = .ListSubItems(7)
        ClassName = .ListSubItems(8)
        Institute = .ListSubItems(9)
        
    End With
   LvChildIndex = LVChildren.SelectedItem.Index
    
    Call ShowChildForm


End Sub

Private Sub LVChildren_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVChildren.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to delete ?", vbQuestion + vbYesNo) = vbYes Then
            LVChildren.ListItems.Remove LVChildren.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LVDTItems_DblClick()
    If LVDTItems.ListItems.count = 0 Then Exit Sub
    With txtDTQtyEdit
        .Move LVDTItems.Left + LVDTItems.ColumnHeaders(3).Left, LVDTItems.Top + LVDTItems.SelectedItem.Top, LVDTItems.ColumnHeaders(3).Width
        .Text = Val(LVDTItems.SelectedItem.SubItems(2))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub LVDTItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVDTItems.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Remove?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LVDTItems.SelectedItem.Tag) > 0 Then
                strDelListEmpDT = strDelListEmpDT & Val(LVDTItems.SelectedItem.Tag) & ","
            End If
            LVDTItems.ListItems.Remove LVDTItems.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LVHistory_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVHistory.ListItems.count = 0 Then Exit Sub
    If KeyCode <> 46 Then Exit Sub
    If MsgBox("Are you sure ?", vbQuestion + vbYesNo) = vbYes Then
        LVHistory.ListItems.Remove LVHistory.SelectedItem.Index
    End If
End Sub

Private Sub cmdAddTraining_Click(Index As Integer)
    
    
    Exit Sub
    Dim strName As String, strAddress As String, strKnownBy As String, strPhoneNo As String
    Dim f As Form, ITM As ListItem, myLV As ListView
    
    Set f = New frmAddGuarantor
    
    If f.ShowMe(strName, strAddress, strKnownBy, strPhoneNo) = False Then Exit Sub
    
    Set myLV = LVGuarantors
    Set ITM = myLV.ListItems.add(, , myLV.ListItems.count + 1)
    ITM.ListSubItems.add , , strName
    ITM.ListSubItems.add , , strAddress
    ITM.ListSubItems.add , , strPhoneNo
    ITM.ListSubItems.add , , strKnownBy
    
End Sub

Private Sub cmdCancel_Click()

    If DeleteJob = False Then
        con.Execute "DELETE FROM EmpJobs WHERE EmpID='" & TBox(40) & "'"
    End If
    
    Unload Me
    If SqlSettings <> "" Then
        Unload frmSettings
    End If
    
    
End Sub

Private Sub cmdCopy_Click()
    TBox(23) = TBox(8).Text & ""
End Sub

Private Sub cmdNext_Click()
    If saveEmp Then
        strTemp = cmbDept.Text
        Unload Me
        Me.SqlSettings = ""
        Me.add = True
        Me.edit = False
        Me.cmbDept = strTemp
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If saveEmp Then
        Unload Me
        If SqlSettings <> "" Then Unload frmSettings
    End If
End Sub


Private Sub cmdPost_Click()
Dim ITM As ListItem
With LVQ
    Set ITM = .ListItems.add(, , txts(0))
    ITM.ListSubItems.add , , txts(1)
    ITM.ListSubItems.add , , txts(2)
    ITM.ListSubItems.add , , DT.year
    ITM.ListSubItems.add , , txts(3)
    ITM.ListSubItems.add , , txts(4)
End With

For i = 0 To txts.count - 1
    txts(i) = ""
Next

NextQual

End Sub

Private Sub RefreshQual(EmpId As String)
Dim Rec As New ADODB.Recordset, ITM As ListItem
With Rec
    .CursorLocation = adUseClient
    .Open "SELECT * FROM EmpQualification WHERE EmpID='" & EmpId & "'", con, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            Set ITM = LVQ.ListItems.add(, , !SrNo)
            ITM.ListSubItems.add , , !Diploma & ""
            ITM.ListSubItems.add , , !Institute & ""
            ITM.ListSubItems.add , , ![year] & ""
            ITM.ListSubItems.add , , !Grad & ""
            ITM.ListSubItems.add , , !Division & ""
            .MoveNEXT
        Next
    End If
End With
End Sub

Private Sub SaveQual(EmpId As String)

con.Execute "DELETE FROM EmpQualification WHERE EmpID='" & EmpId & "'"

For i = 1 To LVQ.ListItems.count
    With LVQ.ListItems(i)
        con.Execute "INSERT INTO EmpQualification VALUES('" & EmpId _
            & "'," & .Text & ",'" & .SubItems(1) & "','" & .SubItems(2) & "'," & .SubItems(3) _
            & ",'" & .SubItems(4) & "','" & .SubItems(5) & "')"
    End With
Next

End Sub

Private Sub DOB_Change()
    Dim Months As Double
    Dim Years As Long
    Dim TotDays As Long
    TotDays = DateDiff("d", Format(DOB.Value, "MMM-dd-yyyy"), Format(Date, "MMM-dd-yyyy"))
    Years = TotDays / 365
    'Months = Totdays Mod 365
    'Months = Round((Tot - Years) * 12)
    'TBox(3) = Round((GetServerDate - DOB.value) / (365))
    TBox(3) = Years
End Sub


Private Sub DTLunch_Change()
    Stxts(9) = CalcNet
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
'    TEnter = Date
'    TExit = Date
'    DTLunch = Date
    Me.KeyPreview = True
'    DTPicker1.value = GetServerDate
    DOB.MaxDate = Date
    Call FillCmbs
    SqlSettings = ""
    cmbReligion.AddItem "Islam"
    cmbReligion.AddItem "Chirstian"
    cmbReligion.AddItem "Other"
    
    cmbReligion.ListIndex = 0
    
    'cmbStatus.AddItem "Temporary"
    

    cmbStatus.AddItem "Salary"
    cmbStatus.AddItem "Wages"
    
    cmbMaritalStatus.AddItem "Married"
    cmbMaritalStatus.AddItem "Single"
    cmbMaritalStatus.AddItem "Divorced"
    cmbMaritalStatus.AddItem "Widowed"
    cmbMaritalStatus.AddItem "Not Specified"
    
    cmbEmpType.AddItem "Staff Member"
    cmbEmpType.AddItem "Director"
    'cmbEmpType.AddItem "Production Staff Member"
    
    cmbEmpType.ListIndex = 0
    
    JoinDate = GetServerDate
    'JoinDate = #7/1/2004#
    RetireDate = DateAdd("yyyy", 25, GetServerDate)
    EmpTab.Tab = 0
    
    cmbEmpType.ListIndex = 0
    Loaded = True
   
    Call AddToCombo(cmbProcess, "Description,ProcessID", "Processes", "Order By ProcessID")
    cmbProbationDays.ListIndex = 3
    
    EmpTab.TabVisible(5) = False
    EmpTab.TabVisible(6) = False
        
        
    bTempEmp = False
    TBox(40).Locked = True
        
    cmbGuarantors.ListHeight = 2400
    cmbGuarantors.AddVals con, "REPLACE(strFullName,'#','No.') + ' ' + REPLACE(strFatherName,'#','No.') + ' {' + CAST(intGuarantorID AS VARCHAR(5))+ '}'", "GuarantorList_Imported", "intGuarantorID"
        
    strDelListEmpDT = ""
    strDelListEmpCD = ""
    strCVID = ""
    
End Sub

Private Function saveEmp() As Boolean

    On Error GoTo err

    Dim OpBalance As Double

    If Not cmbDeptID.MatchFound Then
        MsgBox "Invalid Department Name.", vbInformation
        saveEmp = False
        Exit Function
    ElseIf Trim(TBox(0)) = "" Then
        MsgBox "Invalid Employee Name.", vbInformation
        saveEmp = False
        Exit Function
    ElseIf cmbBloodGroup.ListIndex = -1 Then
        MsgBox "Please Select Blood Group.", vbInformation
        saveEmp = False
        Exit Function
    ElseIf cmbDesignation.ListIndex = -1 Then
        MsgBox "Please Select Designation from the List.", vbInformation
        Exit Function
    End If

    If TableName = "Employees" Then
        If chkTaxable.Value Then
            If Val(TBox(5)) <= 0 Then
                MsgBox "Tax Amount Can't Be Zero.", vbInformation
                TBox(5).SetFocus
                Exit Function
            End If
        ElseIf Val(TBox(3)) < 18 Then
            MsgBox "Employee's age must be at least 18 years.", vbInformation
            DOB.SetFocus
            Exit Function
        
        ElseIf chkSocial.Value Then
'            If Trim(TBox(24)) = "" Then
'                MsgBox "Social Security Number Can't Be Null.", vbInformation
'                TBox(24).SetFocus
'                Exit Function
'            '        ElseIf Val(TBox(22)) = 0 Then
'            '            MsgBox "Social Security Salary Is Invalid.", vbInformation
'            '            TBox(22).SetFocus
'            '            Exit Function
'            End If
    
        ElseIf chkOldAge.Value Then
'            If Trim(TBox(2)) = "" Then
'                MsgBox "Old Age Number Can't Be Null.", vbInformation
'                TBox(2).SetFocus
'                Exit Function
'            End If
'
        ElseIf Val(TBox(17)) = 0 Then
            MsgBox "Must Enter Salary Amount.", vbInformation
            EmpTab.Tab = 2

            Exit Function

            'ElseIf Val(con.Execute("Select count(*) From Employees where upper(name)='" & Trim(UCase(TBox(0))) & "' and empid<>'" & Tbox(40) & "'").Fields(0) & "") > 0 Then
            '    MsgBox "An Employee With The Name """ & TBox(0) & """ Already Exists.", vbInformation
            '     Exit Function
        
        End If
    

    End If

    Dim OldAge As Byte, Social As Byte, TaxPayee As Byte, Bonus As Byte, ArmyRetired As Byte
    Dim Special As String

    'If chkOldAge.value = vbChecked Then OldAge = True Else OldAge = False
    'If chkSocial.value = vbChecked Then Social = True Else Social = False
    'If chkTaxable.value = vbChecked Then TaxPayee = True Else TaxPayee = False

    OldAge = Abs(chkOldAge.Value)
    Social = Abs(chkSocial.Value)
    TaxPayee = Abs(chkTaxable.Value)
    Bonus = Abs(chkBonus.Value)
    If IsNull(chkArmy.Value) Then
        ArmyRetired = 0
    Else
        ArmyRetired = Abs(chkArmy.Value)
    End If
    If IsNull(chkShowInExternal) Then chkShowInExternal = True
    
    Dim Exempt As Integer
    If chkExempt.Value = True Then Exempt = 1
    
    Dim ProcessID As Long
    
    If cmbProcess.ListIndex >= 0 Then
        ProcessID = cmbProcess.List(cmbProcess.ListIndex, 1)
    Else
        ProcessID = 0
    End If
    Dim strEmpID As String
    Dim dSalary As Double, dRate As Double
    
    If chkPerPiece.Value Then
        dSalary = 0
        dRate = Val(TBox(17))
    Else
        dRate = 0
        dSalary = Val(TBox(17))
    End If
    
    Call StartTrans(con)

    If add Then
        strEmpID = TBox(40)
        con.Execute "INSERT INTO " & TableName & "(DeptID,EmpID,Name,Rel,fname,dob,religion,sex,address,NIC," & _
        "Phone1,Phone2,designation,refName,IdentityMark,lastEmpFirm,lastEmpDesig,lastEmpReason,JoinDate," & _
        "Retirement,Salary,Bonus,StartingSalary,oldage,Social,taxpayee,remarks,GInfo,RefFName,RefPhone1," & _
        "RefPhone2,TempAddress," & "RefAddress,SocialNo,SocialSal,OldAgeNo,OldAgeSal,Status,EmpType,Active," & _
        "ExemptSettings,BloodGroup,ProcessID,BankPymt,ShowInExternal,EmergencyName,EmergencyFName," & _
        "EmergencyPhone1,EmergencyPhone2,EmergencyAddress,OTByQty,MinQty,UnionFund,Lunch,TaxAmt,ArmyRetired" & _
        ",EmpAccNo,PerPieceRate,Rate,MainContractorID,ProbationPeriod,FixAllowance,EOBIAmt,SocialSecurityAmt,EmpLunchAmt,BondPeriod,RelationWithCompEmp,RelativeEmpID,TempEmp,intGuarantorID,Pay_Full_Salary,Caste,Maslak) VALUES('" & cmbDeptID & "','" & _
        TBox(40) & "','" & TBox(0) & "','" & TBox(1).Tag & "','" & TBox(1) & "','" & DOB & "','" & cmbReligion & _
        "','" & cmbSex & "','" & TBox(8) & "','" & TBox(4) & "','" & TBox(6) & "','" & TBox(7) & "','" & _
        cmbDesignation.Text & "','" & TBox(10) & "','" & TBox(11) & "','" & TBox(12) & "','" & TBox(13) & "','" & _
        TBox(14) & "','" & JoinDate & "','" & RetireDate & "'," & Val(TBox(15)) & "," & Bonus & "," & _
        dSalary & "," & OldAge & "," & Social & "," & TaxPayee & ",'" & TBox(16) & "','" & Special & "','" & _
        TBox(18) & "','" & TBox(19) & "','" & TBox(20) & "','" & TBox(23) & "','" & TBox(21) & "','" & _
        TBox(24) & "'," & Val(TBox(22)) & ",'" & TBox(2) & "'," & Val(TBox(25)) & ",'" & cmbStatus & _
        "'," & cmbEmpType.ListIndex & ",1," & Exempt & ",'" & cmbBloodGroup.Text & "'," & ProcessID & "," & Abs(chkBankPymt.Value) & _
        "," & Abs(chkShowInExternal.Value) & ",'" & TBox(27) & "','" & TBox(28) & "','" & _
        TBox(29) & "','" & TBox(30) & "','" & TBox(31) & "'," & Abs(chkOTByQTy.Value) & "," & _
        Val(TBox(32)) & "," & Abs(chkUnionFund) & "," & Abs(chkLunch) & "," & Val(TBox(5)) & _
        "," & ArmyRetired & ",'" & TBox(33) & "'," & Abs(chkPerPiece.Value) & "," & dRate & ",'" & _
        cmbMainContractor.Text & "'," & Val(cmbProbationDays.Text) & "," & Val(TBox(34)) & "," & Val(TBox(35)) & "," & Val(TBox(36)) & "," & Val(TBox(37)) & "," & Val(TBox(9)) & ",'" & TBox(39) & "','" & cmbRelEmp.ID & "'," & Abs(bTempEmp) & "," & Val(cmbGuarantors.ID) & "," & chkPayFullSalary.Value & ",'" & TBox(50) & "','" & TBox(51) & "')"
        
        con.Execute "Insert Into SalaryLedger(EmpID,DT,Description," & _
        "Salary,SalaryExt) Values('" & TBox(40) & _
        "','" & JoinDate & "','Starting Salary On Joining'," & Val(TBox(17)) & "," & Val(TBox(15)) & ")"
    
'        If DeleteCV Then
'            con.Execute "DELETE FROM CVs WHERE ID='" & CVID & "'"
'        End If
        If strCVID <> "" Then
            con.Execute "UPDATE CVs SET EmpID='" & TBox(40) & "' WHERE CVID='" & strCVID & "'"
        End If
        
    Else
    
        strEmpID = Me.Tag
        con.Execute "UPDATE " & TableName & " Set Name='" & TBox(0) & "',Rel='" & TBox(1).Tag & "',fName='" & _
        TBox(1) & "',DOB='" & DOB & "',Religion='" & cmbReligion & "',Sex='" & cmbSex & "',Address='" & TBox(8) & _
        "',NIC='" & TBox(4) & "',Phone1='" & TBox(6) & "',Phone2='" & TBox(7) & "',designation='" & cmbDesignation.Text & _
        "',RefName='" & TBox(10) & "',IdentityMark='" & TBox(11) & "',lastEmpFirm='" & TBox(12) & _
        "',lastEmpDesig='" & TBox(13) & "',lastEmpReason='" & TBox(14) & "',JoinDate='" & JoinDate & "',Retirement='" & _
        RetireDate & "'," & "StartingSalary=" & dSalary & "," & "Salary=" & Val(TBox(15)) & ",OldAge=" & OldAge & ",Social=" & _
        Social & ",TaxPayee=" & TaxPayee & ",Remarks='" & TBox(16) & "',GInfo='" & Special & "',RefFName='" & TBox(18) & _
        "',RefPhone1='" & TBox(19) & "',RefPhone2='" & TBox(20) & "',TempAddress='" & TBox(23) & "',RefAddress='" & _
        TBox(21) & "',SocialNo='" & TBox(24) & "',SocialSal=" & Val(TBox(22)) & ",OldAgeNo='" & TBox(2) & "',OldAgeSal=" & _
        Val(TBox(25)) & ",Status='" & cmbStatus & "',EmpType=" & cmbEmpType.ListIndex & ",Bonus=" & Bonus & _
        ",ExemptSettings = " & Exempt & ",BloodGroup='" & cmbBloodGroup.Text & "',ProcessID=" & ProcessID & ",BankPymt=" & _
        chkBankPymt.Value & ",ShowInExternal=" & Abs(chkShowInExternal.Value) & ",EmergencyName='" & TBox(27) & _
        "',EmergencyFName='" & TBox(28) & "',EmergencyPhone1='" & TBox(29) & "',EmergencyPhone2='" & TBox(30) & _
        "',EmergencyAddress='" & TBox(31) & "',OTByQty=" & Abs(chkOTByQTy) & ",MinQty=" & Val(TBox(32)) & ",UnionFund=" & _
        Abs(chkUnionFund) & ",lunch=" & Abs(chkLunch) & ",TaxAmt=" & Val(TBox(5)) & ",ArmyRetired=" & ArmyRetired & _
        ",EmpAccNo='" & TBox(33) & "',PerPieceRate=" & Abs(chkPerPiece.Value) & ",Rate=" & dRate & ",MainContractorID='" & _
        cmbMainContractor.Text & "',FixAllowance=" & Val(TBox(34)) & ",EOBIAmt=" & Val(TBox(35)) & ",SocialSecurityAmt=" & _
        Val(TBox(36)) & ",EmpLunchAmt=" & Val(TBox(37)) & ",BondPeriod=" & Val(TBox(9)) & ",RelationWithCompEmp='" & TBox(39) & "',RelativeEmpID='" & cmbRelEmp.ID & "',ProbationPeriod=" & Val(cmbProbationDays.Text) & ",intGuarantorID=" & Val(cmbGuarantors.ID) & ",Pay_Full_Salary=" & chkPayFullSalary.Value & ",Caste='" & TBox(50) & "',Maslak='" & TBox(51) & "' WHERE EmpID='" & Me.Tag & "'"
     
    End If

    If IsNull(DTInsurable.Value) Then
        con.Execute "Update Employees Set InsurableDT=Null Where EmpID='" & strEmpID & "'"
    Else
        con.Execute "Update Employees Set InsurableDT='" & DTInsurable.Value & "' Where EmpID='" & strEmpID & "'"
    End If

    SaveQual (TBox(40))
        
        
    '------------------
    SaveGuarantors (TBox(40))
    SaveHistory (TBox(40))
    SaveAssets (TBox(40))
    SaveMedicalInfo (TBox(40))
    SaveFamily (TBox(40))
    SaveSuccessions (TBox(40))
    SaveReviews (TBox(40))
    SaveNotices (TBox(40))
    SaveDailyTargets (TBox(40))
    SaveCostingDetails (TBox(40))
    
    Call AttachCNICPDF(TBox(40))
    
    If chkExempt.Value = True Then
        If SaveSetting = False Then
            con.RollbackTrans
            Exit Function
        End If
    Else
        GetDeptSetting
    End If
 
SavePic:

    If (Image1.Picture = 0 Or Trim(PicFileName) = "") And (Image2.Picture = 0 Or Trim(ThumbFileName) = "") Then GoTo TheEnd

    Dim Rec As New ADODB.Recordset
    Dim rsPic As New ADODB.Stream
    Dim rsThumb As New ADODB.Stream
        
    Rec.Open "Select EmpID,EmpPic,ThumbExp from " & TableName & " Where EmpID='" & TBox(40) & "'", con, adOpenStatic, adLockOptimistic
    'Rec.Open "Select UserName,UserPic From UserMainSettings Where UserName='nadeem'", con, adOpenStatic, adLockOptimistic
    If Image1.Picture <> 0 And Trim(PicFileName) <> "" Then
      'Save Picture
        rsPic.Open
        rsPic.Type = adTypeBinary
        rsPic.LoadFromFile PicFileName
        Rec.Fields(1) = rsPic.Read
        rsPic.Close
    End If
   
    'Save Thumb Expression
    If Image2.Picture <> 0 And Trim(ThumbFileName) <> "" Then
        rsThumb.Open
        rsThumb.Type = adTypeBinary
        rsThumb.LoadFromFile ThumbFileName
        Rec.Fields("thumbExp") = rsThumb.Read
        rsThumb.Close
    End If
    
    Rec.Update
    Rec.Close
    Set Rec = Nothing
    Set rsPic = Nothing
    Set rsThumb = Nothing
    
    
    'If Special = "GENERAL" Then Con.Execute "delete from GeneralSettings Where ID='" & Tbox(40) & "'"
    
TheEnd:
    saveEmp = True
    con.CommitTrans
Exit Function
err:
    MsgBox err.Description
End Function


Private Sub AddEmpPic(Optional ThumbExp As Boolean)
    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        If ThumbExp Then
            Image2.Picture = LoadPicture(Picfile)
            ThumbFileName = Picfile
            Label2.Visible = False
        Else
            Image1.Picture = LoadPicture(Picfile)
            PicFileName = Picfile
            lblPic.Visible = False
        End If
        
    End If
End Sub

Private Sub Image1_DblClick()
    Call AddEmpPic
End Sub

Private Sub Image2_DblClick()
    Call Label2_DblClick
End Sub

Private Sub Label2_DblClick()
    Call AddEmpPic(True)
End Sub

Private Sub lblPic_DblClick()
    AddEmpPic
End Sub

Private Sub List1_Click()

    If UCase(Screen.ActiveControl.Name) <> "CMBDEPT" Then
        cmbDept = List1
        List1.Visible = False
        cmbDept.SetFocus
    End If

End Sub


Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LV.ListItems.count > 0 Then
        If Button = vbRightButton Then
            PopUpMenu mnuPopup
        End If
    End If

End Sub

Private Sub LVGuarantors_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVGuarantors.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVGuarantors.ListItems.Remove LVGuarantors.SelectedItem.Index
        End If
    End If
    
End Sub


Private Sub LVNotices_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVNotices.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If CurrentUserName <> "Administrator" Then
            MsgBox "You do not have Permissions to Delete Notices.", vbInformation
            Exit Sub
        End If
        If MsgBox("Are you sure to Delete ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        LVNotices.ListItems.Remove LVNotices.SelectedItem.Index
    End If
End Sub

Private Sub LVQ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If LVQ.ListItems.count > 0 Then
        If Button = vbRightButton Then
            PopUpMenu mnupopup1
        End If
    End If

End Sub

Private Sub mnuDelete_Click()

    con.Execute "DELETE FROM EmpJobs WHERE JobID=" & LV.SelectedItem.Tag & ""
    con.Execute "UPDATE EmpJobs SET SrNo=SrNo-1 WHERE EmpID='" & TBox(40) & "' AND SrNo>" & LV.SelectedItem & ""

    FillJobs

End Sub

Private Sub mnuDeleteQual_Click()
Dim SrNo As Integer
SrNo = Val(LVQ.SelectedItem)
If SrNo <> 0 Then
If LVQ.ListItems.count > 0 Then
    LVQ.ListItems.Remove (LVQ.SelectedItem.Index)
    For i = SrNo To LVQ.ListItems.count
        LVQ.ListItems(i).Text = Val(LVQ.ListItems(i)) - 1
    Next
End If
End If
    NextQual

End Sub

Private Sub SO_Click(Index As Integer)

    On Error GoTo err
    If Not Loaded Then Exit Sub
    Dim sSex As String
    If SO(0) Then
        TBox(1).Tag = "S/O"
        sSex = "Male"
        TBox(1).SetFocus
        
    ElseIf SO(1) Then
        TBox(1).Tag = "D/O"
        sSex = "Female"
        TBox(1).SetFocus
    ElseIf SO(2) Then
        TBox(1).Tag = "W/O"
        sSex = "Female"
        TBox(1).SetFocus
    Else
        TBox(1).Tag = ""
    End If
    cmbSex.Text = sSex
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub SO_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    On Error GoTo err
    If KeyAscii = 13 Then
        If SO(Index).Value = False Then
            SO(Index).Value = True
        Else
            TBox(1).SetFocus
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Stxts_GotFocus(Index As Integer)

    Stxts(Index).SelStart = 0
    Stxts(Index).SelLength = Len(Stxts(Index))

End Sub


Private Sub TBox_GotFocus(Index As Integer)

    If Not Loaded Then Exit Sub
'    Select Case Index
'        Case 10, 18, 19, 20, 21, 27, 28, 29, 30, 31
'            EmpTab.Tab = 1
'        Case 12 To 14
'            EmpTab.Tab = 1
'        Case 15, 16, 17, 9, 5, 24, 2, 22, 25, 32, 33, 34, 35, 36, 37
'            EmpTab.Tab = 2
'        Case Else
'            EmpTab.Tab = 0
'    End Select
    
    '''''''''' Will Select The Previous Text ''''''''''''''
    TBox(Index).SelStart = 0
    TBox(Index).SelLength = Len(TBox(Index))
    
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 15
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
End Sub



Private Function GetNextItemID(CatID As String) As String
    Dim NextVal As String
    NextVal = GetMax("CAST(Right(ItemID,2) as int)", "Items", "Where CatID ='" & CatID & "'")
    GetNextItemID = CatID & "-" & Format(Val(NextVal) + 1, "000")
End Function

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)
    'If Index = 4 Or Index = 5 Then
    '    TBox(Index) = Format(Val(TBox(Index)), "'0.00")
    'ElseIf Index = 2 Or Index = 3 Then
    '    TBox(Index) = Format(Val(TBox(Index)), "'0")
    'End If
End Sub


Public Sub GetCVInfo(ID As String)
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    CVID = ID
    
    With rs
        .Open "Select * From CVs Where ID='" & ID & "'", con, adOpenForwardOnly, adLockReadOnly
        TBox(0) = ![Name] & ""
        DeleteCV = True
        If UCase$(![Rel]) = "S/O" Then
            SO(0).Value = True
            TBox(1).Tag = "S/O"
        ElseIf UCase$(![Rel]) = "D/O" Then
            SO(1).Value = True
            TBox(1).Tag = "D/O"
        Else
            SO(2).Value = True
            TBox(1).Tag = "W/O"
        End If
        
        TBox(1) = ![FName] & ""
        DOB = ![DOB] & ""
        Call DOB_Change
        cmbReligion = ![Religion] & ""
        'TBox(3) = ![age] & ""
        cmbSex = ![Sex] & ""
        TBox(4) = ![NIC] & ""
        TBox(6) = ![Phone1] & ""
        TBox(7) = ![Phone2] & ""
        TBox(8) = ![Address] & ""
'        cmbDesignation.Text = ![Designation] & ""
        TBox(10) = ![RefName] & ""
        TBox(18) = ![RefFName] & ""
        TBox(19) = ![RefPhone1] & ""
        TBox(20) = ![RefPhone2] & ""
        TBox(23) = ![TempAddress] & ""
        TBox(21) = ![RefAddress] & ""
'       TBox(11) = ![IdentityMark] & ""
        TBox(12) = ![lastEmpFirm] & ""
        TBox(13) = ![LastEmpDesig] & ""
        TBox(14) = ![LastEmpReason] & ""
'       TBox(15) = ![Salary] & ""
'       TBox(17) = ![StartingSalary] & ""
        TBox(16) = ![Remarks] & ""
        cmbStatus = ![Status] & ""
        
        JoinDate = Date
        'Get The Employee Type
 '       If LCase(TableName) = "employees" Then cmbEmpType.ListIndex = Val(![EmpType] & "")
        
        .Close
    End With

    Set rs = Nothing
    
    If TBox(11) = "" Then
      cmbMaritalStatus.ListIndex = -1
    Else
      cmbMaritalStatus = TBox(11)
    End If
    
    RefreshCVQual (ID)
    
    NextQual
    Loaded = True
    Me.Show
    
    cmbDept.SetFocus
    Me.Hide
 
    
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
    Resume Next
End Sub

Private Sub RefreshCVQual(ID As String)
Dim Rec As New ADODB.Recordset, ITM As ListItem
With Rec
    .CursorLocation = adUseClient
    .Open "SELECT * FROM CVQualification WHERE ID='" & ID & "'", con, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            Set ITM = LVQ.ListItems.add(, , !SrNo)
            ITM.ListSubItems.add , , !Diploma & ""
            ITM.ListSubItems.add , , !Institute & ""
            ITM.ListSubItems.add , , ![year] & ""
            ITM.ListSubItems.add , , !Grad & ""
            ITM.ListSubItems.add , , !Division & ""
            .MoveNEXT
        Next
    End If
End With
End Sub



Public Sub Editrec()

    On Error GoTo err
    DeleteJob = True
    'cmbProbationDays.Locked = True
    
    Loaded = False
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM " & TableName & " WHERE EmpID='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        TBox(40) = ![EmpId] & ""
        cmbDeptID = ![DeptID] & ""
        TBox(0) = ![Name] & ""
        FillJobs
        
        If UCase$(![Rel]) = "S/O" Then
            SO(0).Value = True
            TBox(1).Tag = "S/O"
        ElseIf UCase$(![Rel]) = "D/O" Then
            SO(1).Value = True
            TBox(1).Tag = "D/O"
        Else
            SO(2).Value = True
            TBox(1).Tag = "W/O"
        End If
        
        TBox(1) = ![FName] & ""
        If IsNull(!DOB) Then
            MsgBox "Date of Birth not Feeded, Please update DOB.", vbInformation
        Else
            DOB = ![DOB] & ""
            Call DOB_Change
        End If
        cmbReligion = ![Religion] & ""
        'TBox(3) = ![age] & ""
        cmbSex = ![Sex] & ""
        TBox(4) = ![NIC] & ""
        TBox(5) = Val(![TaxAmt] & "")
        TBox(6) = ![Phone1] & ""
        TBox(7) = ![Phone2] & ""
        TBox(8) = ![Address] & ""
        
        For i = 0 To cmbDesignation.ListCount - 1
            If cmbDesignation.List(i) = ![Designation] & "" Then
                cmbDesignation.ListIndex = i
                Exit For
            End If
        Next
        
        TBox(10) = ![RefName] & ""
        TBox(18) = ![RefFName] & ""
        TBox(19) = ![RefPhone1] & ""
        TBox(20) = ![RefPhone2] & ""
        TBox(23) = ![TempAddress] & ""
        TBox(21) = ![RefAddress] & ""
        TBox(11) = ![IdentityMark] & ""
        TBox(12) = ![lastEmpFirm] & ""
        TBox(13) = ![LastEmpDesig] & ""
        TBox(14) = ![LastEmpReason] & ""
        TBox(15) = ![Salary] & ""
        cmbProbationDays = Val(!ProbationPeriod & "")
        If !PerPieceRate & "" = True Then
            TBox(17) = Val(![Rate] & "")
        Else
            TBox(17) = Val(![StartingSalary] & "")
        End If
        'TBox(17).Locked = True
        
        TBox(16) = ![Remarks] & ""
        'TBox(26) = ![BloodGroup] & ""
        cmbBloodGroup = !BloodGroup & ""
        
        TBox(27) = ![EmergencyName] & ""
        TBox(28) = ![EmergencyFName] & ""
        TBox(29) = ![EmergencyPhone1] & ""
        TBox(30) = ![EmergencyPhone2] & ""
        TBox(31) = ![EmergencyAddress] & ""
        
        TBox(34) = Val(![FixAllowance] & "")
        
        TBox(35) = Val(!EOBIAmt & "")
        TBox(36) = Val(!SocialSecurityAmt & "")
        TBox(37) = Val(!EmpLunchAmt & "")
        
        TBox(9) = Val(!BondPeriod & "")
        TBox(39) = !RelationWithCompEmp & ""
        cmbRelEmp.ID = !RelativeEmpID & ""
        
        If IsNull(!InsurableDT) Then
            DTInsurable.Value = Date
            DTInsurable.Value = ""
        Else
            DTInsurable.Value = !InsurableDT
        End If
        'DTInsurable.value = IIf(IsNull(!InsurableDT), vbNull, !InsurableDT)
        
        
        cmbStatus = ![Status] & ""
        chkExempt.Value = ![ExemptSettings]
        chkExempt_Click
        
        If ![ExemptSettings] Then
            EditSetting
        End If
        
        TBox(24) = ![SocialNo] & ""
        TBox(22) = ![SocialSal] & ""
        
        TBox(2) = ![OldAgeNo] & ""
        TBox(25) = ![OldAgeSal] & ""
        
        JoinDate = ![JoinDate] & ""
        RetireDate = ![Retirement] & ""
        
        chkArmy.Value = !ArmyRetired & ""
        chkOldAge.Value = ![OldAge] & ""
        chkSocial.Value = ![Social] & ""
        chkTaxable.Value = ![TaxPayee] & ""
        chkBonus.Value = ![Bonus] & ""
        chkBankPymt.Value = IIf(IsNull(![bankpymt]), 0, Abs(![bankpymt]))
        TBox(33).Text = !EmpAccNo & ""
        chkShowInExternal = ![ShowinExternal] & ""
        chkOTByQTy = !OTByQty & ""
        
        chkUnionFund.Value = IIf(IsNull(![UnionFund]), 0, Abs(![UnionFund]))
        chkLunch.Value = IIf(IsNull(![Lunch]), 0, Abs(![Lunch]))
        chkPerPiece.Value = IIf(IsNull(![PerPieceRate]), 0, Abs(![PerPieceRate]))
        
        cmbMainContractor = !MainContractorID & ""
        'Get The Employee Type
        If LCase(TableName) = "employees" Then cmbEmpType.ListIndex = Val(![Emptype] & "")
        
        If Not IsNull(![EmpPic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "EmpPic"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
        
        If Not IsNull(![ThumbExp]) Then
            Set Image2.DataSource = rs
            Image2.DataField = "ThumbExp"
            Label2.Visible = False
        Else
            Label2.Visible = True
        End If
        
        cmbGuarantors.ID = Val(!intGuarantorId & "")
                
        chkPayFullSalary.Value = Abs(!Pay_Full_Salary)
        
        txtFilePath = !CNIC_PDF_FileName & ""
        txtFilePath.Tag = !CNIC_PDF_FileName & ""
                
        TBox(50) = !Caste & ""
        TBox(51) = !Maslak & ""
        
        .Close
    End With

    Set rs = Nothing

    TBox(15).Enabled = False
    TBox(17).Enabled = True
    'TBox(17).Locked = True
    If TBox(11) = "" Then
        cmbMaritalStatus.ListIndex = -1
    Else
        cmbMaritalStatus = TBox(11)
    End If
    
    RefreshQual (TBox(40))
    NextQual
    
    '------------
    LoadGuarantors (TBox(40))
    AddHistory (TBox(40))
    LoadAssets (TBox(40))
    LoadMedicalInfo (TBox(40))
    LoadFamily (TBox(40))
    LoadSuccessions (TBox(40))
    LoadReviews (TBox(40))
    LoadSalaryHistory (TBox(40))
    LoadNotices (TBox(40))
    LoadDailyTargets (TBox(40))
    LoadCostingDetails (TBox(40))
    
    fCat.Enabled = False
    Loaded = True
    Me.Show
    TBox(0).SetFocus
    Me.Hide
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
    Resume Next
    
End Sub
Private Sub NextQual()

    If LVQ.ListItems.count > 0 Then
        txts(0) = Val(LVQ.ListItems(LVQ.ListItems.count).Text) + 1
    Else
        txts(0) = "1"
    End If

End Sub


Private Sub TEnter_Change()

Stxts(9) = CalcNet

End Sub

Private Sub TExit_Change()
Stxts(9) = CalcNet
End Sub


Private Sub txtAsset_Validate(cancel As Boolean)
    txtAsset = StrConv(txtAsset, vbProperCase)
End Sub

Private Sub txtAssetRemarks_GotFocus()
    Me.KeyPreview = False
End Sub

Private Sub txtAssetRemarks_KeyPress(KeyAscii As Integer)

    Dim ITM As ListItem
    If KeyAscii = 13 Then
        
        Set ITM = LVAssets.ListItems.add(, , txtAssetSNo)
        ITM.SubItems(1) = txtAsset
        ITM.SubItems(2) = Format(DTAsset, "dd-MMM-yyyy")
        ITM.SubItems(3) = txtAssetRemarks
        txtAssetSNo = Val(txtAssetSNo) + 1
        txtAsset = ""
        txtAssetRemarks = ""
        
    End If
    
End Sub

Private Sub txtAssetRemarks_LostFocus()
    Me.KeyPreview = True
End Sub

Private Sub txtAssetRemarks_Validate(cancel As Boolean)
    txtAssetRemarks = StrConv(txtAssetRemarks, vbProperCase)
End Sub

Private Sub LVAssets_Click()
    PicEditAssets.Visible = False
End Sub

Private Sub LVAssets_DblClick()

    If LVAssets.ListItems.count = 0 Then Exit Sub
    
    PicEditAssets.Move LVAssets.Left + LVAssets.ColumnHeaders(2).Left, LVAssets.Top + LVAssets.SelectedItem.Top
    PicEditAssets.Visible = True
    
    txtEditAssets(0) = LVAssets.SelectedItem.SubItems(1)
    txtEditAssets(1) = LVAssets.SelectedItem.SubItems(2)
    txtEditAssets(2) = LVAssets.SelectedItem.SubItems(3)
    
    With txtEditAssets(0)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LVAssets_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVAssets.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to delete this Asset ?", vbQuestion + vbYesNo) = vbYes Then
            LVAssets.ListItems.Remove LVAssets.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub txtDTQty_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
        
        If cmbDTItem.MatchFound = False And chkGroupofItems.Value = vbUnchecked Then
            MsgBox "Please Select Item.", vbInformation
            Exit Sub
        End If
        If cmbDTProc.MatchFound = False Then
            MsgBox "Please Select Process.", vbInformation
            Exit Sub
        End If
        If Val(txtDTQty) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
        Dim strItemID As String, strItemName As String, lGroupID As Long
        If chkGroupofItems.Value = vbChecked Then
            strItemID = "Group"
            strItemName = "Group of Items"
            lGroupID = "0"
        Else
            strItemID = "" 'cmbDTItem.ID
            strItemName = cmbDTItem.Text
            lGroupID = cmbDTItem.ID
        End If
        
        Dim i As Integer
        For i = 1 To LVDTItems.ListItems.count
            If lGroupID = Val(LVDTItems.ListItems(i).ListSubItems(1).Tag) And Val(cmbDTProc.ID) = Val(LVDTItems.ListItems(i).ListSubItems(2).Tag) Then
                MsgBox "Already Added.", vbInformation
                Exit Sub
            End If
        Next
        
        
        Dim ITM As ListItem
        Set ITM = LVDTItems.ListItems.add(, , cmbDTProc.Text)
        ITM.Tag = "0"
        ITM.ListSubItems.add(, , strItemName).Tag = lGroupID
        ITM.ListSubItems.add(, , Val(txtDTQty)).Tag = cmbDTProc.ID
        ITM.ListSubItems.add(, , IIf(chkDTOverTime.Value = vbChecked, "Yes", "No")).Tag = chkDTOverTime.Value
        
    End If
End Sub

Private Sub txtDTQtyEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LVDTItems.SelectedItem.SubItems(2) = Val(txtDTQtyEdit)
        txtDTQtyEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtDTQtyEdit.Visible = False
    End If
End Sub

Private Sub txtDTQtyEdit_LostFocus()
    txtDTQtyEdit.Visible = False
End Sub

Private Sub txtDuty_GotFocus()
    Me.KeyPreview = False
End Sub

Private Sub txtDuty_LostFocus()
    Me.KeyPreview = True
End Sub

Private Sub txtDuty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    
    If TBox(40) = "" Then
        MsgBox "Please First Fill The Particulers Of The Employee.", vbInformation
        Exit Sub
    End If
    
    If txtDuty = "" Then
        MsgBox "Job Description Could Not Be Null.", vbInformation
        Exit Sub
    End If
    
    If CheckPrevSrNo = True Then
        PrevSrNo.lblMessage = "Sr No. | " & txtSrNo & " | Already Exists, Select The Option That Want."
        PrevSrNo.Show 1
        
        Select Case PrevSrNo.SelOption
            Case 1
            
                con.Execute "UPDATE EmpJobs SET SrNo=SrNo+1 WHERE SrNo>=" & txtSrNo & " AND EmpID='" & TBox(40) & "'"
                con.Execute "INSERT INTO EmpJobs(EmpID,SrNo,JobDescription) VALUES('" & TBox(40) & "'," & txtSrNo & ",'" & txtDuty & "')"
                
            Case 2
                con.Execute "UPDATE EmpJobs SET SrNo=SrNo+1 WHERE SrNo>" & txtSrNo & " AND EmpID='" & TBox(40) & "'"
                con.Execute "INSERT INTO EmpJobs(EmpID,SrNo,JobDescription) VALUES('" & TBox(40) & "'," & Val(txtSrNo) + 1 & ",'" & txtDuty & "')"
            Case 3
                con.Execute "UPDATE EmpJobs SET JobDescription='" & txtDuty & "' WHERE EmpID='" & TBox(40) & "' AND SrNo=" & txtSrNo & ""
            Case 4
        End Select
        
        FillJobs
        txtDuty = ""

        Exit Sub
    End If
    
    con.Execute "INSERT INTO EmpJobs(EmpID,SrNo,JobDescription) VALUES('" & TBox(40) & "'," & txtSrNo & ",'" & txtDuty & "')"
    txtDuty = ""
    
    FillJobs
    
End If

End Sub

Private Function CheckPrevSrNo() As Boolean

If LV.ListItems.count > 0 Then

    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Text = txtSrNo Then
                CheckPrevSrNo = True
            End If
        End With
    Next
End If

End Function

Private Sub FillJobs()

LV.ListItems.Clear
Dim ITM As ListItem

Dim Rec As New ADODB.Recordset
With Rec
    .CursorLocation = adUseClient
    .Open "SELECT * FROM EmpJobs WHERE EmpID='" & TBox(40) & "' ORDER BY SrNo", con, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            Set ITM = LV.ListItems.add(, , !SrNo)
            ITM.Tag = !JobID
            ITM.ListSubItems.add , , !JobDescription
            txtSrNo = Val(!SrNo) + 1
            .MoveNEXT

        Next
    End If
End With
End Sub


Private Function SaveSetting() As Boolean

'    If Val(Stxts(3)) = 0 Then
'        MsgBox "Over Time Hours Per Day Could Not Be Zero !", vbInformation
'        Stxts(3).SetFocus
'        Exit Function
'    End If
    
    Dim BGHoliday As Integer, AGHoliday As Integer, _
    BPHoliday As Integer, APHoliday As Integer

    'On Error GoTo ERR
    
    
    If Not cmbDept.MatchFound Then
      MsgBox "Please Select The Departement.", vbInformation
      cmbDept.SetFocus
      Exit Function
    
    End If
    
    Dim SQL As String, LHours As Double, LunchTime As Double
    
    BGHoliday = IIf(optGeneral(0).Value = True, 1, 0)
    AGHoliday = IIf(optGeneral(2).Value = True, 1, 0)
    BPHoliday = IIf(optPublic(0).Value = True, 1, 0)
    APHoliday = IIf(optPublic(2).Value = True, 1, 0)
   
    LHours = Round((dtLHours.Hour * 60 + dtLHours.Minute) / 60, 2)
    LunchTime = Round((DTLunch.Hour * 60 + DTLunch.Minute) / 60, 2)
    
        SQL = "Update EmpSettings set AbsAllowed=" & Val(Stxts(0) & "") & _
         ", TotalAbsAllowed=" & Val(Stxts(1) & "") & ",TimeAllowed=" & LHours & _
         ", NTimes=" & Val(Stxts(2) & "") & ",OverTime=" & Val(Stxts(3) & "") & _
         ", DinnerAmt=" & Val(Stxts(4) & "") & ",ShortTermPer=" & Val(Stxts(5) & "") & _
         ", AdvSalPer=" & Val(Stxts(6) & "") & _
         ", LongTermTime=" & Val(Stxts(7) & "") & ", LongTermLimit = '" & Stxts(8) & _
         "',EnterTime='" & Format(TEnter.Value, "HH:mm") & _
         "',ExitTime='" & Format(TExit.Value, "HH:mm") & _
         "',LunchTime=" & LunchTime & ", SocialAmt=" & Val(Stxts(10)) & _
         ", BGHoliday='" & BGHoliday & "', AGHoliday='" & AGHoliday & _
         "',BPHoliday='" & BPHoliday & "', APHoliday='" & APHoliday & _
         "',WorkingHrs=" & Val(Stxts(9).Tag) & ",AllowMultiple=" & Val(chkMultiple) _
         & ",DeductAbsents=" & Val(chkAbsents) & ",DeductLHrs=" & Val(chkLateHrs) & _
         ",LunchInTime='" & Format(LunchStart.Value, "HH:mm") & _
         "',LunchOutTime='" & Format(LunchEnd.Value, "HH:mm") & "',EmpAttCat_RefID=" & cmbCatagory.Text & _
         ",FridayBreakStartTime='" & Format(FridayBreakStart, "HH:mm") & "',FridayBreakEndTime='" & Format(FridayBreakEnd, "HH:mm") & _
         "',FridayCheckOut='" & Format(FridayCheckOut, "HH:mm") & _
         "' WHERE EmpID='" & TBox(40) & "'"
         
    Frame1.Enabled = False
    If IsNew Then
        NewEmployee.SqlSettings = SQL
        Me.Hide
    Else
        Dim a As Long
        con.Execute SQL, a
        'Unload Me
    End If
    
    SaveSetting = True
        
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub GetDeptSetting()

Dim rs As New ADODB.Recordset
With rs
    .CursorLocation = adUseClient
    .Open "SELECT * FROM GeneralSettings WHERE DeptID='" & cmbDeptID & "'", con, adOpenForwardOnly, adLockReadOnly
    If .RecordCount > 0 Then
    
    con.Execute "Update EmpSettings set AbsAllowed=" & Val(!AbsAllowed & "") & _
        ",TotalAbsAllowed=" & Val(!TotalAbsAllowed & "") & ",TimeAllowed=" & !TimeAllowed & _
        ",NTimes=" & Val(!NTimes & "") & ",OverTime=" & Val(!OverTime & "") & _
        ",DinnerAmt=" & Val(!DinnerAmt & "") & ",ShortTermPer=" & Val(!ShortTermPer & "") & _
        ",AdvSalPer=" & Val(!AdvSalPer & "") & _
        ",LongTermTime=" & Val(!longtermtime & "") & ", LongTermLimit = '" & !LongTermLimit & _
        "',EnterTime='" & !EnterTime & _
        "',ExitTime='" & !ExitTime & _
        "',LunchTime=" & !LunchTime & ", SocialAmt=" & !SocialAmt & _
        ", BGHoliday='" & Abs(!BGHoliday) & "', AGHoliday='" & Abs(!AGHoliday) & _
        "',BPHoliday='" & Abs(!BPHoliday) & "', APHoliday='" & Abs(!APHoliday) & _
        "',WorkingHrs=" & Val(Stxts(9).Tag) & ",DeductAbsents=" & Abs(!DeductAbsents) & ",DeductLHrs=" & _
        Abs(!DeductLHrs) & ",EmpAttCat_RefID=" & Val(!EmpAttCat_RefID & "") & " Where EmpID='" & TBox(40) & "'", a
    End If
   .Close
End With
Set rs = Nothing
End Sub


Public Sub EditSetting()
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM EmpSettings WHERE EmpID='" & TBox(40) & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            FirstEntry = False
            TEnter.Value = CDate(IIf(IsNull(!EnterTime), "09:00", !EnterTime))
            TExit.Value = CDate(IIf(IsNull(!ExitTime), "17:00", !ExitTime))
            LunchStart.Value = CDate(IIf(IsNull(!LunchInTime), "12:30", !LunchInTime))
            LunchEnd.Value = CDate(IIf(IsNull(!LunchOutTime), "13:30", !LunchOutTime))
            DTLunch.Value = CDate(IIf(IsNull(!LunchTime), "01:00", MinsToTime(Val(!LunchTime & "") * 60)))
            
            FridayBreakStart.Value = CDate(IIf(IsNull(!FridayBreakStartTime), "12:15", !FridayBreakStartTime))
            FridayBreakEnd.Value = CDate(IIf(IsNull(!FridayBreakEndTime), "14:00", !FridayBreakEndTime))
            FridayCheckOut.Value = CDate(IIf(IsNull(!FridayCheckOut), "18:00", !FridayCheckOut))
            
            Stxts(0) = ![AbsAllowed] & ""
            Stxts(1) = ![TotalAbsAllowed] & ""
            
            dtLHours = CDate(IIf(IsNull(!TimeAllowed), "09:00", MinsToTime(!TimeAllowed * 60)))
            Stxts(2) = ![NTimes] & ""
            Stxts(3) = ![OverTime] & ""
            
            Stxts(4) = ![DinnerAmt] & ""
            Stxts(5) = ![ShortTermPer] & ""
            Stxts(6) = ![AdvSalPer] & ""
            
            Stxts(7) = ![longtermtime] & ""
            Stxts(8) = ![LongTermLimit] & ""
            Stxts(10) = Val(!SocialAmt & "")
            
            chkAbsents = Abs(![DeductAbsents])
            chkLateHrs = Abs(![DeductLHrs])
            
            optGeneral(0).Value = IIf((!BGHoliday = True), True, False)
            optGeneral(1).Value = IIf((!BGHoliday = False), True, False)
            optGeneral(2).Value = IIf((!AGHoliday = True), True, False)
            optGeneral(3).Value = IIf((!AGHoliday = False), True, False)
            
            optPublic(0).Value = IIf((!BPHoliday = True), True, False)
            optPublic(1).Value = IIf((!BPHoliday = False), True, False)
            optPublic(2).Value = IIf((!APHoliday = True), True, False)
            optPublic(3).Value = IIf((!APHoliday = False), True, False)
            Stxts(9) = CalcNet
            chkMultiple = IIf(!AllowMultiple & "" = "", 0, Abs(!AllowMultiple))
            If Val(!EmpAttCat_RefID & "") <> 0 Then
                cmbCatagory.Text = Val(!EmpAttCat_RefID & "")
            End If
        Else
            FirstEntry = True
        End If
        .Close
    End With
    
    Set rs = Nothing
End Sub


Private Function CalcNet() As String

    Dim NetHours As Long

    NetHours = DateDiff("n", TEnter, TExit)
    NetHours = NetHours - (DTLunch.Hour + DTLunch.Minute) * 60

    Stxts(9).Tag = Round(NetHours / 60, 2)
    CalcNet = MinsToTime(NetHours)

End Function


Private Sub txtEditAssets_KeyPress(Index As Integer, KeyAscii As Integer)
        
    If KeyAscii = 13 Then
        If Index = txtEditAssets.UBound Then
            LVAssets.SelectedItem.SubItems(1) = StrConv(txtEditAssets(0), vbProperCase)
            LVAssets.SelectedItem.SubItems(2) = StrConv(txtEditAssets(1), vbProperCase)
            LVAssets.SelectedItem.SubItems(3) = StrConv(txtEditAssets(2), vbProperCase)
            PicEditAssets.Visible = False
        Else
            If txtEditAssets(Index) <> "" Then
                With txtEditAssets(Index + 1)
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
            End If
        End If
    ElseIf KeyAscii = 27 Then
        PicEditAssets.Visible = False
    End If
    
End Sub

Private Sub txtFilePath_DblClick()
    If txtFilePath.Tag <> "" Then
        Call ShowTechDrawing("Employees", "CNIC_PDF", " WHERE EmpID='" & TBox(40) & "'", TBox(1), CD2)
    End If
End Sub

Private Sub txts_GotFocus(Index As Integer)
    
    txts(Index).SelStart = 0
        txts(Index).SelLength = Len(txts(Index))

End Sub

Private Sub SaveGuarantors(EmpId As String)
       
    Exit Sub
    con.Execute "DELETE FROM EmpGuarantors WHERE EmpID='" & EmpId & "'"
    
    For i = 1 To LVGuarantors.ListItems.count
        With LVGuarantors.ListItems(i)
            con.Execute "INSERT INTO EmpGuarantors(EmpID,Name,Address,PhoneNo,KnownBy) VALUES('" & _
             EmpId & "','" & .SubItems(1) & "','" & .SubItems(2) & "','" & _
             .SubItems(3) & "','" & .SubItems(4) & "')"
        End With
    Next
    
End Sub


Private Sub LoadGuarantors(EmpId As String)
    
    Exit Sub
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM EmpGuarantors WHERE EmpID='" & EmpId & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVGuarantors.ListItems.Clear
        Do Until .EOF
           Set ITM = LVGuarantors.ListItems.add(, , LVGuarantors.ListItems.count + 1)
           ITM.ListSubItems.add , , !Name & ""
           ITM.ListSubItems.add , , !Address & ""
           ITM.ListSubItems.add , , !PhoneNo & ""
           ITM.ListSubItems.add , , !KnownBy & ""
           .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
End Sub


Private Sub AddHistory(p_strEmpID As String)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM EmpEmployementHistory WHERE EmpID='" & p_strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        LVHistory.ListItems.Clear
        Do Until .EOF
            Set ITM = LVHistory.ListItems.add(, , Format(!FromDT, "dd-MMM-yyyy"))
            ITM.ListSubItems.add , , Format(!ToDT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !company & ""
            ITM.ListSubItems.add , , !Designation & ""
            ITM.ListSubItems.add , , !Industry & ""
            ITM.ListSubItems.add , , !Responsibilities & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub SaveHistory(p_strEmpID As String)
    
    con.Execute "DELETE FROM EmpEmployementHistory WHERE EmpID='" & p_strEmpID & "'"
    Dim i As Integer
    For i = 1 To LVHistory.ListItems.count
        With LVHistory.ListItems(i)
            con.Execute "INSERT INTO EmpEmployementHistory(EmpID,FromDT,ToDT,Company,Designation,Industry,Responsibilities) " & _
            "VALUES('" & p_strEmpID & "','" & .Text & "','" & .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & _
             .SubItems(4) & "','" & .SubItems(5) & "')"
        End With
    Next
    
End Sub

Private Sub SaveAssets(EmpId As String)

    con.Execute "Delete From EmpAssets Where EmpID='" & EmpId & "'"
    For i = 1 To LVAssets.ListItems.count
        With LVAssets.ListItems(i)
            con.Execute "Insert Into EmpAssets(EmpID,SrNo,AssetDescription,AssetDT,Remarks) " & _
             "Values('" & EmpId & "'," & Val(.Text) & ",'" & .SubItems(1) & "','" & _
             .SubItems(2) & "','" & .SubItems(3) & "')"
        End With
    Next
    
End Sub

Private Sub LoadAssets(EmpId As String)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        
        .Open "Select * From EmpAssets Where EmpID='" & EmpId & "' Order By SrNo", con, adOpenForwardOnly, adLockReadOnly
        LVAssets.ListItems.Clear
        Do Until .EOF
            Set ITM = LVAssets.ListItems.add(, !EntryID & "'", ![SrNo] & "")
            ITM.SubItems(1) = !AssetDescription & ""
            ITM.SubItems(2) = Format(!AssetDT, "dd-MMM-yyyy")
            ITM.SubItems(3) = !Remarks & ""
            
            .MoveNEXT
        Loop
        .Close
        txtAssetSNo = LVAssets.ListItems.count + 1
    End With
        
    Set rs = Nothing
    
End Sub



Private Sub LoadMedicalInfo(EmpId As String)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM EmpMedicalCheckUps WHERE EmpID='" & EmpId & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVMC.ListItems.Clear
        Do Until .EOF
           Set ITM = LVMC.ListItems.add(, , LVMC.ListItems.count + 1)
           ITM.ListSubItems.add , , !MedicalTest & ""
           ITM.ListSubItems.add , , !Lab & ""
           ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
           ITM.ListSubItems.add , , !Status & ""
           ITM.ListSubItems.add , , !Treatment & ""
           ITM.ListSubItems.add , , !Remarks & ""
           .MoveNEXT
        Loop
        .Close
       
        .Open "SELECT * FROM EmpMedicalHistory WHERE EmpID='" & EmpId & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            cmbHBsAg.ListIndex = Val(!HBsAg & "")
            cmbAntiHcv.ListIndex = Val(!AntiHcv & "")
            cmbBloodSugar.ListIndex = Val(!BloodSugar & "")
            cmbBloodPressure.ListIndex = Val(!BloodPressure & "")
            cmbAllergy.ListIndex = Abs(!Allergy)
            TBox(38) = !AllergyDescription & ""
        Else
            cmbHBsAg.ListIndex = 2
            cmbAntiHcv.ListIndex = 2
            cmbBloodSugar.ListIndex = 2
            cmbBloodPressure.ListIndex = 2
            cmbAllergy.ListIndex = 0
            TBox(38) = ""
        End If
        .Close
    End With
    
    Set rs = Nothing
    
End Sub

Private Sub SaveMedicalInfo(EmpId As String)
   
    con.Execute "DELETE FROM EmpMedicalCheckUps WHERE EmpID='" & EmpId & "'"
    
    For i = 1 To LVMC.ListItems.count
        With LVMC.ListItems(i)
            con.Execute "INSERT INTO EmpMedicalCheckUps(EmpID,SrNo,MedicalTest,Lab,DT,Status,Treatment,Remarks) VALUES('" & _
             EmpId & "'," & Val(.Text) & ",'" & .SubItems(1) & "','" & .SubItems(2) & "','" & _
             .SubItems(3) & "','" & .SubItems(4) & "','" & .SubItems(5) & "','" & .SubItems(6) & "')"
        End With
    Next
    
    Dim iAffect As Integer
    con.Execute "UPDATE EmpMedicalHistory SET HBsAg=" & cmbHBsAg.ListIndex & ",AntiHcv=" & cmbAntiHcv.ListIndex & _
     ",BloodSugar=" & cmbBloodSugar.ListIndex & ",BloodPressure=" & cmbBloodPressure.ListIndex & ",Allergy=" & _
     cmbAllergy.ListIndex & ",AllergyDescription='" & TBox(17) & "' WHERE EmpID='" & EmpId & "'", iAffect
      
    If iAffect = 0 Then
        con.Execute "INSERT INTO EmpMedicalHistory(EmpID,HBsAg,AntiHcv,BloodSugar,BloodPressure,Allergy,AllergyDescription) VALUES('" & _
         EmpId & "'," & cmbHBsAg.ListIndex & "," & cmbAntiHcv.ListIndex & "," & cmbBloodSugar.ListIndex & "," & _
         cmbBloodPressure.ListIndex & "," & cmbAllergy.ListIndex & ",'" & TBox(17) & "')"
    End If

End Sub

Private Sub SaveFamily(EmpId As String)
    
    'Save Dependents
    con.Execute "DELETE From EmpDependants Where EmpID='" & EmpId & "'"
    
    For i = 1 To LVChildren.ListItems.count
        With LVChildren.ListItems(i)
            con.Execute "INSERT INTO EmpDependants(EmpID,Name,Relation,NICNo,MaritalStatus,DOB,Education,Occupation,Dependent,Class,Institute) Values('" & _
             EmpId & "','" & .Text & "','" & .SubItems(1) & "','" & .SubItems(2) & "','" & _
             .SubItems(5) & "','" & .SubItems(3) & "','" & .SubItems(6) & "','" & .SubItems(7) & "',1,'" & .SubItems(8) & "','" & .SubItems(9) & "')"
        End With
    Next
    
    For i = 1 To LVBro.ListItems.count
        With LVBro.ListItems(i)
            con.Execute "Insert Into EmpDependants(EmpID,Name,Relation,NICNo,MaritalStatus,DOB,Education,Occupation,Dependent) Values('" & _
             EmpId & "','" & .Text & "','" & .SubItems(1) & "','" & .SubItems(2) & "','" & _
             .SubItems(5) & "','" & .SubItems(3) & "','" & .SubItems(6) & "','" & .SubItems(7) & "',0)"
        End With
    Next
    
End Sub

Private Sub LoadFamily(EmpId As String)
    
    On err GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem, LV As ListView
    Dim myDT As Date
    Dim iYears As Integer, iMonths As Integer, iDays As Integer
    Dim Relation As String, MaritalStatus As String, Birth As Date
    Dim bDependent As Boolean
    With rs
        .Open "Select * From EmpDependants Where EmpID='" & EmpId & "' Order By DOB", con, adOpenForwardOnly, adLockReadOnly
        LVChildren.ListItems.Clear
        LVBro.ListItems.Clear
        Do Until .EOF
            bDependent = !Dependent
            Relation = !Relation & ""
            MaritalStatus = !MaritalStatus & ""
            Birth = !DOB
            If Relation = "Daughter" Then
                If MaritalStatus = "Single" Then
                    Set LV = LVChildren
                Else
                    Set LV = LVBro
                End If
            ElseIf Relation = "Son" Then
                If DateDiffInYears(Birth) < 18 Then
                    Set LV = LVChildren
                Else
                    Set LV = LVBro
                End If
            ElseIf Relation = "Wife" Or Relation = "Husband" Then
                Set LV = LVChildren
            Else
                If bDependent Then
                    Set LV = LVChildren
                Else
                    Set LV = LVBro
                End If
            End If
            
            Set ITM = LV.ListItems.add(, !EntryID & "'", !Name & "")
            
            ITM.SubItems(1) = ![Relation] & ""
            ITM.SubItems(2) = ![NICNo] & ""
            ITM.SubItems(3) = Format(![DOB], "dd-MMM-yyyy")
            ITM.SubItems(4) = DateDiffInText(![DOB])
            ITM.SubItems(5) = !MaritalStatus & ""
            ITM.SubItems(6) = !Education & ""
            ITM.SubItems(7) = !Occupation & ""
            ITM.ListSubItems.add , , ![Class] & ""
            ITM.ListSubItems.add , , ![Institute] & ""
            .MoveNEXT
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub SaveSuccessions(p_strEmpID As String)
    
    con.Execute "DELETE FROM EmpSuccessions WHERE EmpID='" & p_strEmpID & "'"
    Dim i As Integer
    For i = 1 To LVSuccessions.ListItems.count
        With LVSuccessions.ListItems(i)
            con.Execute "INSERT INTO EmpSuccessions(EmpID,SuccessionEmpID) VALUES('" & _
             p_strEmpID & "','" & .key & "')"
        End With
    Next
    
End Sub


Private Sub LoadSuccessions(p_strEmpID As String)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM VEmpSuccessions WHERE EmpID='" & p_strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        LVSuccessions.ListItems.Clear
        Do Until .EOF
            Set ITM = LVSuccessions.ListItems.add(, !SuccessionEmpID, LVSuccessions.ListItems.count + 1)
            ITM.ListSubItems.add , , "{ " & !EmpId & " } " & !Name
            ITM.ListSubItems.add , , !Designation & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub SaveReviews(EmpId As String)
    
    con.Execute "Delete From EmpReviews Where EmpID='" & EmpId & "'"
    
    For i = 1 To LVReview.ListItems.count
        With LVReview.ListItems(i)
            con.Execute "INSERT INTO EmpReviews(EmpID,DT,ReviewedBy,Review,UserName,MachineName,ReviewType,Rating) Values('" & _
             EmpId & "','" & .SubItems(1) & "','" & .SubItems(4) & "','" & .SubItems(5) & "','" & _
             CurrentUserName & "','" & strComputerName & "','" & .SubItems(2) & "','" & .SubItems(3) & "')"
        End With
    Next
    
End Sub

Private Sub LoadReviews(strEmpID As String)
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM EmpReviews WHERE EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        LVReview.ListItems.Clear
        Do Until .EOF
            Set ITM = LVReview.ListItems.add(, !EntryID & "'", LVReview.ListItems.count + 1)
            ITM.SubItems(1) = Format(!DT, "dd-MMM-yyyy")
            ITM.SubItems(2) = !ReviewType & ""
            ITM.SubItems(3) = !Rating & ""
            ITM.SubItems(4) = !ReviewedBy & ""
            ITM.SubItems(5) = !Review & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub SaveDailyTargets(EmpId As String)
        
    If strDelListEmpDT <> "" Then
        strDelListEmpDT = Left(strDelListEmpDT, Len(strDelListEmpDT) - 1)
        con.Execute "DELETE FROM EmpDailyTargets WHERE EntryID IN(" & strDelListEmpDT & ")"
    End If
    
    For i = 1 To LVDTItems.ListItems.count
        With LVDTItems.ListItems(i)
            If Val(.Tag) = 0 Then
                con.Execute "INSERT INTO EmpDailyTargets(EmpID,GroupID,ProcessID,Qty,UserName,MachineName,OverTime) VALUES('" & _
                 EmpId & "'," & Val(.ListSubItems(1).Tag) & "," & Val(.ListSubItems(2).Tag) & "," & Val(.SubItems(2)) & ",'" & _
                 CurrentUserName & "','" & strComputerName & "'," & .ListSubItems(3).Tag & ")"
            Else
                con.Execute "UPDATE EmpDailyTargets SET Qty=" & Val(.SubItems(2)) & ",DTLastUpdated=getDate() " & _
                 " WHERE EntryID=" & Val(.Tag)
            End If
        End With
    Next
    
End Sub

Private Sub SaveNotices(EmpId As String)
    
    con.Execute "DELETE FROM EmpNotices Where EmpID='" & EmpId & "'"
    
    For i = 1 To LVNotices.ListItems.count
        With LVNotices.ListItems(i)
            con.Execute "INSERT INTO EmpNotices(EmpID,DT,NoticeBy,Notice,UserName,MachineName) VALUES('" & _
             EmpId & "','" & .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & _
             CurrentUserName & "','" & strComputerName & "')"
        End With
    Next
    
End Sub

Private Sub LoadDailyTargets(strEmpID As String, Optional p_bCopy As Boolean = False)

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strItemName As String, strOverTime As String
    With rs
        .Open "SELECT * FROM VEmpDailyTargets WHERE EmpID='" & strEmpID & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVDTItems.ListItems.Clear
        Do Until .EOF
            Set ITM = LVDTItems.ListItems.add(, , !Description & "")
            If p_bCopy = False Then
                ITM.Tag = Val(!EntryID & "")
            Else
                ITM.Tag = "0"
            End If
            
'            If !ITemID = "Group" Then
'                strItemName = "Group of Items"
'            Else
'                strItemName = !ItemName & ""
'            End If
             strItemName = !ItemGroupName
            If !OverTime Then
                strOverTime = "Yes"
            Else
                strOverTime = "No"
            End If
            ITM.ListSubItems.add(, , strItemName).Tag = Val(!GroupID & "")
            ITM.ListSubItems.add(, , Val(!Qty & "")).Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add(, , strOverTime).Tag = Abs(!OverTime)
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub LoadNotices(strEmpID As String)

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM EmpNotices WHERE EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        LVNotices.ListItems.Clear
        Do Until .EOF
            Set ITM = LVNotices.ListItems.add(, !EntryID & "'", LVNotices.ListItems.count + 1)
            ITM.SubItems(1) = Format(!DT, "dd-MMM-yyyy")
            ITM.SubItems(2) = !NoticeBy & ""
            ITM.SubItems(3) = !Notice & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LoadSalaryHistory(strEmpID As String)
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    
    Cond = Cond & " AND TempDept=0 AND Active=1"
    SQL = "select * from VSalaryLedger " & Cond & " Order By DT Asc"

    With rs
        
        .Open "SELECT * FROM VSalaryLedger WHERE EmpID='" & strEmpID & "' ORDER BY DT ASC", con, adOpenForwardOnly, adLockReadOnly
        LVSalHistory.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LVSalHistory.ListItems.add(, , LVSalHistory.ListItems.count + 1)
            ITM.SubItems(1) = Format(!DT, "dd-MMM-yyyy")
            ITM.SubItems(2) = ![Description] & ""
            ITM.SubItems(3) = Val(![Salary] & "")
            .MoveNEXT
        Next
        .Close
    End With
    Set rs = Nothing
    
End Sub

Public Function AddNewEmployee(p_strDeptID As String, p_bTempEmp As Boolean, Optional p_strCVID As String)

    bTempEmp = p_bTempEmp
    cmbDept.Text = p_strDeptID
    TBox(40).Locked = False
    strCVID = p_strCVID
    
    If strCVID <> "" Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT CVID,Name,FName,DOB,MaritalStatus,Religion,Address,NIC,Phone1 FROM CVs WHERE CVID='" & strCVID & "'", con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                TBox(0) = !Name & ""
                TBox(1) = !FName & ""
                SO(0).Value = True
                DOB.Value = !DOB
                cmbMaritalStatus = !MaritalStatus & ""
                cmbReligion = !Religion & ""
                TBox(8) = !Address & ""
                TBox(4) = !NIC & ""
                TBox(6) = !Phone1 & ""
            End If
            .Close
        End With
    End If
    Set rs = Nothing
    
    Me.Show 1
    
End Function

Private Sub SaveCostingDetails(EmpId As String)
        
    If strDelListEmpCD <> "" Then
        strDelListEmpCD = Left(strDelListEmpCD, Len(strDelListEmpCD) - 1)
        con.Execute "DELETE FROM EmployeeCostingDetails WHERE EntryID IN(" & strDelListEmpCD & ")"
    End If
    
    For i = 1 To LVCD.ListItems.count
        With LVCD.ListItems(i)
            If Val(.Tag) = 0 Then
                con.Execute "INSERT INTO EmployeeCostingDetails(EmpID,ProcessID,ItemGroupID) VALUES('" & _
                 EmpId & "','" & .ListSubItems(1).Tag & "'," & Val(.ListSubItems(2).Tag) & ")"
'            Else
'                con.Execute "UPDATE EmpDailyTargets SET Qty=" & Val(.SubItems(2)) & ",DTLastUpdated=getDate() " & _
'                 " WHERE EntryID=" & Val(.Tag)
            End If
        End With
    Next
    
End Sub

Private Sub LoadCostingDetails(strEmpID As String, Optional p_bCopy As Boolean = False)

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strItemName As String, strOverTime As String
    With rs
        .Open "SELECT * FROM VEmployeeCostingDetails WHERE EmpID='" & strEmpID & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVCD.ListItems.Clear
        Do Until .EOF
            Set ITM = LVCD.ListItems.add(, , "")
            
            If p_bCopy = False Then
                ITM.Tag = Val(!EntryID & "")
            Else
                ITM.Tag = "0"
            End If
           
            ITM.ListSubItems.add(, , !ProcessDescription & "").Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add(, , !ItemGroupDescription & "").Tag = Val(!ItemGroupID & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdBrowseBarcode_Click()

    On Error Resume Next
    Dim Picfile As String
    CD2.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD2.FileName
        txtFilePath = Picfile
        txtFilePath.Tag = ""
        strFileName = CD2.FileTitle
    End If

End Sub

Private Sub AttachCNICPDF(p_strEmpID As String)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" And txtFilePath.Tag = "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EmpID,CNIC_PDF,CNIC_PDF_FileName FROM Employees WHERE EmpID='" & p_strEmpID & "'", con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            Rec.Fields(2) = strFileName
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub

