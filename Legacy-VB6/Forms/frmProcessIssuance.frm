VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmProcessIssuance 
   ClientHeight    =   10140
   ClientLeft      =   1110
   ClientTop       =   360
   ClientWidth     =   11805
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic_Iss_Multi 
      Appearance      =   0  'Flat
      BackColor       =   &H008ACC99&
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   1830
      ScaleHeight     =   4755
      ScaleWidth      =   8235
      TabIndex        =   77
      Top             =   1050
      Visible         =   0   'False
      Width           =   8265
      Begin VB.CommandButton cmdCancel_M 
         BackColor       =   &H008ACC99&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   6315
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   4050
         Width           =   1830
      End
      Begin VB.CommandButton cmdOK_M 
         BackColor       =   &H008ACC99&
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4410
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   4050
         Width           =   1830
      End
      Begin MSComctlLib.ListView LVIssd_Multi 
         Height          =   3945
         Left            =   -15
         TabIndex        =   78
         Top             =   -15
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   6959
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   255
         BackColor       =   14217174
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Process"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Rate"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin ComboList.Usercontrol1 cmbProcess 
      Height          =   285
      Left            =   1290
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   810
      Width           =   2355
      _ExtentX        =   4154
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
   Begin ComboList.Usercontrol1 cmbCountedBy 
      Height          =   285
      Left            =   7170
      TabIndex        =   82
      Top             =   4605
      Width           =   4125
      _ExtentX        =   7276
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
   Begin VB.TextBox txtBatchNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9225
      TabIndex        =   81
      Top             =   5745
      Width           =   1425
   End
   Begin VB.TextBox txtOpenQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10380
      TabIndex        =   76
      Top             =   810
      Width           =   810
   End
   Begin ComboList.Usercontrol1 cmbIssEmp 
      Height          =   285
      Left            =   1665
      TabIndex        =   72
      Top             =   4605
      Width           =   4125
      _ExtentX        =   7276
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
      Caption         =   "Special Instructions :"
      Height          =   630
      Left            =   420
      TabIndex        =   35
      Top             =   4920
      Width           =   10890
      Begin VB.TextBox txtSpecialInstructions 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   60
         TabIndex        =   36
         Top             =   210
         Width           =   10740
      End
   End
   Begin VB.Frame FraSkip 
      Caption         =   "Skip to :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3270
      Left            =   3195
      TabIndex        =   47
      Top             =   1545
      Visible         =   0   'False
      Width           =   5040
      Begin ComboList.Usercontrol1 cmbSkipProcess 
         Height          =   285
         Left            =   1005
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
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
      Begin VB.CommandButton cmdSkipCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   2535
         TabIndex        =   52
         Top             =   2760
         Width           =   1320
      End
      Begin VB.CommandButton cmdSkipOk 
         Caption         =   "Ok"
         Height          =   435
         Left            =   1170
         TabIndex        =   51
         Top             =   2760
         Width           =   1320
      End
      Begin VB.Label lblSkipDesc 
         Caption         =   "This will make the selected item to skip the current process && make it Issuable to the selected Process."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   645
         Left            =   210
         TabIndex        =   50
         Top             =   1365
         Width           =   4710
      End
      Begin VB.Label Label6A 
         AutoSize        =   -1  'True
         Caption         =   "Process :"
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
         Height          =   195
         Left            =   225
         TabIndex        =   48
         Top             =   270
         Width           =   750
      End
   End
   Begin ComboList.Usercontrol1 cmbVender 
      Height          =   285
      Left            =   3615
      TabIndex        =   2
      Top             =   810
      Width           =   4845
      _ExtentX        =   8546
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
   Begin VB.Frame FraStockOrder 
      Caption         =   "Article for Stock Order Issuance :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3330
      Left            =   630
      TabIndex        =   53
      Top             =   1380
      Visible         =   0   'False
      Width           =   10485
      Begin VB.TextBox txtSOQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9615
         TabIndex        =   58
         Top             =   465
         Width           =   765
      End
      Begin ComboList.Usercontrol1 cmbSOItems 
         Height          =   285
         Left            =   75
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   465
         Width           =   9510
         _ExtentX        =   16775
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
      Begin VB.CommandButton cmdSOOK 
         Caption         =   "Ok"
         Height          =   435
         Left            =   7680
         TabIndex        =   56
         Top             =   840
         Width           =   1320
      End
      Begin VB.CommandButton cmdSOCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   9060
         TabIndex        =   55
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Qty :"
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
         Height          =   195
         Left            =   9600
         TabIndex        =   59
         Top             =   255
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Article :"
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
         Height          =   195
         Left            =   75
         TabIndex        =   57
         Top             =   255
         Width           =   645
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
      Left            =   405
      TabIndex        =   63
      Top             =   9600
      Width           =   1920
   End
   Begin VB.Frame FraChecker 
      Caption         =   "Checker :"
      Height          =   630
      Left            =   4785
      TabIndex        =   61
      Top             =   4920
      Width           =   4200
      Begin MSForms.ComboBox cmbChecker 
         Height          =   285
         Left            =   45
         TabIndex        =   62
         Top             =   240
         Width           =   4080
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7197;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdUpdateRate 
      Cancel          =   -1  'True
      Caption         =   "Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2250
      TabIndex        =   60
      Top             =   5550
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox txtLotNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   420
      TabIndex        =   0
      Top             =   810
      Width           =   885
   End
   Begin VB.TextBox txtExcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "Fax2"
      DataSource      =   "datPrimaryRS"
      Height          =   330
      Left            =   4275
      MaxLength       =   100
      TabIndex        =   42
      Tag             =   "ExcessQtyPercentage"
      Text            =   "0"
      Top             =   9780
      Width           =   1110
   End
   Begin VB.TextBox txtRcvings 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "Fax2"
      DataSource      =   "datPrimaryRS"
      Height          =   330
      Left            =   3120
      MaxLength       =   100
      TabIndex        =   41
      Tag             =   "MaximumRcvingsAgainstPO"
      Text            =   "50"
      Top             =   9780
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox txtAssignedUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3660
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   5745
      Width           =   810
   End
   Begin MSComCtl2.DTPicker DTReturn 
      Height          =   285
      Left            =   7815
      TabIndex        =   33
      Top             =   5745
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   172032003
      CurrentDate     =   37711
   End
   Begin ComboList.Usercontrol1 cmbReturnProcessOld 
      Height          =   285
      Left            =   120
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4965
      Visible         =   0   'False
      Width           =   5280
      _ExtentX        =   9313
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
   Begin VB.TextBox txtVendRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1815
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   31
      Top             =   5745
      Width           =   1110
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   255
      Left            =   1110
      TabIndex        =   29
      Top             =   7170
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check &All"
      Height          =   255
      Left            =   30
      TabIndex        =   28
      Top             =   7170
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2910
      TabIndex        =   16
      Top             =   5745
      Width           =   765
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
      Left            =   6075
      TabIndex        =   13
      Top             =   9495
      Width           =   5280
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
         TabIndex        =   19
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
         Left            =   1770
         TabIndex        =   18
         Top             =   165
         Width           =   1665
      End
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
         TabIndex        =   5
         Top             =   165
         Width           =   1665
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3435
      Left            =   405
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6060
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   6059
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Vender Balances"
      TabPicture(0)   =   "frmProcessIssuance.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LVBal"
      Tab(0).Control(1)=   "txtTotal"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Vender History"
      TabPicture(1)   =   "frmProcessIssuance.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5(2)"
      Tab(1).Control(1)=   "Label5(3)"
      Tab(1).Control(2)=   "LVMaker"
      Tab(1).Control(3)=   "txtRate"
      Tab(1).Control(4)=   "txtTotalQtyHistory"
      Tab(1).Control(5)=   "txtTotalValueHistory"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Issued Items"
      TabPicture(2)   =   "frmProcessIssuance.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "LVIssued"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtRate1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtTotalValue"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtTotalQty"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "chkSampleProvided"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkDrawingProvided"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkForgingProvided"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "chkSteelProvided"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Other Vender's History"
      TabPicture(3)   =   "frmProcessIssuance.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LVMakerOther"
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtTotalValueHistory 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -67785
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   3000
         Width           =   1140
      End
      Begin VB.TextBox txtTotalQtyHistory 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -69825
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   3000
         Width           =   1080
      End
      Begin VB.CheckBox chkSteelProvided 
         Appearance      =   0  'Flat
         Caption         =   "Steel Provided"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5250
         TabIndex        =   67
         Top             =   3075
         Width           =   1605
      End
      Begin VB.CheckBox chkForgingProvided 
         Appearance      =   0  'Flat
         Caption         =   "Forging Provided"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3525
         TabIndex        =   66
         Top             =   3075
         Width           =   1605
      End
      Begin VB.CheckBox chkDrawingProvided 
         Appearance      =   0  'Flat
         Caption         =   "Drawing Provided"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   65
         Top             =   3075
         Width           =   1605
      End
      Begin VB.CheckBox chkSampleProvided 
         Appearance      =   0  'Flat
         Caption         =   "Sample Provided"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         TabIndex        =   64
         Top             =   3075
         Width           =   1605
      End
      Begin VB.TextBox txtTotalQty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7695
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3030
         Width           =   1080
      End
      Begin VB.TextBox txtTotalValue 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9735
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3030
         Width           =   1140
      End
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   -65730
         TabIndex        =   27
         Top             =   3075
         Width           =   1590
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70065
         TabIndex        =   24
         Top             =   630
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtRate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2145
         TabIndex        =   22
         Top             =   1575
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSComctlLib.ListView LVIssued 
         Height          =   2670
         Left            =   75
         TabIndex        =   23
         Top             =   375
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   4710
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vender"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Return Process"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Buyer"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Order No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Return Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Issued Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Value"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Priority"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Employee"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Batch No."
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView LVMaker 
         Height          =   2580
         Left            =   -74925
         TabIndex        =   25
         Top             =   375
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   4551
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Description"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Rate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Value"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Del. Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Priority"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSComctlLib.ListView LVBal 
         Height          =   2685
         Left            =   -74940
         TabIndex        =   26
         Top             =   360
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   4736
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
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Description"
            Object.Width           =   7761
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Balance"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView LVMakerOther 
         Height          =   2970
         Left            =   -74940
         TabIndex        =   34
         Top             =   375
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   5239
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Maker"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Description"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Rate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Value"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Del. Date"
            Object.Width           =   1940
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Value :"
         Height          =   195
         Index           =   3
         Left            =   -68700
         TabIndex        =   71
         Top             =   3030
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Qty :"
         Height          =   195
         Index           =   2
         Left            =   -70635
         TabIndex        =   70
         Top             =   3030
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Qty :"
         Height          =   195
         Index           =   1
         Left            =   6885
         TabIndex        =   46
         Top             =   3060
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Value :"
         Height          =   195
         Index           =   0
         Left            =   8820
         TabIndex        =   38
         Top             =   3060
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&POST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10695
      Picture         =   "frmProcessIssuance.frx":0070
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   600
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   420
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessIssuance.frx":0172
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessIssuance.frx":032E
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   -195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessIssuance.frx":04EA
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessIssuance.frx":093E
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
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
      Left            =   795
      TabIndex        =   6
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Items"
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
         Left            =   4440
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Items"
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
         Left            =   4455
         TabIndex        =   8
         Top             =   135
         Width           =   1695
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   420
      TabIndex        =   17
      Top             =   5760
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   172097539
      CurrentDate     =   37711
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2820
      Left            =   435
      TabIndex        =   10
      Top             =   1635
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   4974
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
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Maker"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Order No"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   5953
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Issueable Qty"
         Object.Width           =   2328
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Order Qty"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Internal Order"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   165
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3885
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   172097539
      CurrentDate     =   37711
   End
   Begin VB.Frame FraLoad 
      Height          =   3255
      Left            =   420
      TabIndex        =   20
      Top             =   1650
      Visible         =   0   'False
      Width           =   10860
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Loading Data"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   90
         TabIndex        =   21
         Top             =   780
         Width           =   10605
      End
   End
   Begin VB.Label lblCountedBy 
      AutoSize        =   -1  'True
      Caption         =   "Counted By :"
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
      Height          =   195
      Left            =   5925
      TabIndex        =   83
      Top             =   4650
      Width           =   1050
   End
   Begin MSForms.ComboBox cmbSteel 
      Height          =   315
      Left            =   420
      TabIndex        =   75
      Tag             =   "SteelUsed"
      Top             =   1290
      Width           =   10770
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "18997;556"
      MatchEntry      =   1
      ListStyle       =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblSteel 
      Height          =   240
      Left            =   420
      TabIndex        =   74
      Top             =   1080
      Width           =   10770
      BackColor       =   11517387
      Caption         =   " Steel"
      Size            =   "18997;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Employee :"
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
      Height          =   195
      Left            =   525
      TabIndex        =   73
      Top             =   4635
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Excess Qty %"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   4290
      TabIndex        =   44
      Top             =   9570
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max. Rcvings"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   3120
      TabIndex        =   43
      Top             =   9570
      Visible         =   0   'False
      Width           =   960
   End
   Begin MSForms.ComboBox cmbReturnProcess 
      Height          =   285
      Left            =   4455
      TabIndex        =   39
      Top             =   5745
      Width           =   3345
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5900;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbOrderNo 
      Height          =   285
      Left            =   165
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   1965
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3466;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   420
      TabIndex        =   15
      Top             =   5520
      Width           =   10230
      BackColor       =   11517387
      Caption         =   $"frmProcessIssuance.frx":0D92
      Size            =   "18045;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbDesc 
      Height          =   285
      Left            =   165
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4155
      Visible         =   0   'False
      Width           =   3300
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5821;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbItems 
      Height          =   285
      Left            =   8475
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   810
      Width           =   1905
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3360;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   420
      TabIndex        =   9
      Top             =   570
      Width           =   10770
      BackColor       =   11517387
      Caption         =   $"frmProcessIssuance.frx":0E34
      Size            =   "18997;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop up Menu"
      Begin VB.Menu mnuSkipProcess 
         Caption         =   "Skip Process"
      End
      Begin VB.Menu mnuStockOrderIssuance 
         Caption         =   "Stock Order Issuance"
      End
   End
End
Attribute VB_Name = "frmProcessIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyCol As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean
Dim IssIDs As New Collection
Dim rsMain As New ADODB.Recordset
Dim IsLotProcess As Boolean
Dim GotIssID As Long
Dim RcvingLot As Boolean
Dim UserReqAuth As Integer 'Whether Or Not User Required Authorization On Issue For This Process.

Dim lVendBal As Long
Dim iColNo As Integer
Dim bFromMaker As Boolean

Dim bReIssuance As Boolean
Dim bDontUpdateFromProcess As Boolean
Dim bDontUpdateFromMaker As Boolean
Dim lCheckerProcessNo As Long
Dim lVI_EntryID_For_Print As Long

Private Sub AddItems()
    
    On Error GoTo err
    Exit Sub
    Dim OrCond As String
    'Create In List
    
    'For i = 1 To MyCol.Count
    '    OrCond = OrCond & "'" & MyCol(i).CompItemCode & "',"
    'Next i
    
    If Len(InList) > 0 Then
        InList = Mid$(InList, 1, Len(InList) - 1)
    End If
    'InList = InList & ")"
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "SELECT  * FROM VVendAssItems", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            For i = 1 To MyCol.count
                If MyCol(i).CompItemCode = ![ITemID] Then
                    cmbItems.AddItem ![ITemID] 'MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 1) = MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 2) = MyCol(i).lQty
                    cmbItems.List(cmbItems.ListCount - 1, 3) = ![Rate]
                    cmbItemDesc.AddItem MyCol(i).ItemDescription
                    Exit For
                End If
            Next i
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()

    If cmbDesc.MatchFound = False Then Exit Sub
    cmbItems.ListIndex = cmbDesc.ListIndex
    
    
    If RcvingLot Then
        LV.ListItems(cmbDesc.ListIndex + 1).Selected = True
    End If
    
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()

    If Not cmbItems.MatchFound Then Exit Sub
    'cmbDesc.ListIndex = cmbItems.ListIndex
    'txtVendRate = Val(cmbItems.List(cmbItems.ListIndex, 1))
    If bReIssuance Then
        Call FillList(cmbProcess.ID, cmbVender.ID, cmbItems.Text)
    Else
        Call FillList
    End If
    '/\/\/\/\/\/\/\/\/\/
    Dim rs As New ADODB.Recordset, lSNo As Long
    With rs
        .Open "SELECT SNO FROM ItemProcesses WHERE ProcessID=" & cmbProcess.ID & " AND ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lSNo = Val(!SNo & "")
        End If
        .Close
    End With
    
    
'    cmbReturnProcess.ClearVals
'    cmbReturnProcess.AddVals con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & _
'     cmbItems.Text & "' AND ProcessID IN(SELECT ProcessID FROM MakerProcesses WHERE VendID=" & _
'     cmbVender.ID & ") AND ItemSNO>=" & lSNo & " ORDER BY ItemSNo"
'
'    cmbReturnProcess.ID = cmbProcess.ID
'Or cmbProcess.ID = iSemiFinishProcessID

'    If (cmbProcess.ID = iReadyFinishProcessID) Then
'        rs.Open "SELECT Rate,Unit FROM VendAssItems WHERE VendID=" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID & " AND ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
'        With cmbReturnProcess
'            .Clear
'            .AddItem cmbProcess.Text '"Ready Finish"
'            .List(0, 1) = cmbProcess.ID
'            .List(0, 2) = Val(rs!Rate & "")
'            .List(0, 3) = rs!Unit & ""
'        End With
'        cmbReturnProcess.ListIndex = 0
'    Else
'        Set rs = con.Execute("EXEC NextProcessesForIssuance_SP " & cmbProcess.ID & "," & cmbVender.ID & ",'" & cmbItems.Text & "'")
'        Dim i As Integer
'        cmbReturnProcess.Clear
'        Do Until rs.EOF
'            cmbReturnProcess.AddItem rs!Description & ""
'            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 1) = Val(rs!ProcessID & "")
'            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 2) = Val(rs!Rate & "")
'            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 3) = rs!Unit & ""
'            rs.MoveNEXT
'        Loop
'        If cmbReturnProcess.ListCount > 0 Then cmbReturnProcess.ListIndex = 0
'    End If
'    rs.Close
    cmbReturnProcess.Clear
    cmbReturnProcess.AddItem cmbProcess.Text
    cmbReturnProcess.List(0, 1) = cmbProcess.ID
    cmbReturnProcess.List(0, 2) = 1 'Val(rs!Rate & "")
    cmbReturnProcess.List(0, 3) = "Pcs" 'rs!Unit & ""
    cmbReturnProcess.ListIndex = 0
    Set rsz = Nothing
    '/\/\/\/\/\/\/\/\/\/
    'Call FillVendList
    Call ShowOtherVendersHistory
    
    Call ShowOpenQty
End Sub

Private Sub ShowOpenQty()

    If cmbVender.MatchFound = False Or cmbItems.MatchFound = False Then
        txtOpenQty = ""
    Else
        Dim rs As New ADODB.Recordset, ITM As ListItem
        Dim strCondition As String
        strCondition = " AND Closed=0 "
        
        strCondition = strCondition & " AND ItemCode='" & cmbItems.Text & "' AND VendIssued.VendID=" & cmbVender.ID
       
        With rs
            .Open "SELECT SUM(VendIssdDetail_ReturnDTs.IssQty-VendIssdDetail_ReturnDTs.RcvQty+ISNULL(WastageQty,0)) AS Qty " & _
             "FROM VendIssued INNER JOIN VendIssdDetail ON VendIssued.EntryID=VendIssdDetail.RefID " & _
             "INNER JOIN VendIssdDetail_ReturnDTs ON VendIssdDetail.EntryID=VendIssdDetail_ReturnDTs.VIS_RefID " & _
             "INNER JOIN Makers ON VendIssued.VendID=Makers.VendID LEFT OUTER JOIN VendIssdDetail_MoreDetails ON VendIssdDetail.EntryID=VendIssdDetail_MoreDetails.VID_RefID " & _
             "LEFT OUTER JOIN VIssuanceEntryWithWastage ON VendIssdDetail.EntryID=VIssuanceEntryWithWastage.EntryID " & _
             "WHERE LotNo='0' AND VendIssdDetail_ReturnDTs.IssQty+CASE WHEN ISNULL(IssRcvBalance,0)<0 THEN ABS(ISNULL(IssRcvBalance,0)) ELSE 0 END+ISNULL(Wastage_Return,0)>VendIssdDetail_ReturnDTs.RcvQty " & strCondition, con, adOpenForwardOnly, adLockReadOnly
             
            If .EOF Then
                txtOpenQty = ""
            Else
                txtOpenQty = Val(!Qty & "")
            End If
            
            .Close
        End With
        Set rs = Nothing
    End If
    
End Sub

Private Sub ShowOtherVendersHistory()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select Top 10 * From VMakerIssItems Where VendID<>" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID & " AND ItemCode='" & cmbItems & "' Order By DT Desc,EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        
        LVMakerOther.ListItems.Clear
        'LVBal.ListItems.Clear
        lVendBal = 0
        Do Until .EOF
            Set ITM = LVMakerOther.ListItems.add(, , LVMakerOther.ListItems.count + 1)
            ITM.SubItems(1) = ![VendID1] & ""
            ITM.SubItems(2) = ![ItemCode] & ""
            ITM.SubItems(3) = ![ItemName] & ""
            ITM.SubItems(4) = Format(![DT] & "", "dd-MM-yy")
            ITM.SubItems(5) = Val(![IssQty] & "")
            ITM.SubItems(6) = Val(![Rate] & "")
            ITM.SubItems(7) = Val(!IssValue & "") 'Round(Val(![IssQty] & "") * Val(![Rate] & ""))
            ITM.SubItems(8) = Format(![ReturnDT] & "", "dd-MM-yy")
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub FillList(Optional iProcessID As Integer = 0, Optional iMakerID As Integer = 0, Optional strItemID As String = "", Optional strLotNo As String = "")

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    If bReIssuance = False Then
        If cmbItems.MatchFound = False Then Exit Sub
        
        With cmd
            Set .ActiveConnection = con_ServerSide
            .CommandType = adCmdStoredProc
            .CommandText = "IssueItems_SP"
            .Parameters("@ProcessID").Value = cmbProcess.ID
            .Parameters("@ItemID").Value = cmbItems.Text
            Set rs = .Execute
        End With
        
    Else
        With cmd
            Set .ActiveConnection = con_ServerSide
            .CommandType = adCmdStoredProc
            .CommandText = "ReIssueItems_Lotwise_SP"
            .Parameters("@ProcessID").Value = iProcessID
            .Parameters("@ReWorkLot").Value = GetSingleBooleanValue("ReWorkLot", "VendIssdDetail", " WHERE LotNo='" & strLotNo & "'")
            .Parameters("@LotNo").Value = strLotNo
            Set rs = .Execute
        End With
'        With cmd
'            Set .ActiveConnection = con
'            .CommandType = adCmdStoredProc
'            .CommandText = "ReIssueItemsNew_SP"
'            .Parameters("@ProcessID").Value = iProcessID
'            .Parameters("@MakerID").Value = iMakerID
'            .Parameters("@ItemID").Value = strItemID
'            .Parameters("@ReWorkLot").Value = GetSingleBooleanValue("ReWorkLot", "VendRcvdDetail", " WHERE LotNo='" & strLotNo & "'")
'            If strLotNo <> "" Then
'                .Parameters("@LotNo").Value = strLotNo
'            End If
'            Set rs = .Execute
'        End With
        
    End If
    
    Dim lQty As Long
    
    Dim bUpdateItemsCombo As Boolean
    If cmbVender.MatchFound And cmbItems.MatchFound = False Then
        bUpdateItemsCombo = True
    Else
        bUpdateItemsCombo = False
    End If
    
    With rs
        If bUpdateItemsCombo Then cmbItems.Clear
            
        LV.ListItems.Clear
        Do Until .EOF
            
            If bReIssuance Then
                Set ITM = LV.ListItems.add(, !Rcvd_RefID & !VendID & !OrderNo & !CompITemID & "'", LV.ListItems.count + 1)
                ITM.ListSubItems.add(, , !VendID1 & "").Tag = !VendID
                ITM.ListSubItems.add(, , ![OrderNo] & "").Tag = !Rcvd_RefID
            Else
'                Set ITM = LV.ListItems.add(, !OrderNo & !CompITemID & "'", LV.ListItems.count + 1)
                Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
                ITM.ListSubItems.add(, , "-").Tag = ""
                ITM.ListSubItems.add(, , ![OrderNo] & "").Tag = 0
            End If
            
            
            ITM.ListSubItems.add , , ![CompITemID] & ""
            ITM.ListSubItems.add , , ![ItemName] & ""
            
            If (cmbProcess.ID = !InitialProcessID Or ((cmbProcess.ID = iReadyFinishProcessID Or cmbProcess.ID = iSemiFinishProcessID Or cmbProcess.ID = iFilingProcessID Or cmbProcess.ID = iFiling2ProcessID))) And strLotNo = "" Then
                lQty = Val(!OrderQty & "") - Val(!TotIssdProc & "")
            Else
                lQty = Val(!IssQty & "")
            End If
            
            '-----------------------------------------
            If strCompany = "QEL" Then
                If bReIssuance = False Then
                    lQty = lQty - Val(!AdjustedQty & "")
                End If
            End If
            '-----------------------------------------
            ITM.ListSubItems.add(, , lQty).Tag = lQty
            
            ITM.ListSubItems.add(, , Val(!OrderQty & "")).Tag = Val(!TotIssdProc & "")
            ITM.ListSubItems.add , , !InternalRefNo & ""
            
            lColor = RGB(150, 55, 150)
            ITM.ForeColor = lColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lColor
            Next
            If bUpdateItemsCombo Then cmbItems.AddItem !CompITemID & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub cmbProcessOld_matched()
    
    On Error GoTo err
'    If cmbProcess.ID = 9 Then
'        cmdCheckAll.Visible = True
'        cmdClearAll.Visible = True
'    Else
'        cmdCheckAll.Visible = False
'        cmdClearAll.Visible = False
'    End If
    
    Dim cmd As New ADODB.Command
    
    LV.ListItems.Clear
    LVIssued.ListItems.Clear
    LVMaker.ListItems.Clear
    
    IsLotProcess = LotProcess(cmbProcess.ID)
    
    If IsLotProcess And RcvLot(cmbProcess.ID) Then
        RcvingLot = True
        cmbItems.Visible = False
        DTPicker1.Visible = True
        Label2.Caption = "  Process                                    Date                                      Lot No"
        LV.Checkboxes = True
        
        txtQty.Visible = False
        Label3.Caption = "  Venders                                                                Date"
        Label3.Width = 4830
        cmdPost.Left = DTPicker2.Left + DTPicker2.Width + 20
    Else
        RcvingLot = False
        cmbItems.Visible = True
        DTPicker1.Visible = False
        Label2.Caption = "  Process                                    Item Code                        Item Description"
        LV.Checkboxes = False
        
        txtQty.Visible = True
        Label3.Caption = "  Venders                                                                Date                      Quantity"
        Label3.Width = 5970
        cmdPost.Left = txtQty.Left + txtQty.Width + 20
        Me.Refresh
    End If
    
    Dim lColor As Long
    Dim ITM As ListItem
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        
        
        If IsLotProcess Then
            
            If RcvingLot Then
                LV.ColumnHeaders.add 2, , "Lot No.", 1200
                'LV.ColumnHeaders(2).Text = "Lot No/Order No"
                .CommandText = "SP_IssueItemsLots"
                txtQty.Locked = True
                txtRate1.Locked = True
            Else
                LV.ColumnHeaders(3).Text = "Order No"
                .CommandText = "SP_IssueItems"
                txtQty.Locked = False
                txtRate1.Locked = False
            End If
            cmbProcess.Tag = "Lot No"
        Else
            LV.ColumnHeaders(3).Text = "Order No"
            .CommandText = "SP_IssueItems"
            cmbProcess.Tag = ""
        End If
        
        If rsMain.State = 1 Then rsMain.Close
        rsMain.CursorType = adOpenForwardOnly
        rsMain.LockType = adLockReadOnly
        
        .Parameters("@ProcessID").Value = cmbProcess.ID
            
        Set rsMain = .Execute
    End With
    
    Dim isAdded As Boolean
    
    With rsMain
        cmbItems.Clear
        cmbDesc.Clear
        
        If RcvingLot Then
'            DTPicker1.value = ServerDate
            Do Until .EOF
                isAdded = False
                For i = 0 To cmbItems.ListCount - 1
                    If Trim(cmbItems.List(i)) = Trim(![RecieptID] & "") Then
                        isAdded = True
                        Exit For
                    End If
                Next i
                
                If Not isAdded Then
                    cmbItems.AddItem ![RecieptID] & ""
                    cmbDesc.AddItem ![RecieptID] & ""
                End If
                    
                
                Set ITM = LV.ListItems.add(, ![LotNo] & "''", LV.ListItems.count + 1)
                ITM.Tag = ![RecieptID] & ""
                'Debug.Print LV.ListItems.Count
                'ITM.Tag = ![EntryID]
                'itm.ListSubItems.Add , , ![CustCode] & ""
                ITM.ListSubItems.add(, , ![RecieptID] & "").Tag = ![LotNo] & ""
                
                ITM.ListSubItems.add(, , ![ItemCode] & "").Tag = ![ItemCode] & ""
                
                'itm.ListSubItems.Add , , ![ItemCode] & ""
                ITM.ListSubItems.add , , ![ItemDescription] & ""
                ITM.ListSubItems.add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                
                If IsNull(![LockedItemCode]) Then
                    lColor = LV.ForeColor
                Else
                    lColor = vbWhite
                End If
                
                ITM.ForeColor = lColor
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = lColor
                Next
                
                .MoveNEXT
            Loop
        Else
            Do Until .EOF
                isAdded = False
                For i = 0 To cmbItems.ListCount - 1
                    If Trim(cmbItems.List(i)) = Trim(![ItemCode] & "") Then
                        isAdded = True
                        Exit For
                    End If
                Next i
                
                If Not isAdded Then
                    cmbItems.AddItem ![ItemCode] & ""
                    cmbDesc.AddItem ![ItemDescription] & ""
                End If
                
                
            
                Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            
            
                'ITM.Tag = ![EntryID]
                
                ITM.ListSubItems.add , , ![ItemCode] & ""
                ITM.ListSubItems.add , , ![ItemDescription] & ""
                ITM.ListSubItems.add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
                
                If IsNull(![LockedItemCode]) Then
                    lColor = LV.ForeColor
                Else
                    lColor = RGB(150, 55, 150)
                End If
                
                ITM.ForeColor = lColor
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = lColor
                Next
                .MoveNEXT
            Loop
        End If
    End With
    
    If LV.ListItems.count > 0 Then
        LV.ListItems(1).Selected = True
        Call LV_ItemClick(LV.ListItems(1))
    End If
    
    Me.Show
    cmbVender.SetFocus
    SendKeys "a"
    SendKeys Chr(8)
    
    FraLoad.Visible = False
    LV.Visible = True
    'Debug.Print Time
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbProcess_matched()
    
    Dim iVendID As Integer
    
    If cmbVender.MatchFound Then
        iVendID = cmbVender.ID
    End If
    
    cmbVender.ClearVals
    If cmbProcess.ID <> "0" Then
        Dim iOperation As Integer
        iOperation = GetSingleLongValue("Operation", "Processes", " WHERE ProcessID=" & cmbProcess.ID)
        Dim bRepairLot As Boolean
        bRepairLot = GetSingleBooleanValue("ReWorkLot", "VendRcvdDetail", " WHERE LotNo='" & txtLotNo & "'")
        
        If iOperation = 0 Then
            cmbVender.AddVals con_ServerSide, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID=" & iFactoryMakerID
            'cmbVender.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID=" & iFactoryMakerID & " AND VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ")"
        ElseIf iOperation = 1 Then
            cmbVender.AddVals con_ServerSide, "'{ ' + VendID1 + ' } ' + VenderName+ ' C='+CAST(MakerCapacity AS VARCHAR(10))+' B='+CAST(Balance AS VARCHAR(10))", "VMakerBalances", "VendID", " WHERE VendID<>" & iFactoryMakerID & " AND VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ") AND Active=1"
        ElseIf iOperation = 2 Then
            cmbVender.AddVals con_ServerSide, "'{ ' + VendID1 + ' } ' + VenderName+ ' C='+CAST(MakerCapacity AS VARCHAR(10))+' B='+CAST(Balance AS VARCHAR(10))", "VMakerBalances", "VendID", " WHERE VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ") AND Active=1"
        End If
    Else
        cmbVender.AddVals con_ServerSide, "'{ ' + VendID1 + ' } ' + VenderName+ ' C='+CAST(MakerCapacity AS VARCHAR(10))+' B='+CAST(Balance AS VARCHAR(10))", "VMakerBalances", "VendID", " WHERE Active=1"
    End If
    
    If iVendID <> 0 Then
        cmbProcess.Tag = "From Process"
        bDontUpdateFromMaker = True
        cmbVender.ID = iVendID
        bDontUpdateFromMaker = False
    Else
        cmbProcess.Tag = ""
    End If
    
'    If bDontUpdateFromProcess = False Then
'        If bReIssuance Then
'            'Fill Issuable Items...
'            Call FillList(cmbProcess.ID)
'        End If
'    End If
    
    If cmbProcess.ID = lCheckerProcessNo Then
        Dim strDesignation As String
        strDesignation = getGeneralDataValue("CheckerDesignation")
        '-------------------------------------------------------------
        cmbChecker.Clear
        Call AddToCombo(cmbChecker, "EmpID + ' ' + Name", "Employees", " WHERE Designation='" & strDesignation & "' AND Active=1", , "EmpID")
        FraChecker.Visible = True
        Frame1.Width = 4350
        txtSpecialInstructions.Width = 4200
    Else
        FraChecker.Visible = False
        Frame1.Width = 8550
        txtSpecialInstructions.Width = 8415
    End If
    
End Sub

Private Sub cmbReturnProcess_Change()
    Call cmbReturnProcess_Click
End Sub


Private Sub cmbReturnProcess_Click()

    If cmbReturnProcess.MatchFound = False Then Exit Sub
    Dim iProcessNo As Integer, dRate As Double
    dRate = 0
    iProcessNo = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1)
    If (iProcessNo = iReadyFinishProcessID) Or (iProcessNo = iSemiFinishProcessID) Or (iProcessNo = iFilingProcessID) Or (iProcessNo = iFiling2ProcessID) Then 'Ready Finish
        dRate = Val(cmbReturnProcess.List(cmbReturnProcess.ListIndex, 2))
    Else
        Dim i As Integer
        For i = 0 To cmbReturnProcess.ListIndex
            dRate = dRate + Val(cmbReturnProcess.List(i, 2))
        Next
    End If
    txtAssignedUnit = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 3) & ""
    txtVendRate = dRate
    
End Sub

Private Sub cmbVender_matched()

    On Error GoTo err
    If bDontUpdateFromMaker Then Exit Sub
'    If cmbVender.Text = "HICO" Then Exit Sub
'    If RcvingLot Then
'        cmbVender.Tag = Val(con.Execute("Select Rate From VMakerAssItems Where VendID=" & cmbVender.ID & " and ItemID='" & LV.SelectedItem.SubItems(2) & "'").Fields(0).value & "")
'    Else
'        cmbVender.Tag = Val(con.Execute("Select Rate From VMakerAssItems Where VendID=" & cmbVender.ID & " and ItemID='" & LV.SelectedItem.SubItems(1) & "'").Fields(0).value & "")
'    End If
    
    Dim iProcessID As Integer
    
    If cmbProcess.MatchFound Then
        iProcessID = cmbProcess.ID
    End If
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT VendID1 FROM Makers WHERE VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False And cmbProcess.ID = "0" Then
            cmbVender.Tag = !VendID1 & ""
            'cmbProcess.ID = Val(!PhaseID & "")
        End If
        .Close
        
        If cmbProcess.Tag = "" Then
            cmbProcess.ClearVals
            cmbProcess.AddVals con, "Description", "VMakerAssignedProcesses", "ProcessID", " WHERE VendID=" & cmbVender.ID
            
            .Open "SELECT * FROM VMakerInitialProcess WHERE VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
            If iProcessID <> 0 Then
                bDontUpdateFromProcess = True
                cmbProcess.ID = iProcessID
                bDontUpdateFromProcess = False
            ElseIf .EOF = False Then
                cmbProcess.ID = Val(!FirstProcessID & "")
            End If
            .Close
        End If
        
        If bReIssuance = False Then
            '.Open "SELECT * FROM VMakerAssItems WHERE VendID=" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
            '.Open "SELECT ItemID FROM Items WHERE ItemID IN(SELECT ItemID FROM ItemProcesses WHERE ProcessID=" & cmbProcess.ID & ")", con_ServerSide, adOpenForwardOnly, adLockReadOnly
            .Open "SELECT ItemID FROM Items WHERE ItemID IN(SELECT ItemID FROM ItemProcesses WHERE ProcessID=" & cmbProcess.ID & ") AND ItemID IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & cmbVender.ID & " AND ProcessID=" & cmbProcess.ID & ")", con_ServerSide, adOpenForwardOnly, adLockReadOnly
            LV.ListItems.Clear
            cmbItems.Clear
            'cmbDesc.Clear
            Do Until .EOF
                cmbItems.AddItem !ITemID & ""
                cmbItems.List(cmbItems.ListCount - 1, 1) = 1  'Val(!Rate & "")
                'cmbDesc.AddItem !Description & ""
                .MoveNEXT
            Loop
            .Close
        End If
    End With
    Set rs = Nothing
    
    If bDontUpdateFromMaker = False Then
        If bReIssuance Then
            'Fill Issuable Items...do nothing for now...
'            If txtLotNo.Tag <> "" Then
'                Call FillList(cmbProcess.ID, cmbVender.ID, txtLotNo.Tag, txtLotNo.Text)
'            Else
'                Call FillList(cmbProcess.ID, cmbVender.ID)
'            End If
        End If
    End If
    
    Call ShowVenderHistory
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub ShowVenderHistory()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lTotalQty As Long
    Dim dTotalValue As Double
    With rs
        .Open "Select Top 10 * From VMakerIssItems WHERE VendID=" & cmbVender.ID & " Order By DT Desc,EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        
        LVMaker.ListItems.Clear
        LVBal.ListItems.Clear
        lVendBal = 0
        lTotalQty = 0
        dTotalValue = 0
        Do Until .EOF
            Set ITM = LVMaker.ListItems.add(, , LVMaker.ListItems.count + 1)
            ITM.SubItems(1) = !ProcessName & ""
            ITM.SubItems(2) = ![ItemCode] & ""
            ITM.SubItems(3) = ![ItemName] & ""
            ITM.SubItems(4) = Format(![DT] & "", "dd-MM-yy")
            ITM.SubItems(5) = Val(![IssQty] & "")
            lTotalQty = lTotalQty + Val(![IssQty] & "")
            ITM.SubItems(6) = Val(![Rate] & "")
            ITM.SubItems(7) = Val(!IssValue & "") 'Round(Val(![IssQty] & "") * Val(![Rate] & ""))
            dTotalValue = dTotalValue + Val(!IssValue & "")
            ITM.SubItems(8) = Format(![ReturnDT] & "", "dd-MM-yy")
            ITM.SubItems(9) = Val(!Priority & "")
            .MoveNEXT
        Loop
        .Close
        txtTotalQtyHistory = lTotalQty
        txtTotalValueHistory = dTotalValue
        'If cmbVender.ID <> 129 Then
            .Open "Select ProcessName,ItemCode,ItemName,Sum(RemQty) As RemQty From VMakerIssItems Where VendID=" & cmbVender.ID & "  AND RemQty>0 Group By ProcessName,ItemCode,ItemName Having Sum(RemQty)>0", con, adOpenForwardOnly, adLockReadOnly
            
            Do Until .EOF
                Set ITM = LVBal.ListItems.add(, , LVBal.ListItems.count + 1)
                ITM.SubItems(1) = !ProcessName & ""
                ITM.SubItems(2) = ![ItemCode] & ""
                ITM.SubItems(3) = ![ItemName] & ""
                ITM.SubItems(4) = Val(![RemQty] & "")
                lVendBal = lVendBal + Val(![RemQty] & "")
                .MoveNEXT
            Loop
            .Close
            txtTotal = lVendBal
        'End If
        
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbVender_Validate(cancel As Boolean)
    cmdPost.Default = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    con.Execute "Delete From PrintSession"
End Sub

Private Sub cmdCancel_M_Click()
    Pic_Iss_Multi.Visible = False
End Sub

Private Sub cmdOK_M_Click()
    Pic_Iss_Multi.Visible = False
End Sub


Private Sub cmdCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdClearAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdPost_Click()

    On Error GoTo err
'    If LV.SelectedItem.ForeColor = 9844630 Then
'        MsgBox "Can't Issue. Item Is Locked", vbInformation
'        Exit Sub
'    End If
        
    If Not cmbVender.MatchFound Then Exit Sub
    If cmbVender.ID = "" Or cmbVender.Text = "" Then Exit Sub
    
    
    If cmbReturnProcess.MatchFound = False Then
        MsgBox "Please Select Return Process.", vbInformation
        Exit Sub
    End If
    
'    If Val(cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1)) <= 0 Then
'        MsgBox "There is some Problem with Return Process.", vbInformation
'        Exit Sub
'    End If
    
    If Val(txtVendRate) <= 0 And cmbVender.ID <> iFactoryMakerID Then
        MsgBox "Invalid Rate!", vbInformation
        Exit Sub
    End If
    
'    If bReIssuance = False Then
'        If cmbIssEmp.MatchFound = False Then
'            MsgBox "Please Select Employee to Add.", vbInformation
'            Exit Sub
'        End If
'    End If
    'If cmbReturnProcess.MatchFound = False Then Exit Sub
    'cmbVender.Enabled = False
    
    'Check For Max Balance Limit For This Vender.......
    If cmbVender.ID <> 129 Then
        Dim rs As New ADODB.Recordset
        With rs
'            .Open "Select MaxLimit From Venders Where VendID=" & cmbVender.ID, con, adOpenForwardOnly, adLockReadOnly
'            If lVendBal >= Val(![MaxLimit] & "") Then
'                MsgBox "Sorry, The Selected Vender Has Reached His Max. Issuence Limit." & _
'                 vbNewLine & "Can't Post It For This Vender." & vbNewLine & "Vender Maximum Limit : " & Val(![MaxLimit] & "") & vbNewLine & "Vender Curr. Balance  : " & lVendBal, vbInformation
'
'                .Close
'                Exit Sub
'            End If
'            .Close
        End With
        Set rs = Nothing
    End If
    
    If RcvingLot Then
        PostForLots
        Exit Sub
    End If
    
    If Val(txtQty) < 1 Then Exit Sub
    
'    If bReIssuance = False Then
'        If strCompany = "IAA" OR strCompany = "Kami" Then
'            Dim lExtraQty As Long, lTotalQty As Long, lIssuedQty As Long
'            lExtraQty = (Val(LV.SelectedItem.SubItems(6)) / 100) * 5
'            lTotalQty = Val(LV.SelectedItem.SubItems(6)) + lExtraQty
'            lIssuedQty = Val(LV.SelectedItem.ListSubItems(6).Tag)
'
'            If Val(txtQty) > (lTotalQty - lIssuedQty) Then
'                MsgBox "Can't Issue more than 5% of the order qty.", vbInformation
'                Exit Sub
'            End If
'        End If
'    End If
    
    If bReIssuance Then
        If Val(txtQty) > Val(LV.SelectedItem.SubItems(5)) Then
            MsgBox "Invalid Qty", vbInformation
            Exit Sub
        End If
    Else
        If Not (cmbProcess.ID = 88 Or cmbProcess.ID = 141) Then
            If txtBatchNo = "" Then
                MsgBox "Please Enter Batch No.", vbInformation
                Exit Sub
            End If
        End If
    End If
    
    If GotIssID = 0 Then
        'Dim IssID As String
        'For i = 1 To LVIssued.ListItems.Count
        '    If Val(cmbVender.ID) = Val(LVIssued.ListItems(i).ListSubItems(2).Key) Then
        '        IssID = LVIssued.ListItems(i).ListSubItems(2)
        '        GoTo GotID:
        '    End If
        'Next i
    
        GotIssID = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,LEN(RecieptID)-10) AS INT)) FROM VMakerIssItems WHERE Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and Left(RecieptID,3)<>'HSR'").Fields(0).Value & "")
        GotIssID = GotIssID + 1
    Else
        GotIssID = GotIssID + 1
    End If
    'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
     Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
    IssID = "ISU-" & _
     Format(DTPicker2, "ddMMyy") & GotIssID
    
    Dim strBuyerCode As String
    strBuyerCode = con.Execute("SELECT CustCode FROM FCUstomerOrders WHERE OrderNo='" & LV.SelectedItem.SubItems(2) & "'").Fields(0).Value & ""
    
GotID:
        
    Dim ITM As ListItem
    With ITM
        'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
        'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
        Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
        ITM.Tag = LV.SelectedItem.Index
        'itm.ListSubItems.Add , cmbVender.ID & "'", cmbVender.Text
        ITM.ListSubItems.add(, , cmbVender.Tag).Tag = cmbVender.ID
        ITM.ListSubItems.add(, , cmbReturnProcess.Text).Tag = IssID & ""  'DTPicker2
        ITM.ListSubItems.add(, , strBuyerCode).Tag = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1) 'DTPicker2
        ITM.ListSubItems.add(, , LV.SelectedItem.SubItems(2)).Tag = cmbReturnProcess.List(cmbReturnProcess.ListIndex, 1)
        ITM.ListSubItems.add(, , LV.SelectedItem.SubItems(3)).Tag = txtSpecialInstructions.Text
        ITM.ListSubItems.add , , Format(DTReturn.Value, "dd-MMM-yyyy") 'LV.SelectedItem.SubItems(3)
        ITM.ListSubItems.add(, , Val(txtVendRate.Text)).Tag = cmbProcess.ID
        ITM.ListSubItems.add , , Val(txtQty) 'LV.SelectedItem.SubItems(3)
        ITM.ListSubItems.add , , GetValueofItem(cmbItems.Text)   'Val(txtQty) * Val(txtVendRate.Text)
        ITM.ListSubItems.add(, , LVIssued.ListItems.count).Tag = LV.SelectedItem.ListSubItems(2).Tag 'Assign From Rcvd_RefID to (10).tag
        ITM.ListSubItems.add , , ""
        ITM.ListSubItems.add , , txtBatchNo
        If bReIssuance = False Then
            ITM.ListSubItems(11).Tag = cmbIssEmp.ID
        End If
        'ITM.ListSubItems.Add , , Val(txtQty)
        If cmbProcess.Tag = "Lot No" Then
            ITM.ListSubItems(6).Tag = LV.SelectedItem.SubItems(3) & ""
        End If
        
        
        Dim bReWorkLot As Boolean, iRepairType As Integer
        
        With rs
            .Open "SELECT ReWorkLot,Repair_RefID FROM VendRcvdDetail WHERE LotNo='" & (txtLotNo) & "' AND NextProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                bReWorkLot = !ReWorklot
                iRepairType = Val(!Repair_RefID & "")
            Else
                bReWorkLot = False
                iRepairType = 0
            End If
            .Close
        End With
        Set rs = Nothing
        ITM.ListSubItems(8).Tag = bReWorkLot
        ITM.ListSubItems(9).Tag = iRepairType
    End With
    Call getTotalValue
    
    'MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtQty)
    LV.SelectedItem.SubItems(5) = Val(LV.SelectedItem.SubItems(5)) - Val(txtQty)
     
    txtRate.Visible = False
    
'    Dim i As Integer, lBatchNo As Long, lBatchNo_Max As Long
'
'    For i = 1 To LVIssued.ListItems.count
'
'        If LVIssued.ListItems(i).SubItems(5) = LV.SelectedItem.SubItems(3) Then
'            LVIssued.ListItems(i).ListSubItems(5).Tag = txtSpecialInstructions.Text
'        End If
'
'        lBatchNo = Right(LVIssued.ListItems(i).SubItems(12), Len(LVIssued.ListItems(i).SubItems(12)) - 1)
'
'        If lBatchNo > lBatchNo_Max Then
'            lBatchNo_Max = lBatchNo
'        End If
'
'    Next
'    txtBatchNo = "B" & (lBatchNo_Max + 1)
    
    'Removing Item From Combos
'    If LV.SelectedItem.SubItems(4) = 0 Then
'        Dim ItemIndex As Long
'        ItemIndex = cmbItems.ListIndex
'        If ItemIndex <> -1 Then
'
'            cmbItems.RemoveItem ItemIndex
'            cmbDesc.RemoveItem ItemIndex
'            cmbItems.Text = ""
'            cmbDesc.Text = ""
'            If cmbItems.ListCount >= ItemIndex And ItemIndex > 0 Then
'                cmbItems.ListIndex = ItemIndex
'            ElseIf cmbItems.ListCount > 0 Then
'                cmbItems.ListIndex = 0
'            End If
'        End If
'    End If
    'if cmbitems.
    'Call LV_ItemClick(LV.SelectedItem)
    
    txtQty = ""
    
    
    'Select The Next Item Automaticaly....
    If LV.SelectedItem.Index < LV.ListItems.count Then
        If Val(LV.SelectedItem.SubItems(4)) = 0 Then
            LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
        End If
        Call LV_ItemClick(LV.ListItems(LV.SelectedItem.Index))
    End If
        
    'cmbVender.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
        'End If
End Sub

Private Function GetValueofItem(strItemID As String) As Double

    On Error GoTo err
    If txtAssignedUnit = "Kgs" Then
        Dim dWeight As Double
        dWeight = GetSingleDoubleValue("ForgingWeight", "Items", " WHERE ItemID='" & strItemID & "'")
        GetValueofItem = ((dWeight * Val(txtQty)) / 1000) * Val(txtVendRate.Text)
    Else
        GetValueofItem = Val(txtQty) * Val(txtVendRate.Text)
    End If
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub getTotalValue()

    Dim i As Integer, lTotal As Long, lQty As Long
    For i = 1 To LVIssued.ListItems.count
        lTotal = lTotal + Val(LVIssued.ListItems(i).SubItems(9))
        lQty = lQty + Val(LVIssued.ListItems(i).SubItems(8))
    Next
    
    txtTotalValue = lTotal
    txtTotalQty = lQty
    
End Sub

Private Sub cmdSave_Click()

    If Saved Then
        If chkPrint.Value = vbChecked Then
            Call PrintIssuanceSlips
            DoEvents
        End If
        FormProcessID = cmbProcess.ID
        'Unload Me
        'Me.Show
        'cmbProcess.ID = cmbProcess.Tag
        'Call ShowIssForm(FormProcessID)
        'frmVendIssItemsByMaker.Show
        LV.ListItems.Clear
        cmbItems.Clear
        cmbVender.ClearVals
        
        LVIssued.ListItems.Clear
        'cmbProcess.ID = ""
        txtLotNo.SelStart = 0
        txtLotNo.SelLength = Len(txtLotNo.Text)
        If bReIssuance Then
            txtLotNo.SetFocus
        End If
        cmdPost.Default = False
        con.Execute "DELETE FROM PrintSession"
        Call cmbItems_Click
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        If strCompany <> "Banzai" Then
            If chkPrint.Value = vbChecked Then
                Call PrintIssuanceSlips
                DoEvents
            End If
        End If
'        If cmbProcess.ID <= 7 Then
'            If Not UserReqAuth Then Call PrintIssRcvRpts("Original", True)
'        End If
        Unload Me
        con.Execute "DELETE FROM PrintSession"
    End If
    
End Sub

Private Sub PrintIssuanceSlips()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
'    If bReIssuance Then
'        Set rpt = rptApp.OpenReport(RptPath & "\IssSlipGP.rpt")
'    Else
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
'    End If
    strSelection = "{VendIssued.EntryID}=" & lVI_EntryID_For_Print
    'strSelection = "ISNULL({PrintSession.RecieptNo})=False"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    con.Execute "Delete From VendAssItems Where EntryID=" & Val(LV.SelectedItem.key), iAffect
    If iAffect > 0 Then
        'Call cmbNick_Click
        cmdUnAssign.Visible = False
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Command1_Click()

End Sub

Private Sub cmdSkipCancel_Click()
    FraSkip.Visible = False
End Sub

Private Sub cmdSkipOk_Click()

    If cmbSkipProcess.MatchFound = False Then
        MsgBox "Please Select Process.", vbInformation
        Exit Sub
    End If
    
    Dim strItemID As String
    
    strItemID = LV.SelectedItem.SubItems(3)
    Dim iRowAffect As Integer
    
    If Trim(txtLotNo) <> "" Then
        con.Execute "UPDATE VendRcvdDetail SET NextProcessID=" & cmbSkipProcess.ID & " WHERE ItemCode='" & _
         strItemID & "' AND RcvdQty>IssQty AND NextProcessID=" & cmbProcess.ID & " AND LotNo='" & Trim(txtLotNo) & "'", iRowAffect
    Else
        'con.Execute "UPDATE VendRcvdDetail SET NextProcessID=" & cmbSkipProcess.ID & " WHERE ItemCode='" & _
         strItemID & "' AND RcvdQty>IssQty AND NextProcessID=" & cmbProcess.ID
    End If
    Call txtLotNo_KeyPress(13)
    FraSkip.Visible = False
    
End Sub

Private Sub cmdSOCancel_Click()
    FraStockOrder.Visible = False
End Sub

Private Sub cmdSOOK_Click()

    If cmbSOItems.MatchFound = False Then
        MsgBox "Please Select Article.", vbInformation
        Exit Sub
    ElseIf Val(txtSOQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    '-----------------------------------------
    'Check If Opening for Selected Article for Selected Process is feeded ?
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(ItemID)", "StockOrderOpening", " WHERE ItemID='" & cmbSOItems.ID & "' AND ProcessID=" & cmbProcess.ID)
'    If lCount > 0 Then
'        MsgBox "This Item has already Opening feeded at this Process.", vbInformation
'        Exit Sub
'    End If
    
    Call StartTrans(con)
    
    'Check Customer...
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CustCode)", "ForeignCustomers", " WHERE CustCode='Stock' AND Country='PK'")
    If lCount = 0 Then
        con.Execute "INSERT INTO ForeignCustomers(Custcode,Country) VALUES('Stock','PK')"
    End If
    
    'Now Check to see if Stock Order is feeded.
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FCustomerOrders", " WHERE OrderNo='Stock-OrderIss'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerOrders(CustCode,Country,OrderNo,DT) VALUES('Stock','PK','Stock-OrderIss','" & Date & "')"
    End If
    
    'Now Check If Article is Assigned to Customer
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CompItemID)", "FCustomerCatalog", " WHERE Custcode='Stock' AND Country='PK' AND CompItemID='" & cmbSOItems.ID & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerCatalog(Custcode,Country,ItemID,CompItemID) VALUES('Stock','PK','" & cmbSOItems.ID & "','" & cmbSOItems.ID & "')"
    End If
    'Now Check If Article is feeded against this Article
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FOrderItems", " WHERE OrderNo='Stock-OrderIss' AND CompItemCode='" & cmbSOItems.ID & "'")
    If lCount = 0 Then
    
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
            con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty,DeliveryDT,Stamps,Quality,Remarks) VALUES('Stock-OrderIss','" & cmbSOItems.ID & _
             "','" & cmbSOItems.ID & "'," & Val(txtSOQty) & ",'" & Date & "','','','')"
        Else
            con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty) VALUES('Stock-OrderIss','" & cmbSOItems.ID & _
             "','" & cmbSOItems.ID & "'," & Val(txtSOQty) & ")"
        End If
         
    Else
        con.Execute "UPDATE FOrderItems SET Qty=Qty+" & Val(txtSOQty) & " WHERE OrderNo='Stock-OrderIss' AND CompItemCode='" & cmbSOItems.ID & "'"
    End If
    
    'Now Check If Article is feeded in Stock Order
    lCount = 0
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        lCount = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening", " WHERE Shelf_RefID=0 AND ItemID='" & cmbSOItems.ID & "' AND ProcessID=" & Val(cmbProcess.ID))
    Else
        lCount = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening", " WHERE ItemID='" & cmbSOItems.ID & "' AND ProcessID=" & cmbProcess.ID)
    End If
    
    If lCount = 0 Then
        Dim strLocation As String, lQty As Long, lEntryID As Long
        strLocation = "": lQty = Val(txtSOQty)
        
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
            con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName,Shelf_RefID,Remarks) VALUES('" & cmbSOItems.ID & _
             "'," & cmbProcess.ID & "," & lQty & ",'" & strLocation & "','" & CurrentUserName & "','" & strComputerName & "',0,'')"
        Else
            con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName) VALUES('" & cmbSOItems.ID & _
             "'," & cmbProcess.ID & "," & lQty & ",'" & strLocation & "','" & CurrentUserName & "','" & strComputerName & "')"
        End If
        
        lEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
    Else
        lEntryID = lCount
        
        con.Execute "UPDATE StockOrderOpening SET Qty=Qty+" & Val(txtSOQty) & " WHERE EntryID=" & lEntryID
    End If
    
    Dim lRcvEntryID As Long, lPrevProcessID As Long
    
    
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) VALUES(" & _
     cmbVender.ID & ",'" & Date & "','',0," & cmbProcess.ID & ",0)"
    
    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
    
    lPrevProcessID = getPrevProcessID(cmbSOItems.ID, cmbProcess.ID)
    
    con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
     ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & cmbSOItems.ID & _
     "',''," & lQty & ",0,'0'," & cmbProcess.ID & ",'Stock-OrderIss','',0," & lPrevProcessID & _
     ",0," & lEntryID & ")"
     
    con.CommitTrans
    '-----------------------------------------
    
    FraStockOrder.Visible = False
    Call cmbVender_matched
    
    Exit Sub
    
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdUpdateRate_Click()

    If cmbVender.MatchFound = False Then
        MsgBox "No Maker Selected.", vbInformation
        Exit Sub
    ElseIf LV.SelectedItem Is Nothing Then
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    ElseIf cmbProcess.MatchFound = False Or cmbProcess.ID = "0" Then
        MsgBox "No Process Selected.", vbInformation
        Exit Sub
    End If
    
    con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtVendRate) & " WHERE VendID=" & cmbVender.ID & " AND ItemID='" & _
     LV.SelectedItem.SubItems(3) & "' AND ProcessID=" & cmbProcess.ID
     
End Sub

Private Sub DTPicker1_Change()
    Call DTPicker1_Click
End Sub

Private Sub DTPicker1_Click()
    On Error GoTo err
    
    
    With rsMain
    
        If DTPicker1.Value & "" = "" Then
            .Filter = adFilterNone
        Else
            .Filter = "RecieptID Like 'HS-" & Format(DTPicker1, "ddMMyy") & "-*'"
            '.Filter = "Left(RecieptID,9) ='HS-" & Format(DTPicker1, "ddMMyy") & "'"
        End If
                
        Dim isAdded As Boolean
        
        cmbItems.Clear
        cmbDesc.Clear
        LV.ListItems.Clear
        Do Until .EOF
            isAdded = False
            For i = 0 To cmbItems.ListCount - 1
                If Trim(cmbItems.List(i)) = Trim(![RecieptID] & "") Then
                    isAdded = True
                    Exit For
                End If
            Next i
            
            If Not isAdded Then
                cmbItems.AddItem ![RecieptID] & ""
                cmbDesc.AddItem ![RecieptID] & ""
            End If
            .MoveNEXT
        Loop
    End With
    
    
        
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub DTPicker2_Change()
    DTReturn.Value = DateAdd("d", 30, DTPicker2.Value)
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Exit Sub 'For Now
    
    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If

    On Error Resume Next
    LV.SelectedItem.EnsureVisible

End Sub


Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    
    If bReIssuance = True Then
        mnuSkipProcess.Visible = True
        mnuStockOrderIssuance.Visible = False
    Else
        mnuStockOrderIssuance.Visible = True
        mnuSkipProcess.Visible = False
    End If
    
    If strCompany = "Weldon Instruments" Or strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        Me.PopUpMenu mnuPop
    End If
    
End Sub

Private Sub LVIssued_DblClick()
    
    
    If LVIssued.ListItems.count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub
    If LV.ColumnHeaders(3).Text = "Lot No/Order No" And RcvingLot Then Exit Sub
    
    If CurrentUserName <> "Administrator" And iColNo = 8 Then Exit Sub
    With txtRate1
        cmdPost.Default = False
        .Move LVIssued.Left + LVIssued.ColumnHeaders(iColNo).Left, LVIssued.Top + LVIssued.SelectedItem.Top, LVIssued.ColumnHeaders(iColNo).Width
        .Text = LVIssued.SelectedItem.SubItems(iColNo - 1) 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
        
    End With

End Sub

Private Sub LVIssued_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If KeyCode = 46 Then
    
'        If RcvingLot Then
'            LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
'            Exit Sub
'        End If
        
        'MyCol(LVIssued.SelectedItem.Tag).IssueableQty = MyCol(LVIssued.SelectedItem.Tag).IssueableQty + Val(LVIssued.SelectedItem.SubItems(7))
        'LV.ListItems(LVIssued.SelectedItem.Tag).ListSubItems(4) = LV.ListItems(LVIssued.SelectedItem.Tag).ListSubItems(4).Tag
        'LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(5) = MyCol(LVIssued.SelectedItem.Tag).IssueableQty
        LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
        Call LV_ItemClick(LV.SelectedItem)
        
        'Reset The IssID's
        Dim TempID As Long
        TempID = 0
        If LVIssued.ListItems.count > 0 Then
            For i = LVIssued.SelectedItem.Index To LVIssued.ListItems.count
                TempID = Val(Right(LVIssued.ListItems(i).SubItems(2), 3)) - 1
                If TempID > 1 Then
                    LVIssued.ListItems(i).SubItems(2) = "ISU-" & _
                     Format(ServerDate, "ddMMyy") & "-" & Format(TempID, "000")
                End If
            Next i
        Else
            GotIssID = 0
            
        End If
        If GotIssID > 0 Then GotIssID = GotIssID - 1
        Call getTotalValue
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LVIssued_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > LVIssued.ColumnHeaders(10).Left Then
        iColNo = 10
    ElseIf x > LVIssued.ColumnHeaders(9).Left Then
        iColNo = 9
    Else
        iColNo = 8
    End If
End Sub

Private Sub LVIssued_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If LVIssued.ListItems.count = 0 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Multiple Processes")
'    If bReIssuance Then
'        strRet = ShowPopUpMenu1("Multiple Processes")
'    Else
'        'strRet = ShowPopUpMenu1("Steel Assignment")
'        Exit Sub
'    End If
    If strRet = "Multiple Processes" Then
        Dim lVendID As Long, lProcessID As Long, strItemID As String, lSNo As Long, dRate As Double
        Dim ITM As ListItem
        lVendID = cmbVender.ID
        lProcessID = cmbProcess.ID
        strItemID = LVIssued.SelectedItem.SubItems(5)
        lSNo = GetSingleLongValue("SNo", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND ProcessID=" & lProcessID)
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT * FROM VItemProcesses WHERE IsExist='" & strItemID & "' AND ItemSNo>" & lSNo & " ORDER BY ItemSNo", con, adOpenForwardOnly, adLockReadOnly
            LVIssd_Multi.ListItems.Clear
            Do Until .EOF
                Set ITM = LVIssd_Multi.ListItems.add(, , !Description & "")
                ITM.Tag = Val(!ProcessID & "")
                dRate = GetSingleDoubleValue("Rate", "VVendAssItems", " WHERE VendID=" & cmbVender.ID & " AND ProcessID=" & Val(!ProcessID & "") & " AND ItemID='" & strItemID & "'")
                If dRate = 0 Then dRate = 1
                ITM.ListSubItems.add , , dRate
                .MoveNEXT
            Loop
            .Close
        End With
        Set rs = Nothing
        
        Pic_Iss_Multi.Visible = True
    ElseIf strRet = "Steel Assignment" Then
'        cmbRM_Steel.Clear
'        Call AddToCombo(cmbRM_Steel, "RMID1 + ' ' + RMName", "RM", " WHERE GroupID IN(SELECT RMG_RefID FROM RMGroupsForBatchwiseIssuance)", False, "RMID")
'        Pic_Assign_Steel.Visible = True
        
    End If
    
End Sub

Private Sub LvMaker_DblClick()
    
    Exit Sub
    If LVMaker.ListItems.count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub

    
    With txtRate
        .Move LVMaker.Left + LVMaker.ColumnHeaders(4).Left, LVMaker.Top + LVMaker.SelectedItem.Top, LVMaker.ColumnHeaders(4).Width
        '.Text = MyCol(LV.SelectedItem.SubItems(2) & "|" & LV.SelectedItem.SubItems(3)).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Text = MyCol(LV.SelectedItem.key).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        
        If LV.ColumnHeaders(3).Text = "Lot No" Then
            If RcvingLot Then
                .Locked = True
            Else
                .Locked = False
            End If
        Else
            .Locked = False
        End If
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    mnuPop.Visible = False
    'Call AddToCombo(cmbOrderNo, "OrderNo", "FCustomerOrders", " Order By OrderNo")
    
    bReIssuance = False
    
    'cmbCust.ListHeight = 2600
    'cmbCustName.ListHeight = 2600
    'cmbOrders.ListHeight = 2600
    
    'cmbProcess.Enabled = True
    bFromMaker = False
    cmbProcess.ListHeight = 6000
     
    DTPicker2.Value = ServerDate
    Call DTPicker2_Change
    DTPicker1.Value = ServerDate
    DTPicker1.Value = ""
    DTReturn.Value = DateAdd("d", 30, DTPicker2.Value)
    SSTab1.Tab = 2
    IsLotProcess = False
    cmbVender.ListHeight = 2000
    GotIssID = 0
    
    cmbSkipProcess.ListHeight = 1800
        
    
    lCheckerProcessNo = Val(getGeneralDataValue("CheckerProcessNo"))
    
    cmbIssEmp.ListHeight = 3000
    
    
    
    Call AddToCombo(cmbSteel, "SteelType", "SteelTypes", , , "SteelID")
    
    'txtBatchNo = getNextBatchNo
    
    Me.Hide
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub LockProc(ProcID As Long)
    
    
    cmbProcess.Tag = ProcID
    
    'cmbProcess.Enabled = False
    'Me.Refresh
    
    cmbProcess.ID = ProcID
    Me.Show
    
End Sub
Public Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select *  from FCustomerCatalog Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![CustCode] & ![Country] & ![ITemID], ![ITemID] & "")
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = ![Unit] & ""
            ITM.SubItems(3) = ![FOB] & ""
            ITM.SubItems(4) = ![CIFSea] & ""
            ITM.SubItems(5) = ![CIFAir] & ""
            ITM.SubItems(6) = ![CnFSea] & ""
            ITM.SubItems(7) = ![CnFAir] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_Resize()
      
    Dim TitleBottom As Long
    If bReIssuance Then
        LV.Move Label2.Left, lblSteel.Top, LV.Width, LV.Height + 600
    End If
    
    Exit Sub
    On Error Resume Next
    
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

    TitleBottom = DTPicker1.Top + DTPicker1.Height + 50 '+ fTitle.Height
 
    Label3.Move Label2.Left, LV.Top + LV.Height + 50
'    cmbVender.Move Label3.Left, Label3.Top + Label3.Height - 15
    
    
    DTPicker2.Move cmbVender.Left + cmbVender.Width + 20, Label3.Top + Label3.Height - 15
    txtQty.Move DTPicker2.Left + DTPicker2.Width + 10, Label3.Top + Label3.Height - 15
    cmdPost.Move txtQty.Left + txtQty.Width + 10, Label3.Top
    SSTab1.Move LV.Left, cmbVender.Top + cmbVender.Height + 50
    
End Sub

Private Sub FillVendList()

    On Error GoTo err
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim VendIDList As String
    VendIDList = ""
    For i = 1 To LVIssued.ListItems.count
        With LVIssued.ListItems(i)
            'If LV.SelectedItem.key = .Tag Then
            '    VendIDList = VendIDList & Val(.ListSubItems(1).key) & ","
            'End If
            If LV.SelectedItem.ListSubItems(1) = .SubItems(3) Then
                VendIDList = VendIDList & Val(.ListSubItems(1).Tag) & ","
            End If
        End With
    Next i
    
    If VendIDList <> "" Then
        VendIDList = "(" & Left(VendIDList, Len(VendIDList) - 1) & ")"
    Else
        VendIDList = "(0)"
    End If
    
    'With rs
    '    .Open "Select VendID,Rate,VenderName From VVendAssItems Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and PhaseID=" & cmbProcess.ID & " and VendID Not In" & VendIDList, con, adOpenForwardOnly, adLockReadOnly
    '    LVMaker.ListItems.Clear
    '    Do Until .EOF
    '
    '        'Set ITM = LVMaker.ListItems.Add(, ![VendID] & "'", LVMaker.ListItems.Count + 1)
    '        Set ITM = LVMaker.ListItems.Add(, , LVMaker.ListItems.Count + 1)
    '        ITM.Tag = Val(MyCol(Item.Index).IssueableQty & "")
    '        ITM.ListSubItems.Add , ![VendID] & "'", .Fields(2) & ""
    '        ITM.ListSubItems.Add , , Val(![Rate] & "")
    '
    '        ITM.ListSubItems.Add(, , 0).Tag = LV.SelectedItem.SubItems(2) & ""
    '        .MoveNext
    '    Loop
    '    .Close
    'End With
    Call cmbVender.ClearVals
    'If RcvingLot Then
    '    Call cmbVender.AddVals(con, "VenderName", "VVendAssItems", "VendID", " Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1) ")
    'Else
    If HicoVisible And RcvingLot Then
        Call cmbVender.AddVals(con, "Distinct VenderName", "VMakerAssItems", "VendID", " Where (ItemID='" & LV.SelectedItem.SubItems(2) & "' and PhaseID=" & cmbProcess.ID & "  AND Active=1 and VendID Not In" & VendIDList & " and VendType=1 Or (PhaseID=-1))")
    Else
        Call cmbVender.AddVals(con, "Distinct VenderName", "VMakerAssItems", "VendID", " Where (ItemID='" & LV.SelectedItem.SubItems(1) & "' AND Active=1 and (PhaseID=" & cmbProcess.ID & ") and VendID Not In" & VendIDList & " ) OR VendID=129 ")
    End If
    'End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    txtQty = LV.SelectedItem.ListSubItems(5)
        
    Dim i As Integer
    For i = 1 To LVIssued.ListItems.count
        If LV.SelectedItem.SubItems(3) = LVIssued.ListItems(i).SubItems(5) Then
            txtSpecialInstructions = LVIssued.ListItems(i).ListSubItems(5).Tag
            Exit For
        End If
    Next
    
    Dim bReWorkLot As Boolean, iRepairType As Integer
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ReWorkLot,Repair_RefID FROM VendRcvdDetail WHERE LotNo='" & (txtLotNo) & "' AND NextProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            bReWorkLot = !ReWorklot
            iRepairType = Val(!Repair_RefID & "")
        End If
        .Close
    End With
    
    
'    If bReWorkLot = False Then
'
'        Set rs = con.Execute("EXEC NextProcessesForIssuance_SP " & Val(cmbProcess.ID) & "," & Val(cmbVender.ID) & ",'" & Item.SubItems(3) & "'")
'
'        cmbReturnProcess.Clear
'        Do Until rs.EOF
'            cmbReturnProcess.AddItem rs!Description & ""
'            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 1) = Val(rs!ProcessID & "")
'            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 2) = Val(rs!Rate & "")
'            cmbReturnProcess.List(cmbReturnProcess.ListCount - 1, 3) = rs!Unit & ""
'            rs.MoveNEXT
'        Loop
'    Else
'        cmbReturnProcess.Clear
'        cmbReturnProcess.AddItem cmbProcess.Text
'        cmbReturnProcess.List(0, 1) = cmbProcess.ID
'        Dim strUnit As String, dRate As Double
'
'        With rs
'            .Open "SELECT Unit,Rate FROM VendAssItems WHERE VendID=" & Val(cmbVender.ID) & " AND ItemID='" & Item.SubItems(3) & "' AND ProcessID=" & Val(cmbProcess.ID), con, adOpenForwardOnly, adLockReadOnly
'            If .EOF = False Then
'                strUnit = !Unit & ""
'                dRate = Val(!Rate & "")
'            Else
'                strUnit = ""
'                dRate = 0
'            End If
'            .Close
'        End With
'        cmbReturnProcess.List(0, 2) = dRate
''        cmbReturnProcess.List(0, 3) = strUnit
'    End If
    cmbReturnProcess.AddItem cmbProcess.Text 'rs!Description & ""
    cmbReturnProcess.List(0, 1) = cmbProcess.ID 'Val(rs!ProcessID & "")
    Dim dRateToUse As Double
    
    dRateToUse = GetSingleDoubleValue("Rate", "VendAssItems", " WHERE ItemID='" & Item.SubItems(3) & "' AND ProcessID=" & Val(cmbProcess.ID) & " AND VendID=" & Val(cmbVender.ID))
    dRateToUse = IIf(dRateToUse = 0, 1, dRateToUse)
    cmbReturnProcess.List(0, 2) = dRateToUse
    cmbReturnProcess.List(0, 3) = "Pcs"
    txtVendRate = dRateToUse
    
    If cmbReturnProcess.ListCount > 0 Then cmbReturnProcess.ListIndex = 0
    FraSkip.Visible = False
    FraStockOrder.Visible = False
    
    Set rs = Nothing
    
    If bReIssuance = False Then
        cmbIssEmp.ClearVals
        Dim p_strEmpID As String
        p_strEmpID = GetSingleStringValue("EmpID_Purchaser", "ItemGroups INNER JOIN Items ON Items.GroupID=ItemGroups.ID", " WHERE ItemID='" & Item.SubItems(3) & "'")
        cmbIssEmp.AddVals con, "'{' + EmpID + '} ' + Name + ' {' + Designation + '}'", "Employees", "EmpID", " WHERE EmpID='" & p_strEmpID & "'"
        cmbIssEmp.ID = p_strEmpID
        
        
        
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub mnuSkipProcess_Click()
    ''
    If LV.ListItems.count = 0 Then Exit Sub
    Dim lSNo As Long, strItemID As String
    
    strItemID = LV.SelectedItem.SubItems(3)
    
    Dim lCount As Long, bRepairLot As Boolean
    lCount = GetSingleLongValue("COUNT(*)", "VendRcvdDetail", " WHERE LotNo='" & txtLotNo & "' AND ReWorkLot=1")
    
    If lCount > 0 Then
        bRepairLot = True
    Else
        bRepairLot = False
    End If
    
    If bRepairLot Then
        Dim lRepairType As Long
        lRepairType = GetSingleLongValue("Repair_RefID", "VendRcvdDetail", " WHERE LotNo='" & txtLotNo & "'")
        
        lSNo = GetSingleLongValue("SeqNo", "RepairTypeProcesses", " WHERE Repair_RefID=" & lRepairType & " AND ProcessID=" & cmbProcess.ID)
        
        cmbSkipProcess.ClearVals
        cmbSkipProcess.AddVals con, "Description", "VRepairTypeProcesses", "ProcessID", " WHERE EntryID=" & lRepairType & " AND SeqNo>" & lSNo & " ORDER BY SeqNo"
    Else
        lSNo = GetSingleLongValue("SNo", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND ProcessID=" & cmbProcess.ID)
        
        Dim lNextBillingProccessID As Long, lNBPSNo As Long, lPrevBillingProcessID As Long, lPBPSNo As Long, lSNoToCheck As Long
        lNBPSNo = GetSingleLongValue("MIN(ItemSNo)", "VItemProcesses", " WHERE IsExist='" & strItemID & "' AND ItemSNo>=" & lSNo & " AND BillingProcessID<>0")
        
        lPBPSNo = GetSingleLongValue("MAX(ItemSNo)", "VItemProcesses", " WHERE IsExist='" & strItemID & "' AND ItemSNo<" & lSNo & " AND BillingProcessID<>0")
        'lNextBillingProccessID = GetSingleLongValue("ItemSNo", "VItemProcesses", " WHERE IsExist='" & strItemID & "' AND ItemSNo=" & lBPSNo)
        If lPBPSNo > lSNo Then
            lSNoToCheck = lPBPSNo
        Else
            lSNoToCheck = lNBPSNo
        End If
        
        cmbSkipProcess.ClearVals
        'cmbSkipProcess.AddVals con, "Description", "Processes", "ProcessID", " WHERE ProcessID IN(SELECT ProcessID FROM ItemProcesses WHERE ItemID='" & strItemID & "' AND SNo>" & lSNo & ") ORDER BY SNo"
        cmbSkipProcess.AddVals con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & strItemID & "' AND ItemSNo>" & lSNo & " ORDER BY ItemSNo"
'        If lSNoToCheck > 0 Then
'            cmbSkipProcess.AddVals con, "Description", "Processes", "ProcessID", " WHERE ProcessID IN(SELECT ProcessID FROM ItemProcesses WHERE ItemID='" & strItemID & "' AND SNo>" & lSNo & " AND SNo<=" & lNBPSNo & ") ORDER BY SNo"
'        Else
'            cmbSkipProcess.AddVals con, "Description", "Processes", "ProcessID", " WHERE ProcessID IN(SELECT ProcessID FROM ItemProcesses WHERE ItemID='" & strItemID & "' AND SNo>" & lSNo & ") ORDER BY SNo"
'        End If
    End If
    
    FraSkip.Visible = True
    
End Sub

Private Sub mnuStockOrderIssuance_Click()
        ''
     
    cmbSOItems.ClearVals
    Dim strCondition As String
    strCondition = " WHERE InActive=0 AND ItemID IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & Val(cmbVender.ID) & " AND ProcessID=" & Val(cmbProcess.ID) & ") ORDER BY ItemName"
    Call cmbSOItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", strCondition)
    
    FraStockOrder.Visible = True
    
End Sub

Private Sub txtLotNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtLotNo = (txtLotNo)
        txtLotNo.Tag = (txtLotNo)
        'First Find out Process ID & Item Code of given LotNo.
        Dim rs As New ADODB.Recordset, bRecFound As Boolean
        Dim strItemCode As String, iProcessID As Integer, lEntryID As Long, lCurrentProcessID As Long
        With rs
            .Open "SELECT TOP 1 EntryID,ItemCode,NextProcessID,ProcessID FROM VendRcvdDetail WHERE LotNo='" & txtLotNo & "' ORDER BY EntryID DESC", con_ServerSide, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                bRecFound = True
                lEntryID = Val(!EntryID & "")
                lCurrentProcessID = Val(!ProcessID & "")
                If InspectionDataEntryCheck(lEntryID, lCurrentProcessID) = False Then
                    MsgBox "Inspection Data Entry is not Complete.", vbInformation
                    Exit Sub
                End If
                strItemCode = !ItemCode & ""
                iProcessID = Val(!NextProcessID & "")
            Else
                bRecFound = False
            End If
            .Close
            
        End With
        Set rs = Nothing
        
        Dim lIssCount As Long
        lIssCount = GetSingleLongValue("COUNT(Rcvd_RefID)", "VendIssdDetail", "WHERE Rcvd_RefID=" & lEntryID)
        If lIssCount > 0 Then
            txtLotNo.Tag = ""
            MsgBox "Already Issued.", vbInformation
            Exit Sub
        End If
        
        If bRecFound Then
            cmbProcess.ClearVals
            cmbProcess.AddVals con_ServerSide, "Description", "Processes", "ProcessID", " Order By SNo"
            cmbProcess.ID = iProcessID
            txtLotNo.Tag = strItemCode
            Call FillList(iProcessID, 0, strItemCode, (txtLotNo))
        Else
            txtLotNo.Tag = ""
        End If
    End If
    
End Sub

Private Sub txtQty_GotFocus()


    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    

End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        
        Dim IssID As String
        For i = 1 To LVIssued.ListItems.count
            If Val(LVMaker.SelectedItem.ListSubItems(1).key) = Val(LVIssued.ListItems(i).ListSubItems(2).key) Then
                IssID = LVIssued.ListItems(i).ListSubItems(2)
                GoTo GotID:
            End If
        Next i
        
        IssID = Val(con.Execute("Select Max(right(RecieptID,3)) From VendIssued Where DT='" & Format(ServerDate, "dd-MMM-yyyy") & "'").Fields(0).Value & "")
        
        'IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).Key) & "-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
         IssID = "ISU-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
GotID:
        
        Dim ITM As ListItem
        With ITM
            'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
            'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
            Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
            ITM.Tag = LV.SelectedItem.key
            ITM.ListSubItems.add , LVMaker.SelectedItem.ListSubItems(1).key, LVMaker.SelectedItem.SubItems(1)
            ITM.ListSubItems.add , , IssID & ""
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(2)
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(3)
            ITM.ListSubItems.add , , LV.SelectedItem.SubItems(4)
            ITM.ListSubItems.add , , LVMaker.SelectedItem.ListSubItems(2)
            ITM.ListSubItems.add , , Val(txtRate)
            If cmbProcess.Tag = "Lot No" Then
                ITM.ListSubItems(6).Tag = LV.SelectedItem.SubItems(2) & ""
            End If
        End With
        
        'MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtRate)
        'LV.SelectedItem.ListSubItems(5) = Val(LV.SelectedItem.ListSubItems(5)) - Val(txtRate)
        LV.SelectedItem.SubItems(5) = Val(LV.SelectedItem.SubItems(5)) - Val(txtRate)
        txtRate.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
    ElseIf KeyAscii = 27 Then
        txtRate.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    If Val(txtRcvings) <= 0 Then
        MsgBox "Please Enter Max. Rcvings.", vbInformation
        Exit Function
    End If
            
    If cmbProcess.ID = lCheckerProcessNo And cmbChecker.MatchFound = False Then
        MsgBox "Please Check Checker.", vbInformation
        Exit Function
    End If
    
    'IF Process is PTC Closed for Wrangler
    If strCompany = "IAA" Or strCompany = "Kami" Then
        If cmbProcess.ID = 206 Then
            If getDBPassword("Issuance_ClosePTC") = False Then
                Exit Function
            End If
        End If
    End If
    
    Dim strIssEmpID As String
    If cmbIssEmp.MatchFound = False Then
        strIssEmpID = ""
    Else
        strIssEmpID = cmbIssEmp.ID
    End If
    
    Dim iSteelType As Integer
    If cmbSteel.MatchFound = False Then
        iSteelType = 0
    Else
        iSteelType = cmbSteel.List(cmbSteel.ListIndex, 1)
    End If
    
    Dim lLotNo As String
    Dim SaveIt As Boolean, SaveItReturnDTs As Boolean
    
    Dim VendID As New Collection
    Dim VIssID As New Collection
     
'    Dim strVendnItem() As String
'    Dim iVendnItemCount As Integer
    
    Dim myPOSavedCol As New Collection, myPOSavedColWithOrder As New Collection
    
    Dim cmd As New ADODB.Command
    Dim TempIndex As Integer
    TempIndex = 1
    
    Dim lRefID As Long
    
    Dim IssLot As Boolean
    Dim iAuth As Integer 'Authorize
    
    IssLot = RcvingLot 'RcvLot(cmbProcess.ID)
    
'    With cmd
'        .CommandType = adCmdStoredProc
'        .CommandText = "SP_ReqAuthForIss"
'        Set .ActiveConnection = con
'        .Parameters("@UserID").value = UserID
'        .Parameters("@ProcessID").value = cmbProcess.ID
'        .Execute
'        UserReqAuth = Val(.Parameters("@ReqAuth").value & "")
'    End With
    UserReqAuth = 0
    
    Dim strMasterPONo As String
    If bReIssuance = False Then
        strMasterPONo = "'" & GetNextMasterPONo & "'"
    Else
        strMasterPONo = "NULL"
    End If
    
    Dim lRecAffect As Long
    Dim lVID_EntryID As Long
    Dim iAuthorized As Integer
    If iFactoryMakerID = cmbVender.ID Then
        iAuthorized = 1
    Else
        iAuthorized = 1
    End If
    
    If bReIssuance Then
        If strCompany = "Banzai" Then
            iAuthorized = 1
        End If
    End If
    
    Call StartTrans(con)
    
    For i = 1 To LVIssued.ListItems.count
        With LVIssued.ListItems(i)
            SaveIt = True
'            For j = 1 To VendID.count
'                If Val(LVIssued.ListItems(i).ListSubItems(1).Tag) = VendID(j) Then  'Check Vender
'                    'It Means That It's Been Saved Already.
'                        SaveIt = False
'                        Exit For
'                End If
'            Next j
'            If i = 1 Then
'                iVendnItemCount = 0
'            Else
'                iVendnItemCount = UBound(strVendnItem)
'            End If
'            For j = 0 To iVendnItemCount - 1
'                If strVendnItem(j, 0) = Val(LVIssued.ListItems(i).ListSubItems(1).Tag) And strVendnItem(j, 1) = .SubItems(5) Then
'                    SaveIt = False
'                    Exit For
'                End If
'            Next
            Dim myPOObjTemp As clsPOEntry
            For j = 1 To myPOSavedCol.count
                Set myPOObjTemp = myPOSavedCol(j)
                If myPOObjTemp.iIssueProcess = Val(.ListSubItems(7).Tag) And myPOObjTemp.iReturnProcess = Val(.ListSubItems(3).Tag) And myPOObjTemp.iVendID = Val(.ListSubItems(1).Tag) And myPOObjTemp.strItemCode = .SubItems(5) Then
                    SaveIt = False
                    lRefID = myPOObjTemp.lRefID
                    Set myPOObjTemp = Nothing
                    Exit For
                End If
                Set myPOObjTemp = Nothing
            Next
            
            If SaveIt Then
                Dim lMakerIssNo As Long, strMakerIssNo As String
                lMakerIssNo = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-12) As Int)) From VendIssued Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "'").Fields(0).Value & "")
                lMakerIssNo = lMakerIssNo + 1
                IssID = "ISU-" & Format(DTPicker2, "ddMMyy") & GotIssID
                strMakerIssNo = "M-ISU-" & Format(DTPicker2, "ddMMyy") & lMakerIssNo
                
                con.Execute "INSERT INTO VendIssued(VendID,DT,RecieptID,UserID,ProcessID,ItemID," & _
                "UserName,MachineName,SpecialInstructions,ExcessQtyPercentage," & _
                "MaximumRcvingsAgainstPO,MasterPONo,SampleProvided,DrawingProvided,ForgingProvided,SteelProvided,Authorized,IssEmpID,SteelType_RefID) VALUES(" & Val(LVIssued.ListItems(i).ListSubItems(1).Tag) & ",'" & _
                 DTPicker2.Value & "','" & strMakerIssNo & "'," & UserID & "," & _
                 cmbProcess.ID & ",'" & .SubItems(5) & "','" & CurrentUserName & "','" & _
                 strComputerName & "','" & .ListSubItems(5).Tag & "'," & Val(txtExcess) & "," & _
                 Val(txtRcvings) & "," & strMasterPONo & "," & chkSampleProvided.Value & "," & chkDrawingProvided.Value & "," & _
                 chkForgingProvided.Value & "," & chkSteelProvided.Value & "," & iAuthorized & ",'" & strIssEmpID & "'," & iSteelType & ")"
                 
                'lRefID = Val(con.Execute("SELECT Max(EntryID) FROM VendIssued").Fields(0).Value & "")
                lRefID = GetSingleLongValue("MAX(EntryID)", "VendIssued", " WHERE MachineName='" & strComputerName & "'")
                p_lEntryID = lRefID
                lVI_EntryID_For_Print = p_lEntryID
                
                If cmbVender.ID = iFactoryMakerID And (strCompany = "Weldon Instruments" Or strCompany = "QEL" Or strCompany = "Banzai") Then
                    con.Execute "UPDATE VendIssued SET Authorized=1 WHERE EntryID=" & lRefID
                End If
                
                'Add to Collection After Assigning to Object
                Dim myPOObj As New clsPOEntry
                myPOObj.iIssueProcess = Val(.ListSubItems(7).Tag)
                myPOObj.iReturnProcess = Val(.ListSubItems(3).Tag)
                myPOObj.iVendID = Val(.ListSubItems(1).Tag)
                myPOObj.strItemCode = .SubItems(5)
                myPOObj.lRefID = lRefID
                myPOSavedCol.add myPOObj
                Set myPOObj = Nothing
                
'                ReDim Preserve strVendnItem(iVendnItemCount + 1, 3)
'                ReDim Preserve strVendnItem(iVendnItemCount + 1)
'                ReDim Preserve strVendnItem(iVendnItemCount)
'                strVendnItem(iVendnItemCount, 0) = Val(LVIssued.ListItems(i).ListSubItems(1).Tag)
'                strVendnItem(iVendnItemCount, 1) = .SubItems(5)
'                strVendnItem(iVendnItemCount, 2) = lRefID
                
                VendID.add Val(LVIssued.ListItems(i).ListSubItems(1).Tag)
                VIssID.add lRefID
            Else
                'lRefID = VIssID(j)
                'lRefID = 'strVendnItem(j, 2)
            End If
        
            '1. Insert Into VendIssuedItems
            
            If cmbProcess.Tag = "Lot No" And IssLot Then
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_UpdateForIARLots"
                    
                    .Parameters("@ProcessID").Value = cmbProcess.ID
                    .Parameters("@Qty").Value = Val(LVIssued.ListItems(i).SubItems(6))
                    .Parameters("@LotNo").Value = LVIssued.ListItems(i).ListSubItems(6).Tag
                    .Parameters("@RefID").Value = lRefID
                    .Parameters("@Rate").Value = Val(LVIssued.ListItems(i).SubItems(5))
                    .Parameters("@RecID").Value = LVIssued.ListItems(i).SubItems(2)
                    .Parameters("@ReqAuth").Value = UserReqAuth
                    
                    .Execute
                End With
                
            Else
                If bReIssuance Then
                    lLotNo = GetSingleStringValue("LotNo", "VendRcvdDetail", " WHERE EntryID=" & Val(.ListSubItems(10).Tag))
                Else
                    If strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                        'IF QEL don't Assign Lot No. @ Issuance instead Lot No. will be Assigned upon Very first Receiving.
                        lLotNo = 0
                    Else
                        lLotNo = GetSingleLongValue("MAX(LotNo)", "VendIssdDetail") + 1
                    End If
                End If
                
                SaveItReturnDTs = True
                Dim myPOObjTemp1 As clsPOEntryWithOrder
                For j = 1 To myPOSavedColWithOrder.count
                    Set myPOObjTemp1 = myPOSavedColWithOrder(j)
                    If myPOObjTemp1.iIssueProcess = Val(.ListSubItems(7).Tag) And myPOObjTemp1.iReturnProcess = Val(.ListSubItems(3).Tag) And myPOObjTemp1.iVendID = Val(.ListSubItems(1).Tag) And myPOObjTemp1.strItemCode = .SubItems(5) And myPOObjTemp1.strOrderNo = .SubItems(4) Then
                        SaveItReturnDTs = False
                        lVID_EntryID = myPOObjTemp1.lRefID
                        Set myPOObjTemp1 = Nothing
                        Exit For
                    End If
                    Set myPOObjTemp1 = Nothing
                Next
                
                If SaveItReturnDTs = False Then
                    con.Execute "UPDATE VendIssdDetail SET IssQty=IssQty+" & Val(.SubItems(8)) & " WHERE EntryID=" & lVID_EntryID
                Else
                    con.Execute "INSERT INTO VendIssdDetail(RefID,RecieptID,ItemCode" & _
                     ",Rate,IssQty,ReqAuth,OrderNo,RcvProcessID,ReturnDT,Priority,Rcvd_RefID,LotNo,ReWorkLot,Repair_RefID,Batch_No,CountedBy) VALUES(" & lRefID & ",'" & _
                     .ListSubItems(2).Tag & "','" & .SubItems(5) & "'," & _
                     Val(.SubItems(7)) & "," & Val(.SubItems(8)) & "," & UserReqAuth & ",'" & _
                     .SubItems(4) & "'," & .ListSubItems(3).Tag & ",'" & .SubItems(6) & "'," & .SubItems(10) & "," & _
                     Val(.ListSubItems(10).Tag) & ",'" & lLotNo & "'," & Abs(.ListSubItems(8).Tag) & "," & Val(.ListSubItems(9).Tag) & ",'" & .SubItems(12) & "','" & cmbCountedBy.Text & "')"
                    
                    lVID_EntryID = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail")
                    
                    If bReIssuance = False Then
                        con.Execute "UPDATE VendIssdDetail SET VID_EmpID='" & .ListSubItems(11).Tag & "' WHERE EntryID=" & lVID_EntryID
                    End If
                    'Add to Collection After Assigning to Object
                    Dim myPOObj1 As New clsPOEntryWithOrder
                    myPOObj1.iIssueProcess = Val(.ListSubItems(7).Tag)
                    myPOObj1.iReturnProcess = Val(.ListSubItems(3).Tag)
                    myPOObj1.iVendID = Val(.ListSubItems(1).Tag)
                    myPOObj1.strItemCode = .SubItems(5)
                    myPOObj1.strOrderNo = .SubItems(4)
                    myPOObj1.lRefID = lVID_EntryID
                    myPOSavedColWithOrder.add myPOObj1
                    Set myPOObj1 = Nothing
                    
                    Dim lCount_MAI As Long
                    lCount_MAI = GetSingleLongValue("COUNT(*)", "VendAssItems", " WHERE VendID=" & cmbVender.ID & " AND ProcessID=" & Val(.ListSubItems(7).Tag) & " AND ItemID='" & .SubItems(5) & "'")
                    If lCount_MAI = 0 Then
                        Dim dRate_MAI As Double
                        dRate_MAI = Val(.SubItems(7))
                        If dRate_MAI = 0 Then dRate_MAI = 1
                        con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID,Rate,Unit,Remarks) VALUES(" & cmbVender.ID & "," & Val(.ListSubItems(7).Tag) & ",'" & .SubItems(5) & "'," & dRate_MAI & ",'Pcs','Auto')"
                    End If
                End If
                
                con.Execute "INSERT INTO VendIssdDetail_ReturnDTs(VIS_RefID,IssQty,RcvQty,ReturnDT) VALUES(" & _
                 lVID_EntryID & "," & Val(.SubItems(8)) & ",0,'" & .SubItems(6) & "')"
                
                If cmbProcess.ID = lCheckerProcessNo Then
                    con.Execute "INSERT INTO VendIssdDetail_MoreDetails(VID_RefID,EmpID) VALUES(" & lVID_EntryID & ",'" & cmbChecker.List(cmbChecker.ListIndex, 1) & "')"
                End If
                
                 'Do Not Execute Stored Procedure
                 'If This User Requires Authorization...
                If UserReqAuth = 0 Then
                    With cmd
                        .ActiveConnection = con
                        .CommandType = adCmdStoredProc
                        .CommandText = "SP_UpdateForIARNew"    'Issue After Receive
                        
                        .Parameters("@ItemCode").Value = LVIssued.ListItems(i).SubItems(5)
                        .Parameters("@ProcessID").Value = cmbProcess.ID
                        .Parameters("@Qty").Value = Val(LVIssued.ListItems(i).SubItems(8))
                        .Parameters("@OrderNo").Value = LVIssued.ListItems(i).SubItems(4)
                        .Parameters("@Rcvd_RefID").Value = Val(LVIssued.ListItems(i).ListSubItems(10).Tag)
                        
                        .Execute
                    End With
                    
                End If
            End If
            'Authentication Info...
            If UserReqAuth = 0 Then
                Dim DetailRefID As Long
                DetailRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendIssdDetail").Fields(0).Value & "")
                con.Execute "INSERT INTO VendIssAuthDetail(UserID,IssID,DT) Values(" & _
                 UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
            End If
            
        End With
        con.Execute "INSERT INTO PrintSession(RecieptNo) Values('" & LVIssued.ListItems(i).ListSubItems(2).Tag & "')"
        'Debug.Print i & " of " & LVIssued.ListItems.count
    Next i
                       
    
    
    If IssLot Then
        For i = 1 To LVIssued.ListItems.count
            If Left(LVIssued.ListItems(i).SubItems(2), 2) = "HS" Then
                If UserReqAuth = 0 Then
                    con.Execute "Update VendRcvdDetail Set IssQty=RcvdQty Where RecieptID='" & LVIssued.ListItems(i).SubItems(2) & "'"
                End If
            End If
        Next
    End If
        
    If bReIssuance Then
        For i = 1 To LVIssd_Multi.ListItems.count
            With LVIssd_Multi.ListItems(i)
                If .Checked Then
                    con.Execute "INSERT INTO VID_Multiple_Process_Issuance(VID_RefID,ProcessID,Rate) VALUES(" & _
                        lVID_EntryID & "," & Val(.Tag) & "," & Val(.SubItems(1)) & ")"
                End If
            End With
        Next
    End If
    
    '-----------------------------------------------------------------
    If strCompany = "Banzai" Then
        If chkPrint.Value = vbChecked Then
            Call PrintCTC(lRefID)
        End If
    End If
    
    con.CommitTrans
    
    Saved = True
    
    MsgBox "Saved.", vbInformation
    
    Exit Function
err:
    Call RollBackTransactions
    MsgBox err.Description
    
End Function

Private Sub txtRate1_KeyDown(KeyCode As Integer, Shift As Integer)
'    MsgBox KeyCode
End Sub

Private Sub txtRate1_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    Dim iQty  As Long
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        If Val(txtRate1) <= 0 Then
            MsgBox "Invalid Data.", vbInformation
            Exit Sub
        End If
        Dim ITM As ListItem
        If iColNo = 9 Then
            'iQty = Val(LV.ListItems(Val(LVIssued.SelectedItem.Tag)).ListSubItems(4).Tag) - Val(txtRate1)
            iQty = Val(LV.ListItems(Val(LVIssued.SelectedItem.Tag)).SubItems(5)) - Val(txtRate1)
            If iQty < 0 Then
                MsgBox "Invalid Qty!", vbInformation
                Exit Sub
            End If
            LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(4) = iQty
            LVIssued.SelectedItem.SubItems(iColNo - 1) = Val(txtRate1)
        Else
            LVIssued.SelectedItem.SubItems(iColNo - 1) = Val(txtRate1)
        End If
        
        txtRate1.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
'        If Val(txtRate1) = 0 Then
'            Call LVIssued_KeyDown(46, 0)
'        End If
        
    ElseIf KeyAscii = 27 Then
        txtRate1.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtRate1_LostFocus()
    txtRate1.Visible = False
    'Me.KeyPreview = True
    cmdPost.Default = True
End Sub
Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub
Private Sub PostForLots()

    On Error GoTo err
    Dim ITM As ListItem

    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            IssID = LV.ListItems(i).Tag
            With ITM
                Set ITM = LVIssued.ListItems.add(, , LVIssued.ListItems.count + 1)
                ITM.Tag = LV.ListItems(i).key
                ITM.ListSubItems.add(, cmbVender.ID & "'", cmbVender.Text).Tag = cmbVender.ID
                ITM.ListSubItems.add(, , IssID & "").Tag = DTPicker2
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(2)
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(3)
                ITM.ListSubItems.add , , Val(cmbVender.Tag)
                ITM.ListSubItems.add , , LV.ListItems(i).SubItems(4)
                'itm.ListSubItems.Add , , Val(LV.ListItems(i).SubItems(5))
                If cmbProcess.Tag = "Lot No" Then
                    ITM.ListSubItems(6).Tag = LV.ListItems(i).ListSubItems(1).Tag
                End If
            End With
            LV.ListItems(i).SubItems(4) = 0
            
        End If
    Next
    
    
    'Remove Posted Entries From the List.
    i = 1
    Do While i <= LV.ListItems.count
        If LV.ListItems(i).Checked Then
            LV.ListItems.Remove i
            On Error Resume Next
            cmbDesc.RemoveItem i - 1
            If err.Number = -2147024809 Then Resume Next
                
            i = i - 1
        End If
        i = i + 1
    Loop
    
    'if cmbitems.
    'Call LV_ItemClick(LV.ListItems(i))
    
    Call FillVendList
    
    txtQty = ""
    Exit Sub
err:
    MsgBox err.Description

End Sub

Public Sub ShowMe(m_bReIssuance As Boolean)

    bReIssuance = m_bReIssuance
    
    If bReIssuance = False Then
        cmbProcess.AddItem "<All Processes>", "0"
        txtLotNo.Enabled = False
        
        lblSteel.Visible = True
        cmbSteel.Visible = True
        
        'cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " WHERE ProcessID IN(" & iReadyFinishProcessID & "," & iSemiFinishProcessID & "," & iFilingProcessID & "," & iFiling2ProcessID & "," & iMakerPOProess_I & "," & iMakerPOProess_II & "," & iMakerPOProess_III & "," & iMakerPOProess_IV & "," & iMakerPOProess_V & "," & iMakerPOProess_VI & "," & iMakerPOProess_VII & "," & iMakerPOProess_VIII & ") ORDER BY SNo"
        cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " WHERE ProcessID IN(SELECT ProcessID FROM Processes_Purchase) ORDER BY SNo"
    Else
        Label1(2).Visible = False
        txtExcess.Text = "0"
        txtExcess.Visible = False
        If strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
            txtLotNo.Enabled = True
        Else
            txtLotNo.Enabled = False
        End If
        Label1(0).Caption = "Process Issuance"
        Label1(1).Caption = "Process Issuance"
        cmbProcess.Enabled = False
        
        lblSteel.Visible = False
        cmbSteel.Visible = False
        cmbIssEmp.AddVals con, "'{' + EmpID + '} ' + Name + ' {' + Designation + '}'", "Employees", "EmpID", " WHERE Active=1"
        
        cmbCountedBy.ClearVals
        cmbCountedBy.AddVals con, "'{' + EmpID + '} ' + Name + ' {' + Designation + '}'", "Employees", "EmpID", " WHERE Active=1"
        
        chkPrint.Value = vbChecked
        
        cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " ORDER BY SNo"
    End If
    
    
    If bReIssuance = False Then cmbProcess.ID = "0"
    
    Me.Show
    
End Sub

Private Function GetNextMasterPONo() As String

    Dim Prefix As String
    'ORD-DDMM-001
    Dim myDT As Date
    myDT = Date
    Dim lOrderNo As Long
    lOrderNo = GetSingleLongValue("MAX(CAST(LEFT(MasterPONo,CHARINDEX('/',MasterPONo)-1) AS INT))", "VendIssued", " WHERE RIGHT(MasterPONo,4)='" & Format(myDT, "MMyy") & "' AND MasterPONo IS NOT NULL")
    lOrderNo = lOrderNo + 1
    GetNextMasterPONo = lOrderNo & Format(myDT, "/MMyy")
    
End Function

Private Function InspectionDataEntryCheck(lVRD_EntryID As Long, lCurrentProcessID As Long) As Boolean

    Dim lCount As Long, lCount_Done As Long
    lCount = GetSingleLongValue("COUNT(*)", "Process_Inspection_Points", " WHERE ProcessID=" & lCurrentProcessID)
    lCount_Done = GetSingleLongValue("COUNT(*)", "VRD_PIP_Details", " WHERE VRD_RefID=" & lVRD_EntryID)
    
    If lCount = lCount_Done Then
        InspectionDataEntryCheck = True
    Else
        InspectionDataEntryCheck = False
    End If
    
End Function
