VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmProductionPlanningDashBoard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   10560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductionPlanningDashBoard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   16410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrintDetails 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   10110
      Width           =   1215
   End
   Begin VB.CommandButton cmdTechDrawing 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tech. Draw."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   10110
      Width           =   1215
   End
   Begin VB.CommandButton cmdArticlewiseStatus 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Articlewise Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7725
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   10110
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdPrintAvgProcessDuration 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Avg. Duration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7725
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   10110
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtCompletion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   15345
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   10095
      Width           =   1000
   End
   Begin VB.CommandButton cmdShowPicture 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10110
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintSpecificationSheet 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Specs."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   10110
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintPPR 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PPR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7725
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   10110
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtQtyRequired 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   13515
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   10095
      Width           =   1000
   End
   Begin VB.TextBox txtMinStock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   11355
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   10095
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10110
      Width           =   1215
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   9495
      Index           =   0
      Left            =   60
      ScaleHeight     =   9495
      ScaleWidth      =   16290
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   16290
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   6060
         TabIndex        =   46
         Top             =   6435
         Width           =   5985
         Begin TabDlg.SSTab SSTab3 
            Height          =   2670
            Left            =   60
            TabIndex        =   57
            Top             =   210
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   4710
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Finish Stock"
            TabPicture(0)   =   "frmProductionPlanningDashBoard.frx":014A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label8"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "LVRFStock"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "txtTotalRFStock"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Look alike"
            TabPicture(1)   =   "frmProductionPlanningDashBoard.frx":0166
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "LVRFLA"
            Tab(1).ControlCount=   1
            Begin VB.TextBox txtTotalRFStock 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   4835
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   2310
               Width           =   1000
            End
            Begin MSComctlLib.ListView LVRFStock 
               Height          =   1980
               Left            =   30
               TabIndex        =   59
               Top             =   330
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   3493
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   12582912
               BackColor       =   12632256
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Qty"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Store"
                  Object.Width           =   3351
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Remarks"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LVRFLA 
               Height          =   2250
               Left            =   -74955
               TabIndex        =   61
               Top             =   375
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   3969
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   12582912
               BackColor       =   12632256
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Item"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Qty"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Store"
                  Object.Width           =   3351
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Remarks"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total :"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   4125
               TabIndex        =   60
               Top             =   2310
               Width           =   645
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer Reference"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   12105
         TabIndex        =   44
         Top             =   6435
         Width           =   4095
         Begin MSComctlLib.ListView LVCustCatalog 
            Height          =   2550
            Left            =   75
            TabIndex        =   45
            Top             =   270
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4498
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   12632256
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Customer"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Ref #"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "OCR #"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   1411
            EndProperty
         End
      End
      Begin VB.Frame FForgItems 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Forging Detail"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2610
         Left            =   12105
         TabIndex        =   36
         Top             =   6435
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtTotalForgDetail 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   3050
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   2130
            Width           =   1000
         End
         Begin MSComctlLib.ListView LVFD 
            Height          =   1830
            Left            =   75
            TabIndex        =   37
            Top             =   300
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   3228
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   12632256
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Ord.Qty"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   2355
            TabIndex        =   43
            Top             =   2190
            Width           =   645
         End
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   330
         Left            =   30
         TabIndex        =   28
         Top             =   225
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbArticle 
         Height          =   330
         Left            =   3885
         TabIndex        =   4
         Top             =   225
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   582
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "RM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   12105
         TabIndex        =   24
         Top             =   3525
         Width           =   4095
         Begin VB.CommandButton cmdReserveForging 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Reserve"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtTotalForging 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   3035
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   420
            Visible         =   0   'False
            Width           =   1000
         End
         Begin MSComctlLib.ListView LVForging 
            Height          =   2565
            Left            =   60
            TabIndex        =   26
            Top             =   300
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4524
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   12632256
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "RM"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Qty"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Location"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   2325
            TabIndex        =   27
            Top             =   465
            Visible         =   0   'False
            Width           =   645
         End
      End
      Begin VB.Frame FraVenders 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Vendors"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   30
         TabIndex        =   11
         Top             =   6435
         Width           =   5985
         Begin MSComctlLib.ListView LVVenders 
            Height          =   2550
            Left            =   60
            TabIndex        =   12
            Top             =   315
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   4498
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   12632256
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Vendor"
               Object.Width           =   4233
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Process"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Rate"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin VB.Frame FraSFStock 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Semi Finish Stock"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   6060
         TabIndex        =   10
         Top             =   3525
         Width           =   5985
         Begin TabDlg.SSTab SSTab2 
            Height          =   2625
            Left            =   75
            TabIndex        =   52
            Top             =   255
            Width           =   5850
            _ExtentX        =   10319
            _ExtentY        =   4630
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Semi Finish Stock"
            TabPicture(0)   =   "frmProductionPlanningDashBoard.frx":0182
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label6"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "LVSFStock"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "txtTotalSFStock"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Look alike"
            TabPicture(1)   =   "frmProductionPlanningDashBoard.frx":019E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "LVSFLA"
            Tab(1).ControlCount=   1
            Begin VB.TextBox txtTotalSFStock 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   4785
               Locked          =   -1  'True
               TabIndex        =   53
               Top             =   2265
               Width           =   1000
            End
            Begin MSComctlLib.ListView LVSFStock 
               Height          =   1920
               Left            =   30
               TabIndex        =   54
               Top             =   345
               Width           =   5760
               _ExtentX        =   10160
               _ExtentY        =   3387
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   12582912
               BackColor       =   12632256
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Process"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Qty"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Store"
                  Object.Width           =   3351
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Remarks"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView LVSFLA 
               Height          =   2220
               Left            =   -74955
               TabIndex        =   56
               Top             =   360
               Width           =   5745
               _ExtentX        =   10134
               _ExtentY        =   3916
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   12582912
               BackColor       =   12632256
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   5
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Process"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Item"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "Qty"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Store"
                  Object.Width           =   3351
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Remarks"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total :"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   4110
               TabIndex        =   55
               Top             =   2280
               Width           =   645
            End
         End
      End
      Begin VB.Frame FraOpenPO 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Open Purchase Orders"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   30
         TabIndex        =   8
         Top             =   3525
         Width           =   5985
         Begin VB.TextBox txtTotalOpenPO 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   2475
            Width           =   1000
         End
         Begin MSComctlLib.ListView LVOpenPO 
            Height          =   2190
            Left            =   60
            TabIndex        =   9
            Top             =   300
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   3863
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   12632256
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PO No."
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Maker"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Order No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "P.Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Del."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Maker PO"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4230
            TabIndex        =   21
            Top             =   2505
            Width           =   645
         End
      End
      Begin VB.Frame FraRunningLots 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Running Lots"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   6060
         TabIndex        =   6
         Top             =   615
         Width           =   10125
         Begin VB.TextBox txtTotalRunningLots 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   9050
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   2475
            Width           =   1000
         End
         Begin MSComctlLib.ListView LVRunningLots 
            Height          =   2190
            Left            =   60
            TabIndex        =   7
            Top             =   300
            Width           =   9990
            _ExtentX        =   17621
            _ExtentY        =   3863
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   12632256
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Lot No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Process"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Order No."
               Object.Width           =   3545
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Customer"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "R2 In"
               Object.Width           =   2671
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "From"
               Object.Width           =   3466
            EndProperty
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   8340
            TabIndex        =   19
            Top             =   2505
            Width           =   645
         End
      End
      Begin VB.Frame FraCustomerOrders 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   30
         TabIndex        =   5
         Top             =   615
         Width           =   5985
         Begin TabDlg.SSTab SSTab1 
            Height          =   2685
            Left            =   75
            TabIndex        =   47
            Top             =   195
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   4736
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Customer Orders"
            TabPicture(0)   =   "frmProductionPlanningDashBoard.frx":01BA
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label4"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "LVCustOrders"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "txtTotalCustOrders"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Comp/Feedback"
            TabPicture(1)   =   "frmProductionPlanningDashBoard.frx":01D6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "LVComplaints"
            Tab(1).Control(1)=   "CD1"
            Tab(1).ControlCount=   2
            Begin VB.TextBox txtTotalCustOrders 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   345
               Left            =   4770
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   2310
               Width           =   1000
            End
            Begin MSComctlLib.ListView LVCustOrders 
               Height          =   1935
               Left            =   45
               TabIndex        =   48
               Top             =   360
               Width           =   5730
               _ExtentX        =   10107
               _ExtentY        =   3413
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   12582912
               BackColor       =   12632256
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   6
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Cust Code"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Order No."
                  Object.Width           =   2468
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "Date"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "Qty"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Text            =   "Del."
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Remarks"
                  Object.Width           =   2469
               EndProperty
            End
            Begin MSComctlLib.ListView LVComplaints 
               Height          =   2250
               Left            =   -74940
               TabIndex        =   51
               Top             =   375
               Width           =   5730
               _ExtentX        =   10107
               _ExtentY        =   3969
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   12582912
               BackColor       =   12632256
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Complaint"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Remarks"
                  Object.Width           =   2469
               EndProperty
            End
            Begin MSComDlg.CommonDialog CD1 
               Left            =   -75000
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
               CancelError     =   -1  'True
               DialogTitle     =   "Please Select Picture"
               Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total :"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   4050
               TabIndex        =   50
               Top             =   2310
               Width           =   645
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Processes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2910
         Left            =   7950
         TabIndex        =   22
         Top             =   3525
         Visible         =   0   'False
         Width           =   4095
         Begin MSComctlLib.ListView LVProcesses 
            Height          =   2595
            Left            =   75
            TabIndex        =   23
            Top             =   255
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4577
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   12632256
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Rate"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2070
         TabIndex        =   31
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   30
         Top             =   0
         Width           =   810
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   330
         Left            =   2055
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   225
         Width           =   1800
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3175;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   420
         Left            =   14955
         TabIndex        =   16
         Top             =   165
         Width           =   1230
         ForeColor       =   65280
         BackColor       =   4210752
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2170;741"
         Accelerator     =   82
         FontName        =   "Calibri"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Article"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3885
         TabIndex        =   3
         Top             =   15
         Width           =   540
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comp. :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Index           =   1
      Left            =   14535
      TabIndex        =   39
      Top             =   10140
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Required :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Index           =   0
      Left            =   12405
      TabIndex        =   17
      Top             =   10140
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min. Stock Lvl :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   9750
      TabIndex        =   15
      Top             =   10155
      Width           =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTION PLANNING DASHBOARD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   16305
   End
End
Attribute VB_Name = "frmProductionPlanningDashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long
Dim bHideRate As Boolean

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(220, 219, 244) 'RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    Dim lBorderColor As Long, lForeColor As Long
    For i = 0 To PicWC.count - 1
        If i = 1 Then
            lBorderColor = RGB(153, 204, 255)
            lForeColor = 0
        ElseIf i = 0 Then
            lBorderColor = RGB(220, 219, 244) 'RGB(255, 204, 153)
            lForeColor = 0
        ElseIf i = 3 Then
            lBorderColor = RGB(255, 153, 204)
            lForeColor = 0
        ElseIf i = 2 Then
            lBorderColor = RGB(0, 255, 0)
            lForeColor = 0
        ElseIf i = 4 Then
            lBorderColor = RGB(200, 200, 192)
            lForeColor = 0
        End If
        Call DrawCaptions(PicWC(i), lBorderColor, lForeColor)
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox, lBorderColor As Long, lForeColor As Long)

    Dim lFillColor
    'lBorderColor = 0 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    lFillColor = lBorderColor 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 150)-(Pic.Width, 150), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 300), lFillColor, BF
        Pic.ForeColor = lForeColor 'vbGreen
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 0
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 50)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub

Private Sub cmbArticle_matched()
    Call RefreshDashBoard
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    cmbArticle.ClearVals
    cmbArticle.DropDownOffLine = True
    If cmbCust.ID = "0" Then
        cmbArticle.AddVals con, "'{' + ItemID + '} '+ItemName", "Items", "ItemID", " ORDER BY ItemID"
    Else
        cmbArticle.AddVals con, "'{' + CompItemID + '} ' + ItemID + ' {' + ISNULL(OCR,'') + '} {' + ISNULL(OCR2,'') + '} ' + CompDesc", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
    End If
    
End Sub

Private Sub cmbCust_matched()
    
    If cmbCust.ID = "0" Then
        cmbCountry.Clear
        cmbCountry.AddItem "<Comp>"
        cmbCountry.ListIndex = 0
        cmbCountry.Enabled = False
        Exit Sub
    Else
        cmbCountry.Enabled = True
    End If
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Country,Curr FROM ForeignCustomers WHERE CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
End Sub

Private Sub cmdArticlewiseStatus_Click()
    If cmbArticle.MatchFound = False Then
        MsgBox "Please Select Article.", vbInformation
        Exit Sub
    End If
    Dim f As New frmArticlewiseShippedStatus
    Load f
    Call f.ShowMe(cmbArticle.ID, True)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Function GetTotal() As Double

    Dim i As Integer
    Dim lCustOrdersQty As Long
    Dim lRunLotsQty As Long, lOpenPOQty As Long, lSFStockQty As Long, lRFStockQty As Long, lTotalForging As Long
    
    lCustOrdersQty = 0
    For i = 1 To LVCustOrders.ListItems.count
        With LVCustOrders.ListItems(i)
            lCustOrdersQty = lCustOrdersQty + Val(.SubItems(3))
        End With
    Next
    txtTotalCustOrders = lCustOrdersQty
    
    lRunLotsQty = 0
    For i = 1 To LVRunningLots.ListItems.count
        With LVRunningLots.ListItems(i)
            lRunLotsQty = lRunLotsQty + Val(.SubItems(2))
        End With
    Next
    txtTotalRunningLots = lRunLotsQty
    
    lOpenPOQty = 0
    For i = 1 To LVOpenPO.ListItems.count
        With LVOpenPO.ListItems(i)
            lOpenPOQty = lOpenPOQty + Val(.SubItems(3))
        End With
    Next
    txtTotalOpenPO = lOpenPOQty
    
    lSFStockQty = 0
    For i = 1 To LVSFStock.ListItems.count
        With LVSFStock.ListItems(i)
            lSFStockQty = lSFStockQty + Val(.SubItems(1))
        End With
    Next
    txtTotalSFStock = lSFStockQty
    
    lRFStockQty = 0
    For i = 1 To LVRFStock.ListItems.count
        With LVRFStock.ListItems(i)
            lRFStockQty = lRFStockQty + Val(.Text)
        End With
    Next
    txtTotalRFStock = lRFStockQty
    
    lTotalForging = 0
    For i = 1 To LVForging.ListItems.count
        With LVForging.ListItems(i)
            lTotalForging = lTotalForging + Val(.SubItems(1)) + Val(.SubItems(2))
        End With
    Next
    txtTotalForging = lTotalForging
    
    txtQtyRequired = lCustOrdersQty + Val(txtMinStock) - (lRunLotsQty + lOpenPOQty + lSFStockQty + lRFStockQty)
    
End Function
    
Private Sub cmdPrintAvgProcessDuration_Click()

    If cmbArticle.MatchFound = False Then Exit Sub
    Dim myDTFrom As Date, myDTTo As Date
    Dim frmDT As New frmGetDateDouble
    Load frmDT
    If frmDT.getDate(myDTFrom, myDTTo) = False Then Exit Sub
        
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemWiseAvgProcessDuration.rpt")
    rpt.ParameterFields.GetItemByName("@ItemID").SetCurrentValue cmbArticle.ID
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
End Sub

Private Sub cmdPrintDetails_Click()

    On Error GoTo err
    
    If cmbArticle.MatchFound = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemSpecification_Detail.rpt")
    Dim f As New frmPrevRpt, strCondition As String
    strCondition = "{Items.ItemID}='" & cmbArticle.ID & "'"
    f.ShowReport strCondition, rpt
   
    Exit Sub
err:
    MsgBox err.Description


End Sub

Private Sub cmdPrintPPR_Click()

    Dim strItems As String
    Dim i As Integer
    
    
    If cmbArticle.MatchFound = False Then
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    End If
    strItems = "'" & cmbArticle.ID & "'"
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningReport.rpt")
    
    'rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrderNo
    rpt.ParameterFields.GetItemByName("@FromItems").SetCurrentValue True
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue #1/1/2010# 'DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue Date 'DtTo.Value
    rpt.ParameterFields.GetItemByName("@NoDelivery").SetCurrentValue False
    rpt.ParameterFields.GetItemByName("@Both").SetCurrentValue False
    
    rpt.ParameterFields.GetItemByName("@ItemCode").SetCurrentValue strItems
    rpt.ParameterFields.GetItemByName("@TempList_RefID").SetCurrentValue 0
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Private Sub cmdPrintSpecificationSheet_Click()

    On Error GoTo err
    
    If cmbArticle.MatchFound = False Then
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemSpecification.rpt")
    Dim f As New frmPrevRpt, strCondition As String
    strCondition = "{Items.ItemID}='" & cmbArticle.ID & "'"
    f.ShowReport strCondition, rpt
   
    Exit Sub
err:
    MsgBox err.Description

End Sub

Public Sub cmdRefresh_Click()
    Call cmbArticle_matched
End Sub


Private Sub cmdReserveForging_Click()
    Dim f As New frmReserveForggingOpenPO
    Load f
    f.ReserveForgging cmbArticle.ID, cmbArticle.Text
End Sub

Private Sub cmdShowPicture_Click()

    If cmbArticle.MatchFound = False Then Exit Sub
    
    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromDBField "ItemPic", "Items", " WHERE ItemID='" & cmbArticle.ID & "'"
    
End Sub



Private Sub cmdTechDrawing_Click()

    Call ShowTechDrawing("Items", "TechnicalDrawing", " WHERE ItemID='" & cmbArticle.ID & "'", cmbArticle.ID, MainForm.CD1)
    
End Sub

Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
    
    If (strCompany = "Banzai" Or strCompany = "QEL") Then
        cmdPrintPPR.Visible = False
        cmdPrintAvgProcessDuration.Visible = False
        cmdArticlewiseStatus.Visible = False
    End If
    
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    SSTab3.Tab = 0
        
    If CurrentUserName <> "Administrator" Then
        bHideRate = GetSingleBooleanValue("HideRateMakerAssign", "Users", " WHERE UserID=" & UserID)
    Else
        bHideRate = False
    End If
    
End Sub


Private Sub LVComplaints_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LVComplaints.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print Complaint", "View PDF")
    
    
    If strRet = "Print Complaint" Then
        Dim rpt As CRAXDRT.Report
        Set rpt = rptApp.OpenReport(RptPath & "\CustomerComplaints.rpt")
        Dim f As New frmPrevRpt
        f.ShowReport "{CustomerComplaints.EntryID}=" & Val(LVComplaints.SelectedItem.Tag), rpt
    ElseIf strRet = "View PDF" Then
        Call ShowTechDrawing("CustomerComplaints", "AttachmentPDF", " WHERE EntryID=" & Val(LVComplaints.SelectedItem.Tag), Left(LVComplaints.SelectedItem.Text, 5), CD1)
    End If
    
End Sub

Private Sub LVCustCatalog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVCustCatalog
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
    
End Sub

Private Sub LVCustOrders_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVCustOrders
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
    
End Sub

Private Sub LVCustOrders_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LVCustOrders.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        strRet = ShowPopUpMenu1("Filter by Selected", "Show All", "Print Order Balance", "R1-Order Planning", "R2-Order Planning", "-", "Initial Planning Memo")
    Else
        strRet = ShowPopUpMenu1("Filter by Selected", "Show All", "Print Order Status", "Print Order Balance", "R1-Order Planning", "R2-Order Planning")
    End If
    
    If strRet = "Filter by Selected" Then
        Call FilterByOrderNo(LVCustOrders.SelectedItem.Tag)
    ElseIf strRet = "Show All" Then
        Call FilterByOrderNo("")
    ElseIf strRet = "Print Order Balance" Then
        Call PrintOrderBalanceReport(LVCustOrders.SelectedItem.SubItems(1))
    ElseIf strRet = "R1-Order Planning" Then
        Call PrintOrderPlanningReportR1(LVCustOrders.SelectedItem.SubItems(1), True)
    ElseIf strRet = "R2-Order Planning" Then
        Call PrintOrderPlanningReportR1(LVCustOrders.SelectedItem.SubItems(1), False)
    ElseIf strRet = "Initial Planning Memo" Then
        Call ShowInitialPlanningMemo(Val(LVCustOrders.SelectedItem.ListSubItems(2).Tag), Val(txtTotalRunningLots), Val(txtTotalOpenPO), Val(txtTotalSFStock), Val(txtTotalRFStock), Val(txtTotalForgDetail))
    ElseIf strRet = "Print Order Status" Then
        Call PrintOrderStatusReport(LVCustOrders.SelectedItem.Tag)
    End If
    
End Sub


Private Sub LVFD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVFD
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
    
End Sub

Private Sub LVForging_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVForging
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
    
End Sub

Private Sub LVForging_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim lTotal As Long
    With rs
        .Open "SELECT VItemsRMComp.*,UnShippedQty FROM VItemsRMComp INNER JOIN (SELECT CompItemCode,SUM(Qty-ShippedQty) AS UnShippedQty FROM VFOrderItemsWIthShippedQty WHERE Qty>ShippedQty AND OrderNo IN(SELECT OrderNo FROM VUnShippedOrderList) AND OrderNo NOT IN('Stock-Order','Stock-OrderIss') GROUP BY CompItemCode) T1 ON VItemsRMComp.ItemID=T1.CompItemCode WHERE RMID=" & Val(Item.Tag) & " AND ItemID<>'" & cmbArticle.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LVFD.ListItems.Clear
        Do Until .EOF
            Set ITM = LVFD.ListItems.add(, , !ITemID & "")
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , Val(!UnShippedQty & "")
            lTotal = lTotal + Val(!UnShippedQty & "")
            .MoveNEXT
        Loop
        .Close
    End With
    txtTotalForgDetail = lTotal
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LVForging_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If LVForging.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim str As String
'    If LVForging.SelectedItem.ForeColor = vbRed Then    'In-Active
'        str = "Make Active"
'    Else
'        str = "Make In-Active"
'    End If
    
    Dim strReturn As String, lValue As Long
    strReturn = ShowPopUpMenu1("View RM Ledger", "Print Report")
    
    If strReturn = "Make In-Active" Then
        lValue = 1
        con.Execute "UPDATE ItemsRMComp SET Functional_Status=1 WHERE RMID=" & Val(LVForging.SelectedItem.ListSubItems(1).Tag) & " AND ItemID='" & cmbArticle.ID & "'"
        Call RefreshForging
    ElseIf strReturn = "Make Active" Then
        lValue = 0
        con.Execute "UPDATE ItemsRMComp SET Functional_Status=0 WHERE RMID=" & Val(LVForging.SelectedItem.ListSubItems(1).Tag) & " AND ItemID='" & cmbArticle.ID & "'"
        Call RefreshForging
    ElseIf strReturn = "View RM Ledger" Then
        Dim f As New frmStockLedger
        Load f
        f.cmbMaterial.ID = LVForging.SelectedItem.Text
        f.DtFrom.Value = DateAdd("M", -1, Date)
        f.DtTo.Value = Date
        f.Show
    ElseIf strReturn = "Print Report" Then
        Call PrintRMLinkedReport
    End If
    
    
    
    
End Sub

Private Sub LVOpenPO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVOpenPO
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
End Sub

Private Sub LVOpenPO_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LVOpenPO.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print PO", "Print Master PO Status", "Close PO")
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt
    If strRet = "Print PO" Then
        Set rpt = rptApp.OpenReport(RptPath & "\IssList.rpt")
        f.ShowReport "{VendIssued.MasterPONo}='" & LVOpenPO.SelectedItem.ListSubItems(4).Tag & "'", rpt
    ElseIf strRet = "Print Master PO Status" Then
        Set rpt = rptApp.OpenReport(RptPath & "\MasterPOStatus.rpt")
        f.ShowReport "{VendIssued.MasterPONo}='" & LVOpenPO.SelectedItem.ListSubItems(4).Tag & "'", rpt
    ElseIf strRet = "Close PO" Then
        If getDBPassword("CloseMakerPO") Then
            con.Execute "UPDATE VendIssued SET Closed=1 WHERE EntryID=" & Val(LVOpenPO.SelectedItem.Tag)
            Call RefreshOpenPOs
        End If
    End If
    
End Sub

Private Sub LVRFStock_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVRFStock
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
    
End Sub

Private Sub LVRFStock_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If LVRFStock.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim str As String
'    If LVForging.SelectedItem.ForeColor = vbRed Then    'In-Active
'        str = "Make Active"
'    Else
'        str = "Make In-Active"
'    End If
    
    Dim strReturn As String, lValue As Long
    strReturn = ShowPopUpMenu1("Print Report")
    
    If strReturn = "Print Report" Then
        Call PrintFinishStoreLots
    End If
    
End Sub

Private Sub LVRunningLots_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVRunningLots
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
    
End Sub

Private Sub LVRunningLots_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo err
    If LVRunningLots.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    
    Dim strRet As String
    If Val(LVRunningLots.SelectedItem.ListSubItems(6).Tag) > 0 Then 'IssuanceEntryID
        strRet = ShowPopUpMenu1("Print PTC", "Print PTC Mini", "Add Bookmark for Receiving", "Print Running Lots")
    Else
        strRet = ShowPopUpMenu1("Print PTC", "Print PTC Mini", "Split/Transfer Lot", "Print Running Lots", "Close Lot")
    End If
    
    'strRet = ShowPopUpMenu1("Print PTC")
    
    If strRet = "Add Bookmark for Receiving" Then
        Dim frmBookM As New frmAddBookmarkforReceiving
        Load frmBookM
        frmBookM.AddBookmark (Val(LVRunningLots.SelectedItem.ListSubItems(6).Tag))
    ElseIf strRet = "Print PTC" Then
        Call PrintPTC(LVRunningLots.SelectedItem.Text, False)
'        Dim rpt As CRAXDDRT.Report
'        Dim f As New frmPrevRpt, strSelection As String
'
'        Dim strLotNo As String, bRepairLot As Boolean
'        strLotNo = LVRunningLots.SelectedItem.Text
'        bRepairLot = GetSingleBooleanValue("ReWorkLot", "VendIssdDetail", " WHERE LotNo='" & strLotNo & "'")
'
'        If bRepairLot Then
'            Set rpt = rptApp.OpenReport(RptPath & "\PTC_Repair.rpt")
'        Else
'            Set rpt = rptApp.OpenReport(RptPath & "\PTCQEL.rpt")
'        End If
'
'        If bRepairLot = False And strCompany = "Dr-Frgz" Then
'            rpt.ParameterFields.GetItemByName("@LotNo").SetCurrentValue strLotNo
'            rpt.EnableParameterPrompting = False
'            f.ShowReport "", rpt, "", False
'        Else
'            strSelection = "{VFOrderItemPTC.LotNo}='" & strLotNo & "'"
'            f.ShowReport strSelection, rpt
'        End If
    ElseIf strRet = "Print PTC Mini" Then
        Call PrintPTC(LVRunningLots.SelectedItem.Text, False, 1)
    ElseIf strRet = "Split/Transfer Lot" Then
        Dim frmSplit As New frmMakerRcvList
        Load frmSplit
        Call frmSplit.ShowForSplit(LVRunningLots.SelectedItem.Text, Val(LVRunningLots.SelectedItem.ListSubItems(5).Tag))
    ElseIf strRet = "Print Running Lots" Then
        Dim rpt As CRAXDDRT.Report
        Set rpt = rptApp.OpenReport(RptPath & "\PPDashBoard_RunningLots.rpt")
        Dim f As New frmPrevRpt
        f.ShowReport "{VRunningLots_Simple.ItemCode}='" & cmbArticle.ID & "'", rpt
    ElseIf strRet = "Close Lot" Then
        'If getDBPassword("Close_Lots") = False Then Exit Sub
'        If GetSingleBooleanValue("Close_Lot", "Users", " WHERE UserName='" & CurrentUserName & "'") = False Then
'            MsgBox "You don't have access.", vbInformation
'            Exit Sub
'        End If
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            If getDBPassword("Close_Lot") = False Then Exit Sub
            con.Execute "INSERT INTO Lots_Closed(LotNo,UserName,MachineName) VALUES('" & LVRunningLots.SelectedItem.Text & "','" & CurrentUserName & "','" & strComputerName & "')"
            LVRunningLots.ListItems.Remove LVRunningLots.SelectedItem.Index
            Call GetRunningLotsTotal
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LVSFStock_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVSFStock
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
    
End Sub

Private Sub LVVenders_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LVVenders
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            myLV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        'myLV.ColumnHeaderIcons = Nothing
        'Set myLV.ColumnHeaderIcons = ImageList2
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    If myLV.Sorted = False Then myLV.Sorted = True
    
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = x
                lDownY = y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
End Sub

Public Function ShowMe(Optional strItemID As String = "") As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
'    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
'    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show
    
'    cmbArticle.ListHeight = 5800
'    cmbArticle.DropDownOffLine = True
'    Call cmbArticle.AddVals(con, "'{ ' + ItemID + ' } ' + ItemName ", "Items", "ItemID", " WHERE InActive=0")
    
    cmbCust.ListHeight = 4000
    cmbCust.AddItem "<Comp>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " WHERE Active=1")
    cmbCust.ID = "0"
    
    If strItemID <> "" Then
        cmbArticle.ClearVals
        cmbArticle.DropDownOffLine = False
        cmbArticle.AddVals con, "'{' + ItemID + '} '+ItemName", "Items", "ItemID", " WHERE ItemID='" & strItemID & "' ORDER BY ItemID"
        cmbArticle.ID = strItemID
        'cmbarticle.AddVals con,
    End If
    
    cmbArticle.ListHeight = 5800
    
    Me.Show
    ShowMe = bSaved
    
End Function


Private Sub RefreshDashBoard()
    
    If cmbArticle.MatchFound = False Then Exit Sub
    Dim strItemCode As String
    strItemCode = cmbArticle.ID
    Dim rs As New ADODB.Recordset, ITM As ListItem, lForeColor As Long
    Dim i As Integer
    With rs
        Call RefreshCustomerOrders("")
        
        Call RefreshRunningLots("")
        
        Call RefreshOpenPOs("")
        
        
        'Refresh S/F Stocks
        If strCompany = "QEL" Then
            .Open "SELECT Description,RcvdQty-IssQty AS Qty,'-' AS StoreName,'' AS Remarks FROM VendRcvdDetail INNER JOIN VStockOrderOpening ON " & _
             "VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID WHERE VendRcvdDetail.ItemCode='" & strItemCode & "' AND RcvdQty>IssQty", con, adOpenForwardOnly, adLockReadOnly
        Else
            '.Open "SELECT Description,Qty-IssdQty AS Qty,StoreName + ' (' + RackNo + ') {' + ShelfNo + '}' AS StoreName,Remarks FROM VStockOrderOpening INNER JOIN (SELECT Opening_RefID FROM VendRcvdDetail GROUP BY Opening_RefID) T1 ON T1.Opening_RefID=VStockOrderOpening.EntryID  " & _
             " WHERE VStockOrderOpening.ItemID='" & strItemCode & "' AND Qty>IssdQty", con, adOpenForwardOnly, adLockReadOnly
            .Open "SELECT Description,Qty-IssdQty AS Qty,StoreName + ' (' + RackNo + ') {' + ShelfNo + '}' AS StoreName,LotNo_Manual FROM VStockOrderOpening" & _
             " WHERE VStockOrderOpening.ItemID='" & strItemCode & "' AND Qty>IssdQty", con, adOpenForwardOnly, adLockReadOnly
        End If
        LVSFStock.ListItems.Clear
        Do Until .EOF
            Set ITM = LVSFStock.ListItems.add(, , !Description & "")
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !StoreName & ""
            ITM.ListSubItems.add , , !LotNo_Manual & ""
            .MoveNEXT
        Loop
        .Close
         
        'Refresh S/F Lookalikes.
        If strCompany = "QEL" Then
            .Open "SELECT ItemID,Description,RcvdQty-IssQty AS Qty,'-' AS StoreName,'' AS Remarks FROM VendRcvdDetail INNER JOIN VStockOrderOpening ON " & _
             "VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID WHERE VendRcvdDetail.ItemCode IN(SELECT ItemID_LookAlike FROM ItemLookalikes WHERE ItemID='" & strItemCode & "') AND RcvdQty>IssQty", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT VStockOrderOpening.ItemID,Description,Qty-IssdQty AS Qty,StoreName + ' (' + RackNo + ') {' + ShelfNo + '}' AS StoreName,Remarks FROM VStockOrderOpening INNER JOIN (SELECT Opening_RefID FROM VendRcvdDetail GROUP BY Opening_RefID) T1 ON T1.Opening_RefID=VStockOrderOpening.EntryID  " & _
             " WHERE VStockOrderOpening.ItemID IN(SELECT ItemID_LookAlike FROM ItemLookalikes WHERE ItemID='" & strItemCode & "') AND Qty>IssdQty", con, adOpenForwardOnly, adLockReadOnly
        End If
        LVSFLA.ListItems.Clear
        Do Until .EOF
            Set ITM = LVSFLA.ListItems.add(, , !Description & "")
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !StoreName & ""
            ITM.ListSubItems.add , , !Remarks & ""
            .MoveNEXT
        Loop
        .Close
        
        'Refresh Venders
        .Open "SELECT VenderName,Rate,Description FROM VendAssItems INNER JOIN Makers ON VendAssItems.VendID=Makers.VendID INNER JOIN Processes ON VendAssItems.ProcessID=Processes.ProcessID WHERE ItemID='" & strItemCode & "' AND Makers.VendID<>" & iFactoryMakerID, con, adOpenForwardOnly, adLockReadOnly
        LVVenders.ListItems.Clear
        Dim p_strRate As String
        Do Until .EOF
            Set ITM = LVVenders.ListItems.add(, , !VenderName & "")
            ITM.ListSubItems.add , , !Description & ""
            If bHideRate Then
                p_strRate = "-"
            Else
                p_strRate = Val(!Rate & "")
            End If
            ITM.ListSubItems.add , , p_strRate
            .MoveNEXT
        Loop
        .Close
        
        
        'Refresh R/F Stocks
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
            .Open "SELECT ItemID,StoreName + ' (' + RackNo + ') {' + ShelfNo + '}' AS StoreName,Remarks,SUM(NetQty) AS NetQty FROM VItemsWithShelfWiseStock WHERE ItemID='" & strItemCode & "' GROUP BY ItemID,StoreName + ' (' + RackNo + ') {' + ShelfNo + '}',Remarks HAVING SUM(Netqty)>0", con, adOpenForwardOnly, adLockReadOnly
            LVRFStock.ListItems.Clear
            Do Until .EOF
                Set ITM = LVRFStock.ListItems.add(, , Val(!NetQty & ""))
                ITM.ListSubItems.add , , !StoreName & ""
                ITM.ListSubItems.add , , !Remarks & ""
                .MoveNEXT
            Loop
            .Close
            
            .Open "SELECT ItemID,StoreName + ' (' + RackNo + ') {' + ShelfNo + '}' AS StoreName,Remarks,SUM(NetQty) AS NetQty FROM VItemsWithShelfWiseStock WHERE ItemID IN(SELECT ItemID_LookAlike FROM ItemLookalikes WHERE ItemID='" & strItemCode & "') GROUP BY ItemID,StoreName + ' (' + RackNo + ') {' + ShelfNo + '}',Remarks HAVING SUM(Netqty)>0", con, adOpenForwardOnly, adLockReadOnly
            LVRFLA.ListItems.Clear
            Do Until .EOF
                Set ITM = LVRFLA.ListItems.add(, , !ITemID)
                ITM.ListSubItems.add , , Val(!NetQty & "")
                ITM.ListSubItems.add , , !StoreName & ""
                ITM.ListSubItems.add , , !Remarks & ""
                .MoveNEXT
            Loop
            .Close
        End If
'        'Forging
'        .Open "SELECT VRM.RMID1,VRM.RMName,VRM.QtyInStock,QtyToRcv,VItemsRMComp.RMID FROM VItemsRMComp INNER JOIN VRM ON VRM.RMID=VItemsRMComp.RMID LEFT OUTER JOIN " & _
'         "(SELECT RMID1,SUM(QtyToRcv) AS QtyToRcv FROM VVendOrdersToRcv GROUP BY RMID1) T1 ON VRM.RMID1=T1.RMID1 WHERE ItemID='" & _
'         strItemCode & "' AND VRM.GroupID IN(SELECT Group_ID FROM RMGroupIDsForForging)", con, adOpenForwardOnly, adLockReadOnly
'        LVForging.ListItems.Clear
'        Do Until .EOF
'            Set ITM = LVForging.ListItems.add(, , !RMID1 & "")
'            ITM.Tag = Val(!RMID & "")
'            ITM.ListSubItems.add , , Val(!QtyInStock & "")
'            ITM.ListSubItems.add , , Val(!QtyToRcv & "")
'            .MoveNEXT
'        Loop
'        .Close
'        If LVForging.ListItems.count > 0 Then
'            LVForging.ListItems(1).Selected = True
'            Call LVForging_ItemClick(LVForging.ListItems(1))
'        Else
'            LVFD.ListItems.Clear
'        End If
        
        
        'RM
        Call RefreshForging
        
        'Item Processes
        .Open "SELECT Description,Rate FROM VItemProcesses WHERE IsExist='" & strItemCode & "' ORDER BY ItemSNo", con, adOpenForwardOnly, adLockReadOnly
        LVProcesses.ListItems.Clear
        Do Until .EOF
            Set ITM = LVProcesses.ListItems.add(, , !Description & "")
            ITM.ListSubItems.add , , Val(!Rate & "")
            .MoveNEXT
        Loop
        .Close
        
        'Customer Ref.
        .Open "SELECT FCustomerCatalog.CustCode,ItemID,OCR,Qty FROM FCustomerCatalog " & _
            "LEFT OUTER JOIN (SELECT CustCode,CompItemCode,SUM(Qty) AS Qty FROM FOrderItems INNER JOIN FCustomerOrders " & _
            "ON FCustomerOrders.OrderNo=FOrderItems.OrderNo WHERE DT>='1/1/12' GROUP BY CustCode,CompItemCode) TOrders ON  " & _
            "FCustomerCatalog.CustCode=TOrders.CustCode AND FCustomerCatalog.CompItemID=TOrders.CompItemCode WHERE CompItemID='" & strItemCode & "' AND FCustomerCatalog.CustCode<>'Stock'", con, adOpenForwardOnly, adLockReadOnly
        LVCustCatalog.ListItems.Clear
        Do Until .EOF
            Set ITM = LVCustCatalog.ListItems.add(, , !CustCode & "")
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , !OCR & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            .MoveNEXT
        Loop
        .Close
        
        'Customer Complaints...
        .Open "SELECT CustomerComplaints.EntryID,CustomerComplaints.ComplaintNo,Description FROM CustomerComplaints INNER JOIN CustomerComplaints_Detail ON CustomerComplaints.EntryID=CustomerComplaints_Detail.RefID WHERE ItemID='" & strItemCode & "' AND CustCode<>'Stock'", con, adOpenForwardOnly, adLockReadOnly
        LVComplaints.ListItems.Clear
        Do Until .EOF
            Set ITM = LVComplaints.ListItems.add(, , !ComplaintNo & "")
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !Description & ""
            .MoveNEXT
        Loop
        .Close
        
        txtMinStock = GetSingleLongValue("MinLevel", "Items", "WHERE ItemID='" & strItemCode & "'")
        
    End With
    Set rs = Nothing
    Call GetTotal
    Call CalculateCompletionAvg
    
End Sub

Private Sub CalculateCompletionAvg()

    Dim rs As New ADODB.Recordset, strItemID As String, lTotalProcesses As Long, lTotalOrdQty As Long
    strItemID = cmbArticle.ID
    'lTotalOrdQty = Val(txtTotalCustOrders)
    
    Dim dPercentCompletion As Double, lSNo As Long, lQty As Long
    Dim dTotalCompletion As Double
    With rs
        .Open "SELECT COUNT(*) AS TotalProcesses FROM ItemProcesses WHERE ItemID='" & strItemID & "'", con, adOpenForwardOnly, adLockReadOnly
        lTotalProcesses = Val(!TotalProcesses & "")
        .Close
        
        lTotalOrdQty = GetSingleLongValue("SUM(RcvdQty)", "VendRcvdDetail", " WHERE ItemCode='" & strItemID & "' AND EntryID IN(SELECT MAX(EntryID) FROM VendRcvdDetail GROUP BY LotNo)")
        
        If lTotalOrdQty = 0 Then
            dTotalCompletion = 0
        Else
            .Open "SELECT SNo,SUM(RcvdQty) AS CompletedQty FROM VendRcvdDetail INNER JOIN ItemProcesses " & _
                "ON VendRcvdDetail.ProcessID=ItemProcesses.ProcessID AND VendRcvdDetail.ItemCode=ItemProcesses.ItemID " & _
                "WHERE ItemID='" & strItemID & "' AND VendRcvdDetail.EntryID IN(SELECT MAX(EntryID) FROM VendRcvdDetail GROUP BY LotNo) " & _
                "GROUP BY SNo"
            
            Do Until .EOF
                lQty = Val(!CompletedQty & "")
                lSNo = Val(!SNo & "")
                
                dPercentCompletion = (lSNo / lTotalProcesses) * 100
                dPercentCompletion = (lQty / lTotalOrdQty) * dPercentCompletion
                dTotalCompletion = dTotalCompletion + dPercentCompletion
                
                .MoveNEXT
            Loop
            .Close
        End If
    End With
    Set rs = Nothing
    
    txtCompletion = Round(dTotalCompletion, 2) & "%"
    
End Sub

Private Sub FilterByOrderNo(Optional strOrderNo As String)

    Call RefreshCustomerOrders(strOrderNo)
    
    Call RefreshRunningLots(strOrderNo)
        
    Call RefreshOpenPOs(strOrderNo)
    
    Call GetTotal
    
End Sub

Private Sub RefreshCustomerOrders(Optional strOrderNo As String)
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim strCondition As String
    If strOrderNo <> "" Then
        strCondition = " AND OrderNo='" & strOrderNo & "'"
    End If
    
    Dim lForeColor As Long
    With rs
        'Refresh Customer Orders
        If strCompany = "QEL" Then
            .Open "SELECT * FROM VrptOrders WHERE ItemCode='" & cmbArticle.ID & "' AND OrderNo NOT IN('Stock-Order','Stock-OrderIss') AND OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList) AND Qty>ShippedQty " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT * FROM VFOrderItems WHERE CompItemCode='" & cmbArticle.ID & "' AND OrderNo NOT IN('Stock-Order','Stock-OrderIss') AND OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList) AND Qty>ShippedQty " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        End If
        LVCustOrders.ListItems.Clear
        Dim strRemarks As String, i As Integer
        Dim strOrderNoToShow As String
        Do Until .EOF
        
            If !Remarks & "" = "" Then
                strRemarks = ""
            Else
                strRemarks = !Remarks & ""
            End If
            
            If Val(!DeliveryStatus & "") = 2 Then
                If strRemarks <> "" Then
                    strRemarks = strRemarks & "/"
                End If
                strRemarks = strRemarks & "Urgent"
            End If
            If strCompany = "IAA" Or strCompany = "Kami" Then
                strOrderNoToShow = !InternalRefNo & ""
            Else
                strOrderNoToShow = !OrderNo & ""
            End If
            Set ITM = LVCustOrders.ListItems.add(, , !CustCode & "")
            ITM.Tag = !OrderNo & ""
            
            ITM.ListSubItems.add(, , strOrderNoToShow).Tag = Val(!ID & "")
            ITM.ListSubItems.add(, , Format(!DT, "dd-MM-yy")).Tag = Val(!ID & "")
            ITM.ListSubItems.add , , Val(!Qty & "") - Val(!ShippedQty & "")
            ITM.ListSubItems.add , , Format(!DeliveryDT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , strRemarks
            
            If strRemarks = "" Then
                lForeColor = LVCustOrders.ForeColor
            Else
                lForeColor = vbRed
            End If
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Public Sub RefreshRunningLots(Optional strOrderNo As String)

    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim strCondition As String
    If strOrderNo <> "" Then
        strCondition = " AND TOrders.OrderNo='" & strOrderNo & "'"
    End If
    Dim strR2InDT As String, strFrom As String
    Dim strOrderNoToShow As String
    With rs
        
        .Open "SELECT T1.*,VLotWithFirstProcessOfIssuance.Description AS ProcessFrom,VLotWithR2InDate.DT AS R2InDT,TSOOI.VID_RefID AS SOO_ISS_EntryID,IssuanceEntryID,FCustomerOrders.CustCode FROM " & _
            "(SELECT ItemCode,LotNo,OrderNo,Description,InternalRefNo,IssQty-RcvdQty AS Qty,ProcessID,0 AS SOA_EntryID,EntryID AS IssuanceEntryID,0 AS RcvEntryID FROM VVendIssdDetail_ForRunningLots TOrders " & _
            "WITH (NOLOCK) WHERE ItemCode='" & cmbArticle.ID & "' AND LotNo<>'0' AND IssQty-RcvdQty>0 AND " & _
            "EntryID NOT IN(SELECT Issue_RefID FROM VendRcvdDetail) " & strCondition & _
            " UNION " & _
            "SELECT ItemCode,LotNo,TOrders.OrderNo,Description,InternalRefNo,RcvdQty-IssQty-ISNULL(Wastage,0)-ISNULL(ReWorkQty,0) AS Qty,ProcessID,0 AS SOA_EntryID,0 AS IssuanceEntryID,TOrders.EntryID AS RcvEntryID  " & _
            "FROM VVendRcvdDetail_Simple TOrders WITH (NOLOCK) " & _
            " WHERE ItemCode='" & cmbArticle.ID & "' AND LotNo<>'0' " & _
            "AND RcvdQty-IssQty-ISNULL(Wastage,0)-ISNULL(ReWorkQty,0)>0 " & _
            "AND TOrders.EntryID NOT IN(SELECT Rcvd_RefID FROM VendIssdDetail WHERE Rcvd_RefID IS NOT NULL) AND ISNULL(Opening_RefID,0)=0 " & strCondition & _
            ") T1 LEFT OUTER JOIN VLotWithR2InDate ON T1.LotNo=VLotWithR2InDate.LotNo " & _
            "LEFT OUTER JOIN VLotWithFirstProcessOfIssuance ON T1.LotNo=VLotWithFirstProcessOfIssuance.LotNo " & _
            "LEFT OUTER JOIN (SELECT VID_RefID FROM StockOrderOpening_Issuance GROUP BY VID_RefID) TSOOI ON VLotWithFirstProcessOfIssuance.EntryID=TSOOI.VID_RefID " & _
            "LEFT OUTER JOIN FCustomerOrders ON T1.OrderNo=FCustomerOrders.OrderNo" & _
            " WHERE T1.LotNo NOT IN(SELECT LotNo FROM Lots_Closed) ORDER BY T1.ProcessID", con, adOpenForwardOnly, adLockReadOnly
            
        LVRunningLots.ListItems.Clear
        Do Until .EOF
            Dim strLotNo As String
            If strCompany = "Dr-Frgz" And strSubCompany = "Electro" Then
                strLotNo = GetSingleLongValue("InternalProductionOrderDetailLots.EntryID", "InternalProductionOrderDetailLots INNER JOIN VendIssdDetail ON InternalProductionOrderDetailLots.EntryID=VendIssdDetail.IPODL_RefID", " WHERE VendIssdDetail.EntryID=" & Val(!IssuanceEntryID & ""))
            Else
                strLotNo = !LotNo & ""
            End If
            Set ITM = LVRunningLots.ListItems.add(, , strLotNo)
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            If strCompany = "IAA" Or strCompany = "Kami" Then
                strOrderNoToShow = !InternalRefNo & ""
            Else
                strOrderNoToShow = !OrderNo & ""
            End If
            ITM.ListSubItems.add , , strOrderNoToShow
            ITM.ListSubItems.add , , !CustCode & ""
            If IsNull(!R2InDT) Then
                strR2InDT = "-"
            Else
                strR2InDT = Format(!R2InDT, "dd-MMM-yyyy")
            End If
            ITM.ListSubItems.add(, , strR2InDT).Tag = Val(!RcvEntryID & "")
            If Val(!SOO_ISS_EntryID & "") > 0 Then
                strFrom = "Semi"
            ElseIf Val(!SOA_EntryID & "") > 0 Then
                strFrom = "Pack"
            Else
                strFrom = ""
            End If
            ITM.ListSubItems.add(, , strFrom).Tag = Val(!IssuanceEntryID & "")
            .MoveNEXT
        Loop
        .Close
    End With
    
End Sub


Private Sub RefreshOpenPOs(Optional strOrderNo As String)

    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim strCondition As String
    strCondition = " AND Closed=0 "
    If strOrderNo <> "" Then
        strCondition = strCondition & " AND OrderNo='" & strOrderNo & "'"
    End If
    Dim strOrderNoToShow As String
    With rs
        '.Open "SELECT VendIssued.EntryID,VendID1,VenderName,VendIssued.RecieptID,MasterPONo,OrderNo,VendIssdDetail_ReturnDTs.IssQty+CASE WHEN ISNULL(IssRcvBalance,0)<0 THEN ABS(ISNULL(IssRcvBalance,0)) ELSE 0 END+ISNULL(Wastage_Return,0)-VendIssdDetail_ReturnDTs.RcvQty AS QtyOld,VendIssdDetail_ReturnDTs.IssQty-VendIssdDetail_ReturnDTs.RcvQty+ISNULL(WastageQty,0) AS Qty,VendIssdDetail_ReturnDTs.ReturnDT " & _
         "FROM VendIssued INNER JOIN VendIssdDetail ON VendIssued.EntryID=VendIssdDetail.RefID " & _
         "INNER JOIN VendIssdDetail_ReturnDTs ON VendIssdDetail.EntryID=VendIssdDetail_ReturnDTs.VIS_RefID " & _
         "INNER JOIN Makers ON VendIssued.VendID=Makers.VendID LEFT OUTER JOIN VendIssdDetail_MoreDetails ON VendIssdDetail.EntryID=VendIssdDetail_MoreDetails.VID_RefID " & _
         "LEFT OUTER JOIN VIssuanceEntryWithWastage ON VendIssdDetail.EntryID=VIssuanceEntryWithWastage.EntryID " & _
         "WHERE VendIssdDetail.ItemCode='" & cmbArticle.ID & "' AND LotNo='0' AND VendIssdDetail_ReturnDTs.IssQty+CASE WHEN ISNULL(IssRcvBalance,0)<0 THEN ABS(ISNULL(IssRcvBalance,0)) ELSE 0 END+ISNULL(Wastage_Return,0)>VendIssdDetail_ReturnDTs.RcvQty " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM VItemOpenPOs WHERE ItemCode='" & cmbArticle.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LVOpenPO.ListItems.Clear
        Do Until .EOF
            Set ITM = LVOpenPO.ListItems.add(, , !RecieptID & "")
            ITM.Tag = !EntryID
            ITM.ListSubItems.add , , !VenderName & "" '!VendID1 & ""
            
            strOrderNoToShow = !OrderNo & ""
            If strCompany = "IAA" Or strCompany = "Kami" Then
                strOrderNoToShow = GetSingleStringValue("InternalRefNo", "FCustomerOrders", " WHERE OrderNo='" & strOrderNoToShow & "'")
            End If
            
            ITM.ListSubItems.add , , strOrderNoToShow
            
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add(, , Format(!ReturnDT, "dd-MMM-yy")).Tag = !MasterPONo
            ITM.ListSubItems.add , , !MasterPONo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub GetRunningLotsTotal()

    Dim lRunLotsQty  As Long, i As Integer
    lRunLotsQty = 0
    For i = 1 To LVRunningLots.ListItems.count
        With LVRunningLots.ListItems(i)
            lRunLotsQty = lRunLotsQty + Val(.SubItems(2))
        End With
    Next
    txtTotalRunningLots = lRunLotsQty
    
End Sub

Private Sub RefreshForging()
    
    Dim rs As New ADODB.Recordset, ITM As ListItem, lForeColor As Long, i As Integer
    With rs
        .Open "SELECT * FROM VMaterialLocationWiseBalances WHERE RMID IN(SELECT RMID FROM ItemsRMComp WHERE ItemID='" & cmbArticle.ID & "') AND ShelfQty>0", con, adOpenForwardOnly, adLockReadOnly
        LVForging.ListItems.Clear
        Do Until .EOF
            Set ITM = LVForging.ListItems.add(, , !MaterialID & "")
            ITM.ListSubItems.add , , Val(!ShelfQty & "")
            ITM.ListSubItems.add , , !StoreName & " - " & !RackNo & " - " & !ShelfNo & ""
            
            ITM.ListSubItems(1).Tag = Val(!RMID & "")
            If GetSingleLongValue("Functional_Status", "ItemsRMComp", " WHERE ItemID='" & cmbArticle.ID & "' AND RMID=" & Val(!RMID & "")) = 1 Then 'In-Active for Issuance
                lForeColor = vbRed
            Else
                lForeColor = LVForging.ForeColor
            End If
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
            .MoveNEXT
        Loop
        .Close
    End With
    
End Sub

Private Sub PrintRMLinkedReport()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PPDashBoard_RMLinkedLots.rpt")
    Dim strSelection As String
    strSelection = "{ItemsRMComp.ItemID}='" & cmbArticle.ID & "' AND {VMaterialLocationWiseBalances.ShelfQty}>0"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
End Sub


Private Sub PrintFinishStoreLots()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PPDashBoard_FinishStoreLots.rpt")
    Dim strSelection As String
    strSelection = "{VItemsWithShelfWiseStock.ItemID}='" & cmbArticle.ID & "'"
    rpt.GroupSelectionFormula = "Sum ({VItemsWithShelfWiseStock.NetQty}, {@ForGrouping})>0"
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub PrintOrderStatusReport(strOrderNo As String)

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrder_Status.rpt")

    Dim strSelection As String
    strSelection = "{CustOrders.OrderNo}='" & strOrderNo & "' "
    
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrderNo

    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"

    rpt.EnableParameterPrompting = False
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt, "", False
    
End Sub

