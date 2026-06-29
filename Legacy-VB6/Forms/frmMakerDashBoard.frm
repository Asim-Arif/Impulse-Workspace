VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerDashBoard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "fv"
   ClientHeight    =   10560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMakerDashBoard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrintEvaluation 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Evaluation"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10110
      Width           =   1545
   End
   Begin VB.CommandButton cmdPrintRateList 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print Rate &List"
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
      TabIndex        =   17
      Top             =   10110
      Width           =   1545
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
      ScaleWidth      =   12120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   12120
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Height          =   9495
         Left            =   30
         TabIndex        =   3
         Top             =   -75
         Width           =   12060
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Upcoming delivery promises :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2940
            Left            =   7665
            TabIndex        =   26
            Top             =   3615
            Width           =   4305
            Begin MSComctlLib.ListView LVPromises 
               Height          =   2595
               Left            =   60
               TabIndex        =   27
               Top             =   270
               Width           =   4170
               _ExtentX        =   7355
               _ExtentY        =   4577
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   16711680
               BackColor       =   16761024
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "S#"
                  Object.Width           =   1058
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Detail"
                  Object.Width           =   5644
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Hidden"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Hidden"
                  Object.Width           =   0
               EndProperty
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Pending"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3105
            Left            =   45
            TabIndex        =   18
            Top             =   6330
            Width           =   11955
            Begin VB.TextBox txtPendingItems 
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
               Left            =   3975
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   2655
               Width           =   1140
            End
            Begin VB.TextBox txtTotalValuePending 
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
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   2655
               Width           =   1140
            End
            Begin VB.TextBox txtPOQty 
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
               Left            =   6315
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   2655
               Width           =   1140
            End
            Begin VB.TextBox txtRcvdQty 
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
               Left            =   7485
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   2655
               Width           =   1140
            End
            Begin VB.TextBox txtBalance 
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
               Left            =   9765
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   2655
               Width           =   1140
            End
            Begin MSComctlLib.ListView LVBalances 
               Height          =   2355
               Left            =   45
               TabIndex        =   22
               Top             =   285
               Width           =   11835
               _ExtentX        =   20876
               _ExtentY        =   4154
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   16711680
               BackColor       =   16761024
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   11
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Customer"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Customer Order"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Process"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Item"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "PO Date"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "PO Qty"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "T. Rcvd"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Rcv. Date"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "Balance"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Item Descriptoin"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "PO #"
                  Object.Width           =   3281
               EndProperty
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No. of Items :"
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
               Index           =   2
               Left            =   2565
               TabIndex        =   31
               Top             =   2715
               Width           =   1365
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Value :"
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
               Left            =   75
               TabIndex        =   30
               Top             =   2715
               Width           =   1275
            End
            Begin VB.Label Label3 
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
               ForeColor       =   &H0000C000&
               Height          =   285
               Index           =   0
               Left            =   5520
               TabIndex        =   23
               Top             =   2715
               Width           =   660
            End
         End
         Begin ComboList.Usercontrol1 cmbMaker 
            Height          =   330
            Left            =   60
            TabIndex        =   6
            Top             =   360
            Width           =   10785
            _ExtentX        =   19024
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
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   5355
            TabIndex        =   4
            Top             =   3405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin MSComctlLib.ListView LVDetails 
            Height          =   2790
            Left            =   2910
            TabIndex        =   5
            Top             =   750
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   4921
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   16711680
            BackColor       =   16761024
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Receivings"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            Left            =   45
            TabIndex        =   13
            Top             =   3540
            Width           =   7590
            Begin VB.TextBox txtTotalReWork 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   360
               Left            =   6375
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   2370
               Width           =   1000
            End
            Begin VB.TextBox txtTotalWastage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   360
               Left            =   5355
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   2370
               Width           =   1000
            End
            Begin VB.TextBox txtTotalValueRcvd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   360
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   2370
               Width           =   1635
            End
            Begin VB.TextBox txtTotalRcvd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00404040&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   360
               Left            =   4140
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   2370
               Width           =   1200
            End
            Begin MSComctlLib.ListView LVRcvList 
               Height          =   1830
               Left            =   60
               TabIndex        =   14
               Top             =   525
               Width           =   7485
               _ExtentX        =   13203
               _ExtentY        =   3228
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   16711680
               BackColor       =   16761024
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Item"
                  Object.Width           =   7056
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Qty Rcvd"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "Rej"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "Re-Work"
                  Object.Width           =   1764
               EndProperty
            End
            Begin MSComCtl2.DTPicker DTRcv 
               Height          =   300
               Left            =   2430
               TabIndex        =   15
               Top             =   195
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "MMMM, yyyy"
               Format          =   197459971
               CurrentDate     =   40597
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Value :"
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
               Left            =   60
               TabIndex        =   29
               Top             =   2385
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Month :"
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
               Left            =   1710
               TabIndex        =   16
               Top             =   210
               Width           =   660
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Accounts Details :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2940
            Left            =   7665
            TabIndex        =   8
            Top             =   645
            Width           =   4305
            Begin MSComctlLib.ListView LV 
               Height          =   2370
               Left            =   60
               TabIndex        =   10
               Top             =   525
               Width           =   4170
               _ExtentX        =   7355
               _ExtentY        =   4180
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   16711680
               BackColor       =   16761024
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Vchr"
                  Object.Width           =   2381
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Debit"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Text            =   "Credit"
                  Object.Width           =   2117
               EndProperty
            End
            Begin MSComCtl2.DTPicker DT 
               Height          =   330
               Left            =   2550
               TabIndex        =   11
               Top             =   150
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   582
               _Version        =   393216
               CustomFormat    =   "MMMM, yyyy"
               Format          =   197459971
               CurrentDate     =   40597
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Month :"
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
               Left            =   1890
               TabIndex        =   12
               Top             =   195
               Width           =   660
            End
         End
         Begin MSForms.CommandButton cmdRefresh 
            Height          =   420
            Left            =   10890
            TabIndex        =   9
            Top             =   270
            Width           =   1125
            ForeColor       =   65280
            BackColor       =   4210752
            Caption         =   "Refresh        "
            PicturePosition =   327683
            Size            =   "1984;741"
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
            Caption         =   "Maker"
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
            TabIndex        =   7
            Top             =   150
            Width           =   525
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   2790
            Left            =   45
            Stretch         =   -1  'True
            Top             =   750
            Width           =   2850
         End
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAKER DASHBOARD"
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
      Width           =   12120
   End
End
Attribute VB_Name = "frmMakerDashBoard"
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


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrintEvaluation_Click()

    Dim DtFrom As Date, DtTo As Date
    Dim frm As New frmGetDateDouble
    Load frm
    If frm.getDate(DtFrom, DtTo) = False Then Exit Sub
    
    Dim rpt  As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\SupplierEvaluation.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VVendReceivingList.OnlyDT}=#" & DtFrom & "# TO #" & DtTo & "# AND {VVendReceivingList.VendID}=" & cmbMaker.ID, rpt
    
End Sub

Private Sub cmdPrintRateList_Click()
    On Error GoTo err
    If cmbMaker.MatchFound = False Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerAssignedItems.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendAssItems.VendID}=" & cmbMaker.ID, rpt
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()

    On Error GoTo err
    
    If cmbMaker.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Sub
    End If
    Dim strAccNo As String
    Dim lTotalTaken As Long, lTotalCleared As Long, lDeductionAmt As Long, lAdvance As Long, lBalanceLoan As Long
    Dim rs As New ADODB.Recordset, ITM As ListItem, DTLastMonth As Date, lCurrentBill As Long
    With rs
        .Open "SELECT * FROM Makers WHERE VendID=" & cmbMaker.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            If Not IsNull(![VendPic]) Then
                Set Image1.DataSource = rs
                Image1.DataField = "VendPic"
            Else
                Image1.Picture = LoadPicture("")
            End If
            strAccNo = !AccNo
            LVDetails.ListItems("MakerName").SubItems(1) = !VenderName & ""
            LVDetails.ListItems("Address").SubItems(1) = !Address & ""
            LVDetails.ListItems("PhoneNo").SubItems(1) = !Phone1 & ""
            LVDetails.ListItems("LastPayment").SubItems(1) = Format(GetSingleLongValue("TOP 1 Debit", "Vouchers", " WHERE AccNo='" & !AccNo & "' AND LEFT(VchrNo,3) IN('CPV','BPV') ORDER BY VDate DESC"), "#,##")
            lCurrentBill = getCurrentBillAmount(cmbMaker.ID)
            LVDetails.ListItems("CurrentBill").SubItems(1) = Format(lCurrentBill, "#,##")
'            lTotalTaken = GetSingleLongValue("TotalAmountTaken", "VMakerAdvancesDeduction", " WHERE AccNo='" & strAccNo & "'")
'            lTotalCleared = GetSingleLongValue("TotalAmountCleared", "VMakerAdvancesDeduction", " WHERE AccNo='" & strAccNo & "'")
            lBalanceLoan = GetSingleLongValue("TotalAmountTaken-TotalAmountCleared", "VMakerAdvancesDeduction", " WHERE AccNo='" & strAccNo & "'")
            lDeductionAmt = GetSingleLongValue("ActualDeductionAmount", "VMakerAdvancesDeduction", " WHERE AccNo='" & strAccNo & "'")
            lAdvance = GetSingleLongValue("Amount-AmountCleared", "VMakerAdvancesShortDeduction", " WHERE AccNo='" & strAccNo & "'")
            LVDetails.ListItems("Loan").SubItems(1) = Format(lBalanceLoan, "#,##")
            LVDetails.ListItems("Deduction").SubItems(1) = Format(lDeductionAmt, "#,##")
            LVDetails.ListItems("Advance").SubItems(1) = Format(lAdvance, "#,##")
            
            LVDetails.ListItems("LedgerBalance").SubItems(1) = Format(GetSingleLongValue("Balance", "VAccounts", " WHERE AccNo='" & !AccNo & "'"), "#,## DR;#,## CR;#,##")
            
        End If
        .Close
        
        Dim strLastRcvDT As String
        Dim lPOQty As Long, lRcvdQty As Long, lBalanceQty As Long
        lPOQty = 0
        lRcvdQty = 0
        lBalanceQty = 0
        
        Dim strCondition As String, strOrderNoToShow As String
        strCondition = "WHERE VendID=" & cmbMaker.ID & " AND RcvdQty<IssQty AND Closed=0 AND CustCode NOT IN('Stock')"
        .Open "SELECT * FROM VVendIssdDetail_MakerDashBoard " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        LVBalances.ListItems.Clear
        
        Do Until .EOF
        
            Set ITM = LVBalances.ListItems.add(, , !CustCode & "")
            
            If strCompany = "IAA" Or strCompany = "Kami" Then
                strOrderNoToShow = !InternalRefNo & ""
            Else
                strOrderNoToShow = !OrderNo & ""
            End If
            ITM.ListSubItems.add , , strOrderNoToShow
            
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MM-yy")
            ITM.ListSubItems.add , , Val(!IssQty & "")
            ITM.ListSubItems.add , , Val(!RcvdQty & "")
            
            If IsNull(!RcvDT) Then
                ITM.ListSubItems.add , , "-"
            Else
                ITM.ListSubItems.add , , Format(!RcvDT, "dd-MM-yy")
            End If
            
            ITM.ListSubItems.add , , Val(!IssQty & "") - Val(!RcvdQty & "")
            
            ITM.ListSubItems.add , , !ItemName & ""
            
            ITM.ListSubItems.add , , !MainRecieptID & ""
            
            lPOQty = lPOQty + Val(!IssQty & "")
            lRcvdQty = lRcvdQty + Val(!RcvdQty & "")
            lBalanceQty = lBalanceQty + Val(!IssQty & "") - Val(!RcvdQty & "")
            
            .MoveNEXT
        Loop
        .Close
         
        txtPOQty = Format(lPOQty, "#,##")
        txtRcvdQty = Format(lRcvdQty, "#,##")
        txtBalance = Format(lBalanceQty, "#,##")
        
        .Open "SELECT SUM((IssQty-RcvdQty)*Rate) AS PendingValue,COUNT(DISTINCT ItemCode) AS PendingItems FROM VVendIssdDetail_MakerdashBoard " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        txtPendingItems = Format(Val(!PendingItems & ""), "#,##")
        txtTotalValuePending = Format(Val(!PendingValue & ""), "#,##")
        .Close
        
    End With
    
    Call DT_Change
    Call DTRcv_Change
    
    Call UpdatePromises
    Exit Sub
err:
    MsgBox err.Description
End Sub
    
Private Sub UpdatePromises()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strAlert As String, dBalance As Double, lForeColor As Long, i As Integer
    With rs
        .Open "SELECT EntryIDReturnDTs,ItemName,ItemCode,IssQty,RcvQty,Prefix,FollowUpData.EntryID AS FollowUpEntryID,MasterPONO FROM " & _
        "VPurchaseCalendar INNER JOIN Items ON VPurchaseCalendar.ItemCode=Items.ItemID " & _
        "LEFT OUTER JOIN FollowUpData ON EntryIDReturnDTs=VIDRDTS_RefID " & _
        "AND Prefix=FollowupData.FollowUp WHERE IssQty>RcvQty " & _
        "AND CustCode<>'Stock' AND MasterPONo IS NOT NULL AND VPurchaseCalendar.VendID=" & cmbMaker.ID & " AND ReturnDT BETWEEN '" & Date & "' AND '" & DateAdd("d", 15, Date) & "'", con, adOpenForwardOnly, adLockReadOnly
        
        LVPromises.ListItems.Clear
        Do Until .EOF
        
            Set ITM = LVPromises.ListItems.add(, , LVPromises.ListItems.count + 1)
            ITM.Tag = Val(!EntryIDReturnDTs & "")
            
            strAlert = ""
            If !Prefix <> "" Then
                strAlert = !Prefix & ""
            End If
            dBalance = Val(!IssQty & "") - Val(!RcvQty & "")
            'strAlert = strAlert & "(" & !ItemCode & ") " & !ItemName & " / " & dBalance & " / " & "(" & !VendID1 & ") " & !VenderName
            strAlert = strAlert & " " & dBalance & " Pcs of " & !ItemCode & " (" & !ItemName & ")"
            ITM.ListSubItems.add(, , strAlert).Tag = !Prefix & ""
            ITM.ListSubItems.add(, , "").Tag = Val(!FollowUpEntryID & "")
            ITM.ListSubItems.add(, , "").Tag = !MasterPONo & ""
            
            If Val(!FollowUpEntryID & "") = 0 Then
                If !Prefix & "" = "(D)" Then
                    lForeColor = RGB(50, 50, 255)
                Else
                    lForeColor = RGB(255, 50, 50)
                End If
            Else
                lForeColor = RGB(100, 200, 100)
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
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function getCurrentBillAmount(lMakerID As Long) As Long

    On Error GoTo err
    Dim cmd As New ADODB.Command, dBillAmount As Double, rs As New ADODB.Recordset
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "MakerBilling_RcvingWise_SP"
        .Parameters("@MakerID").Value = lMakerID
        .Parameters("@DTFrom").Value = #1/1/2005#
        .Parameters("@DTTo").Value = DateAdd("yyyy", 10, Date)
        .Parameters("@CatID").Value = "0"
        Set rs = .Execute
    End With
    
    dBillAmount = 0
    With rs
        Do Until .EOF
            dBillAmount = dBillAmount + Round(Val(!BillingProcAmount & ""), 4)
            .MoveNEXT
        Loop
        .Close
    End With
    getCurrentBillAmount = dBillAmount
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub DT_Change()
    If cmbMaker.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset, strAccNo As String, ITM As ListItem
    strAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbMaker.ID)
    With rs
        .Open "SELECT * FROM VLedger WHERE AccNo='" & strAccNo & "' AND MONTH(VDate)=" & DT.Month & " AND YEAR(VDate)=" & DT.year, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !VchrNo & "")
            ITM.ListSubItems.add , , Format(Val(!Debit & ""), "#,##0")
            ITM.ListSubItems.add , , Format(Val(!Credit & ""), "#,##0")
            ITM.ToolTipText = !Description & ""
            ITM.ListSubItems(1).ToolTipText = !Description & ""
            ITM.ListSubItems(2).ToolTipText = !Description & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub DTRcv_Change()

    If cmbMaker.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset, strAccNo As String, ITM As ListItem
    strAccNo = GetSingleStringValue("AccNo", "Makers", " WHERE VendID=" & cmbMaker.ID)
    Dim lTotal As Long, dTotalValue As Double
    Dim lTotalWastage As Long, lTotalReWork As Long
    With rs
        .Open "SELECT ItemCode,ItemName,SUM(RcvdQty) AS RcvdQty,SUM(RcvdQty*Rate) AS RcvdValue,SUM(Wastage) AS Wastage,SUM(ReWorkQty) AS ReWorkQty FROM VVendRcvdDetail WHERE VendID=" & cmbMaker.ID & " AND MONTH(DT)=" & DTRcv.Month & " AND YEAR(DT)=" & DTRcv.year & " GROUP BY ItemCode,ItemName ORDER BY SUM(RcvdQty)", con, adOpenForwardOnly, adLockReadOnly
        LVRcvList.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRcvList.ListItems.add(, , !ItemCode & " " & !ItemName & "")
            ITM.ListSubItems.add , , Val(!RcvdQty & "")
            ITM.ListSubItems.add , , Val(!Wastage & "")
            ITM.ListSubItems.add , , Val(!ReWorkQty & "")
            lTotal = lTotal + Val(!RcvdQty & "")
            lTotalWastage = lTotalWastage + Val(!Wastage & "")
            lTotalReWork = lTotalReWork + Val(!ReWorkQty & "")
            dTotalValue = dTotalValue + Val(!RcvdValue & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    txtTotalRcvd = Format(lTotal, "#,##")
    txtTotalValueRcvd = Format(dTotalValue, "#,##")
    
    txtTotalWastage = Format(lTotalWastage, "#,##")
    txtTotalReWork = Format(lTotalReWork, "#,##")
End Sub

Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
        Else
            LV.SortOrder = lvwAscending
        End If
    Else
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
    End If
    If LV.Sorted = False Then LV.Sorted = True
    
End Sub

Private Sub LVBalances_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LVBalances.SortKey + 1 = ColumnHeader.Index Then
        If LVBalances.SortOrder = lvwAscending Then
            LVBalances.SortOrder = lvwDescending
        Else
            LVBalances.SortOrder = lvwAscending
        End If
    Else
        LVBalances.SortKey = ColumnHeader.Index - 1
        LVBalances.SortOrder = lvwAscending
    End If
    If LVBalances.Sorted = False Then LVBalances.Sorted = True
    
End Sub

Private Sub LVRcvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LVRcvList.SortKey + 1 = ColumnHeader.Index Then
        If LVRcvList.SortOrder = lvwAscending Then
            LVRcvList.SortOrder = lvwDescending
        Else
            LVRcvList.SortOrder = lvwAscending
        End If
    Else
        LVRcvList.SortKey = ColumnHeader.Index - 1
        LVRcvList.SortOrder = lvwAscending
    End If
    If LVRcvList.Sorted = False Then LVRcvList.Sorted = True
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

Public Function ShowMe() As Boolean
    
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
    
    DT.Value = Date
    DTRcv.Value = Date
    
    cmbMaker.ListHeight = 5800
    Call cmbMaker.AddVals(con, "'{' + VendID1 + '} ' + VenderName", "Makers", "VendID", " ORDER BY VenderName")
    
    LVDetails.ListItems.add(, "MakerName", "MakerName").ForeColor = vbWhite
    LVDetails.ListItems.add(, "Address", "Address").ForeColor = vbWhite
    LVDetails.ListItems.add(, "PhoneNo", "Phone").ForeColor = vbWhite
    LVDetails.ListItems.add , "", ""
    
    LVDetails.ListItems.add(, "LastPayment", "Last Payment").ForeColor = vbWhite
    LVDetails.ListItems.add(, "CurrentBill", "Current Bill").ForeColor = vbWhite
    LVDetails.ListItems.add(, "Loan", "Loan").ForeColor = vbWhite
    LVDetails.ListItems.add(, "Deduction", "Deduction").ForeColor = vbWhite
    
    LVDetails.ListItems.add(, "Advance", "Advance").ForeColor = vbWhite
    LVDetails.ListItems.add , "", ""
    LVDetails.ListItems.add(, "LedgerBalance", "Ledger Balance").ForeColor = vbWhite
    Me.Show
    ShowMe = bSaved
    
End Function

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

