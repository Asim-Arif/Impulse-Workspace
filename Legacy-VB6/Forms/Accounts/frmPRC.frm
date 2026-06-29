VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmPRC 
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPRC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0F3EB&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   7905
      Index           =   0
      Left            =   0
      ScaleHeight     =   7905
      ScaleWidth      =   7635
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "PRC Entry"
      Top             =   0
      Width           =   7635
      Begin VB.TextBox txtNetAmt 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   5790
         Width           =   1860
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   1695
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.TextBox txtPostedAmt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0F3EB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1815
      End
      Begin VB.TextBox txtForeignExchDiff 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2295
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5790
         Width           =   1860
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3300
         TabIndex        =   30
         Top             =   4515
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txtRunningSerialNo 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   1350
         Width           =   2055
      End
      Begin VB.TextBox txtBillNo 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   3585
         TabIndex        =   3
         Top             =   1350
         Width           =   2055
      End
      Begin VB.TextBox txtAmtInPakRs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0F3EB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1815
      End
      Begin VB.TextBox txtExchRate 
         Height          =   330
         Left            =   2835
         TabIndex        =   10
         Top             =   2700
         Width           =   1020
      End
      Begin VB.TextBox txtAmtRealized 
         Height          =   330
         Left            =   1515
         TabIndex        =   14
         Top             =   2700
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker DTRealization 
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   2700
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   127401987
         CurrentDate     =   39644
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0F3EB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5670
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1350
      End
      Begin MSComctlLib.ListView LV 
         Height          =   2235
         Left            =   30
         TabIndex        =   15
         Top             =   3465
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   32768
         BackColor       =   14742507
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S/N"
            Object.Width           =   1085
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Detail of Deduction"
            Object.Width           =   8361
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2593
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Curr"
            Object.Width           =   1244
         EndProperty
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0F3EB&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   4980
         Picture         =   "frmPRC.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6810
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0F3EB&
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
         Height          =   1020
         Left            =   6300
         Picture         =   "frmPRC.frx":09A9
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6810
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTBillNo 
         Height          =   330
         Left            =   5670
         TabIndex        =   4
         Top             =   1350
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   127401987
         CurrentDate     =   39644
      End
      Begin MSComCtl2.DTPicker DTSerialNo 
         Height          =   330
         Left            =   2220
         TabIndex        =   2
         Top             =   1350
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   72613891
         CurrentDate     =   39644
      End
      Begin MSComCtl2.DTPicker DTPosting 
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   7455
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   72613891
         CurrentDate     =   39644
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   16
         Left            =   4500
         TabIndex        =   38
         Top             =   5835
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posted Amount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   15
         Left            =   5715
         TabIndex        =   37
         Top             =   2460
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Foreign Exch. Difference :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   14
         Left            =   120
         TabIndex        =   34
         Top             =   5820
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   13
         Left            =   150
         TabIndex        =   33
         Top             =   7215
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   32
         Top             =   6105
         Width           =   1245
      End
      Begin MSForms.ComboBox cmbBanks 
         Height          =   390
         Left            =   135
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   6330
         Width           =   7410
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13070;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   11
         Left            =   2220
         TabIndex        =   29
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Running Serial No. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   28
         Top             =   1110
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   9
         Left            =   5670
         TabIndex        =   27
         Top             =   1110
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Number :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   8
         Left            =   3585
         TabIndex        =   26
         Top             =   1110
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt In Rs. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   7
         Left            =   3900
         TabIndex        =   25
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exch. Rate :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   6
         Left            =   2835
         TabIndex        =   23
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt Realised :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   5
         Left            =   1515
         TabIndex        =   22
         Top             =   2475
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Realiz. Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9EBCC&
         BackStyle       =   0  'Transparent
         Caption         =   "DEDUCTIONS:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   30
         TabIndex        =   20
         Top             =   3090
         Width           =   2460
      End
      Begin VB.Image ImgEuro 
         Height          =   480
         Left            =   7080
         Picture         =   "frmPRC.frx":1208
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgDollar 
         Height          =   480
         Left            =   7080
         Picture         =   "frmPRC.frx":176B
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. Amount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   3
         Left            =   5670
         TabIndex        =   19
         Top             =   1815
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   18
         Top             =   1815
         Width           =   900
      End
      Begin MSForms.ComboBox cmbCustomInvoice 
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   2040
         Width           =   1500
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2646;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   17
         Top             =   1815
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
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   0
         Left            =   1710
         TabIndex        =   16
         Top             =   1815
         Width           =   915
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   3780
         TabIndex        =   6
         Top             =   2040
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
      Begin VB.Image Image2 
         Height          =   960
         Left            =   6555
         Picture         =   "frmPRC.frx":1CAB
         Top             =   60
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   15
         Picture         =   "frmPRC.frx":272F
         Top             =   60
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmPRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const lCaptionHeight = 600
Dim bMouseDown As Boolean
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim lPymtEntryID As Long
Dim strExchDiffAccNo As String
Dim dAmtForED As Double
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
    lBorderColor = RGB(150, 195, 95)
    lFillColor = RGB(150, 195, 95)
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
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 20, Pic.Height), lBorderColor  'Right

End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    On Error GoTo err
    
    If cmbCountry.MatchFound = False Then Exit Sub
           
    If cmbCountry.List(cmbCountry.ListIndex, 1) = "EUR" Then
        ImgEuro.Visible = True
        ImgDollar.Visible = False
    Else
        ImgEuro.Visible = False
        ImgDollar.Visible = True
    End If
   
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "Select Country,Curr,AccNo From ForeignCustomers "
    
    If cmbCust.ID <> "0" Then
        SQL = SQL & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(.AbsolutePosition - 1, 1) = !Curr & ""
            cmbCountry.List(.AbsolutePosition - 1, 2) = !AccNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim iBankID As Integer
    Dim i As Integer
    With rs
        .Open "SELECT TotalCustomAmt FROM VCustomInvoiceAmt WHERE CustomInvoice='" & cmbCustomInvoice.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        txtAmount.Text = Val(!TotalCustomAmt & "")
        .Close
        iBankID = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 1)
'        If iBankID = 7 Then
'            StrSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE AccNo='15-002-16001'"
'        ElseIf iBankID = 8 Then
'            StrSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE AccNo='15-002-21001'"
'        Else
'            StrSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE BankID=" & iBankID
'        End If
        strSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE BankID=" & iBankID
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        cmbBanks.Clear
        Do Until .EOF
            cmbBanks.AddItem !AccTitle & ""
            cmbBanks.List(cmbBanks.ListCount - 1, 1) = !AccNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    If cmbBanks.ListCount > 0 Then cmbBanks.ListIndex = 0
    
    txtPostedAmt = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 2)
    cmbCust.ID = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 3)
    cmbCountry.Text = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 4)
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    If cmbCustomInvoice.MatchFound = False Then
        MsgBox "Please Select Invoice.", vbInformation
        Exit Sub
    ElseIf Val(txtAmtRealized) <= 0 Then
        MsgBox "Invalid Realized Amount.", vbInformation
        Exit Sub
    ElseIf Val(txtExchRate) <= 0 Then
        MsgBox "Invalid Exchange Rate.", vbInformation
        Exit Sub
    ElseIf Val(txtAmtInPakRs) <= 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Sub
    ElseIf cmbBanks.MatchFound = False Then
        MsgBox "Please Select Bank Account.", vbInformation
        Exit Sub
    ElseIf txtRunningSerialNo = "" Then
        MsgBox "Invalid Serial No.", vbInformation
        Exit Sub
    ElseIf txtBillNo = "" Then
        MsgBox "Invalid Bill No.", vbInformation
        Exit Sub
    End If
    
    Dim strVchrNo As String, lEntryID As Long, iBankID As Integer
    Call StartTrans(con)
    iBankID = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 1)
    
    con.Execute "INSERT INTO PRC(CustomInvoice,RunningSerialNo,SerialNoDT,BillNo,BillNoDT,RealizationDT,AmtRealized," & _
     "Commission,FBCharges,EDS,WithHoldingTax,ExchRate,UserName,MachineName) VALUES('" & _
     cmbCustomInvoice.Text & "','" & txtRunningSerialNo & "','" & DTSerialNo & "','" & txtBillNo & "','" & _
     DTBillNo & "','" & DTRealization & "'," & Val(txtAmtRealized) & "," & Val(LV.ListItems(1).SubItems(2)) & _
     "," & Val(LV.ListItems(2).SubItems(2)) & "," & Val(LV.ListItems(3).SubItems(2)) & "," & _
     Val(LV.ListItems(4).SubItems(2)) & "," & Val(txtExchRate) & _
     ",'" & CurrentUserName & "','" & strComputerName & "')"
     
    lEntryID = GetSingleLongValue("MAX(EntryID)", "PRC")
'    If iBankID = 7 Then
'        strVchrNo = AddVoucherRBS(lEntryID)
'    Else
'        strVchrNo = AddVoucher
'    End If
    strVchrNo = AddVoucher
    con.Execute "UPDATE PRC SET VchrNo='" & strVchrNo & "' WHERE EntryID=" & lEntryID
    con.Execute "UPDATE FCustPayments SET PRCVchrNo='" & strVchrNo & "' WHERE EntryID=" & lPymtEntryID
    con.CommitTrans
    
'    Call SendPRCEmail(lEntryID)
    
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

'Private Sub SendPRCEmail(lEntryID As Long)
'
'    On Error GoTo err
'    Dim rpt As CRAXDDRT.Report
'    Set rpt = rptApp.OpenReport(rptExportPath & "\PRC.rpt")
'    rpt.RecordSelectionFormula = "{PRC.EntryID}=" & lEntryID
'
'    With rpt.ExportOptions
'        .FormatType = crEFTHTML40
'        .DestinationType = crEDTDiskFile
'        .HTMLEnableSeparatedPages = False
'        .HTMLFileName = "\\Server\Software-Auto\PRC.html"
'        .HTMLHasPageNavigator = False
'    End With
'    On Error Resume Next
'    rpt.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'    On Error GoTo 0
'    rpt.Export False
'    DoEvents
'
'    Dim oFso As New FileSystemObject
'    Dim ts As TextStream
'    Dim strTemp As String
'    Dim str As String
'    If oFso.FileExists("\\Server\Software-Auto\PRC.html") Then
'        Set ts = oFso.OpenTextFile("\\Server\Software-Auto\PRC.html", 1, False)
'        str = ts.ReadAll
'        ts.Close
'        Set ts = Nothing
'    End If
'    Set oFso = Nothing
'
'    Call SendEmail("afzal@vision.com.pk,imtiaz@vision.com.pk,finance@vision.com.pk", "PRC Feeded", str, , "asim@vision.com.pk")
'    'Call SendEmail("asim@vision.com.pk", "PRC Feeded", str)
'
'    Exit Sub
'err:
'    MsgBox err.Description
'End Sub

Private Function AddVoucherRBS(lEntryID As Long) As String

    'In case of RBS... there will be a voucher for every Expense in order to reflect that on Bank Statement...
    AddVoucherRBS = ""
    Dim strVchrNo As String
    Dim VNo As String, m_DTVchr As Date
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double, Desc As String, dAmt As Double, dDeductionAmt As Double
    Dim strBankAccNo As String, strCustomerAccNo As String
    
    m_DTVchr = DTPosting
    strBankAccNo = cmbBanks.List(cmbBanks.ListIndex, 1)
    strCustomerAccNo = cmbCountry.List(cmbCountry.ListIndex, 2)
    
    Desc = cmbCustomInvoice.Text & " " & Val(txtAmtRealized) & " " & cmbCountry.List(cmbCountry.ListIndex, 1) & " @ " & Val(txtExchRate) & " " & txtBillNo & " " & cmbCust.ID & "(" & cmbCountry.Text & ")"
    
    Dim objVchr As New clsVouchers
    objVchr.VoucherType = "PRC"
    objVchr.VoucherDate = m_DTVchr
    'Debit Bank Account
    objVchr.AddVoucherHead strBankAccNo, Val(txtAmtInPakRs), Desc
    If Right$(txtForeignExchDiff, 2) = "Dr" Then
        objVchr.AddVoucherHead strExchDiffAccNo, Val(txtForeignExchDiff.Tag), Desc
    End If
    'Credit Customer Account, Credit with Inovice's Posted Amount, and Adjust Diff in Exch. Diff Account....
    objVchr.AddVoucherHead strCustomerAccNo, -Val(txtPostedAmt), Desc
    If Right$(txtForeignExchDiff, 2) = "Cr" Then
        objVchr.AddVoucherHead strExchDiffAccNo, -Val(txtForeignExchDiff.Tag), Desc
    End If
    strVchrNo = objVchr.PostVoucher(False, True)
    
    con.Execute "INSERT INTO PRCVouchers(RefID,VchrNo) VALUES(" & lEntryID & ",'" & strVchrNo & "')"
    AddVoucherRBS = strVchrNo
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
        If .SubItems(3) <> "PKR" Then
            dDeductionAmt = (Val(.SubItems(2)) * Val(txtExchRate))
        Else
            dDeductionAmt = (Val(.SubItems(2)))
        End If
        If dDeductionAmt > 0 Then
            Set objVchr = New clsVouchers
            objVchr.VoucherType = "PRC"
            objVchr.VoucherDate = m_DTVchr
            'Debit
            objVchr.AddVoucherHead .key, dDeductionAmt, Desc
            'Credit
            objVchr.AddVoucherHead strBankAccNo, -dDeductionAmt, Desc
            strVchrNo = objVchr.PostVoucher(False, True)
            con.Execute "INSERT INTO PRCVouchers(RefID,VchrNo) VALUES(" & lEntryID & ",'" & strVchrNo & "')"
        End If
        End With
    Next
    
    Exit Function
err:
    AddVoucherRBS = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
End Function

Private Function AddVoucher() As String
    
    'Post Net Amount... In Case of Banks other Than RBS...
    AddVoucher = ""
    
    Dim VNo As String, m_DTVchr As Date
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double, Desc As String, dAmt As Double
    
    m_DTVchr = DTPosting
'    VNo = getNextVchrNo(m_DTVchr, "JV")
'    INextSno = getNextSno(m_DTVchr)
    
    Dim i As Integer
    
    Dim dNetAmt As Double, dDeductionAmt As Double
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SubItems(3) <> "PKR" Then
                dDeductionAmt = dDeductionAmt + (Val(.SubItems(2)) * Val(txtExchRate))
            Else
                dDeductionAmt = dDeductionAmt + (Val(.SubItems(2)))
            End If
        End With
    Next
    dNetAmt = Round(Val(txtAmtInPakRs), 2) - Round(dDeductionAmt, 2)
    
    Dim strBankAccNo As String
    strBankAccNo = cmbBanks.List(cmbBanks.ListIndex, 1)
    
    'Bank Account Debit Entry
    
    'Remember Sno For Voucher Details
    SnoToUse = INextSno
'    dNetAmt = Val(txtAmtInPakRs)
    
'    Bal = getBalance(strBankAccNo, m_DTVchr)
'    Bal = Bal + dNetAmt
    
    Desc = cmbCustomInvoice.Text & " " & Val(txtAmtRealized) & " " & cmbCountry.List(cmbCountry.ListIndex, 1) & " @ " & Val(txtExchRate) & " " & txtBillNo & " " & cmbCust.ID & "(" & cmbCountry.Text & ")"
     
    Dim objVchr As New clsVouchers
    objVchr.VoucherType = "JV"
    objVchr.VoucherDate = m_DTVchr
     
    objVchr.AddVoucherHead strBankAccNo, dNetAmt, Desc
          
    Dim strAccNo As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SubItems(3) <> "PKR" Then
                dDeductionAmt = (Val(.SubItems(2)) * Val(txtExchRate))
            Else
                dDeductionAmt = (Val(.SubItems(2)))
            End If
            dDeductionAmt = Round(dDeductionAmt, 2)
            If dDeductionAmt > 0 Then
                'INextSno = INextSno + 1
                strAccNo = .key
'                Bal = getBalance(strAccNo, m_DTVchr)
'                Bal = Bal + dDeductionAmt
                
                'con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                     Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strAccNo & _
                    "','" & Desc & "'," & dDeductionAmt & ",0," & Bal & ",'',0)"
                      
                'Call adjustLedger(strAccNo, m_DTVchr, INextSno, dNetAmt)
                objVchr.AddVoucherHead strAccNo, dDeductionAmt, Desc
            End If
        End With
    Next
    
    'Foreign Exchange Diff. If Debit .. .Post Now...
    If Right$(txtForeignExchDiff, 2) = "Dr" Then
        strAccNo = strExchDiffAccNo
        dAmt = Val(txtForeignExchDiff.Tag)
        INextSno = INextSno + 1
'        Bal = getBalance(strAccNo, m_DTVchr)
'        Bal = Bal + dAmt
'
'        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
'            Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strAccNo & _
'            "','" & Desc & "'," & dAmt & ",0," & Bal & ",'',0)"
'
'        Call adjustLedger(strAccNo, m_DTVchr, INextSno, dAmt)
        objVchr.AddVoucherHead strAccNo, dAmt, Desc
    End If
    
    'It Finishes Debit Entries...
    'Starting Credit Entries Now
    
'    INextSno = INextSno + 1
    'Foreign Customer Credit Entry...
    strAccNo = cmbCountry.List(cmbCountry.ListIndex, 2)
    'dAmt = Val(txtPostedAmt)
    dAmt = dAmtForED
'    Bal = getBalance(strAccNo, m_DTVchr)
'    Bal = Bal - dAmt
'
'    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
'        Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strAccNo & _
'        "','" & Desc & "',0," & dAmt & "," & Bal & ",'',0)"
'
'    Call adjustLedger(strAccNo, m_DTVchr, INextSno, -dAmt)
    objVchr.AddVoucherHead strAccNo, -dAmt, Desc
    
    If Right$(txtForeignExchDiff, 2) = "Cr" Then
        strAccNo = strExchDiffAccNo '= "32-06001"
        dAmt = Val(txtForeignExchDiff.Tag)
'        INextSno = INextSno + 1
'        Bal = getBalance(strAccNo, m_DTVchr)
'        Bal = Bal - dAmt
'
'        con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
'            Format(m_DTVchr, "dd-MMM-yyyy") & "','" & VNo & "','" & strAccNo & _
'            "','" & Desc & "',0," & dAmt & "," & Bal & ",'',0)"
'
'        Call adjustLedger(strAccNo, m_DTVchr, INextSno, -dAmt)
        objVchr.AddVoucherHead strAccNo, -dAmt, Desc
    End If
    
    'con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    VNo = objVchr.PostVoucher(False, True, , True)
    Set objVchr = Nothing
    
    AddVoucher = VNo
    
    Exit Function
err:
    AddVoucher = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
    
End Function

Private Sub DTRealization_Change()
    DTPosting.Value = DTRealization.Value
End Sub

Private Sub Form_Load()
    cmbCust.ListHeight = 2600
    DTSerialNo = Date
    DTBillNo = Date
    DTRealization = Date
    DTPosting = Date
    Call DrawPicCaptions
    bMouseDown = False
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LV.SelectedItem.Tag = "Debit" Then
        iColNo = 3
'    ElseIf LV.SelectedItem.Tag = "Credit" Then
'        iColNo = 4
'    Else
        iColNo = iColNoTemp
    End If
    
    If iColNo <> 3 Then Exit Sub
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > LV.ColumnHeaders(3).Left And X < LV.ColumnHeaders(3).Left + LV.ColumnHeaders(3).Width Then
        iColNoTemp = 3
    ElseIf X > LV.ColumnHeaders(4).Left And X < LV.ColumnHeaders(4).Left + LV.ColumnHeaders(4).Width Then
        iColNoTemp = 4
    End If
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = X
                lDownY = Y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((X - lDownX) / 15), Me.Top + ((Y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = False
End Sub

Public Sub ShowMe(m_lPymtEntryID As Long, strInvoiceNo As String, dValue As Double, strBillNo As String, dExchRate As Double)
    
    lPymtEntryID = m_lPymtEntryID
    cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode", " Order By ForeignCustomers.CustCode"
    cmbCust.Enabled = False
    cmbCountry.Enabled = False
    txtBillNo = strBillNo
    txtAmtRealized = dValue
    
    Dim rs As New ADODB.Recordset
    
    With rs
        '.Open "SELECT CustCode,Country,OurBankID,CustomInvoice,AmtInPakRs FROM VCustomInvoiceAuthorized WHERE " & _
        " CustomInvoice NOT IN(SELECT CustomInvoice FROM PRC) Order By CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT CustCode,Country,OurBankID,CustomInvoice,AmtInPakRs,ExchRate FROM VCustomInvoiceAuthorized ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbCustomInvoice.Clear
        Do Until .EOF
            cmbCustomInvoice.AddItem !CustomInvoice & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 1) = !OurBankID & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 2) = Val(!AmtInPakRs & "")
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 3) = !CustCode & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 4) = !Country & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 5) = Val(!ExchRate & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    cmbCustomInvoice.Text = strInvoiceNo
    txtExchRate.Text = dExchRate
     
     
    Dim ITM As ListItem
     
    Dim strAccNo As String
    strAccNo = getGeneralDataValue("PRC_FBCharges_AccNo")   '44-06001
    
    If strAccNo = "" Then
        MsgBox "Bank Service Charges Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "Bank Charges (F B Charges)"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(3) = cmbCountry.List(cmbCountry.ListIndex, 1) & ""
   'LV.ListItems(1).SubItems(3) = cmbCountry.List(cmbCountry.ListIndex, 1)
'    Set ITM = LV.ListItems.add(, "44-002-03001", LV.ListItems.count + 1)
'    ITM.Tag = "Debit"
'    ITM.SubItems(1) = "Commission"
'    ITM.SubItems(2) = "0"
'    ITM.SubItems(3) = "0"
'    ITM.SubItems(3) = "PKR"
    
    strAccNo = getGeneralDataValue("PRC_EDS_AccNo")   '44-001
    
    If strAccNo = "" Then
        MsgBox "EDS Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "EDS"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(3) = "PKR"
     
    strAccNo = getGeneralDataValue("PRC_BSC_AccNo")   '"44-02001"
    
    If strAccNo = "" Then
        MsgBox "Bank Service Charges Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "Bank Service Charges"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(3) = "PKR"
    
'    Set ITM = LV.ListItems.add(, "45-05002", LV.ListItems.count + 1)
'    ITM.Tag = "Debit"
'    ITM.SubItems(1) = "Excise Duty"
'    ITM.SubItems(2) = "0"
'    ITM.SubItems(3) = "0"
'    ITM.SubItems(3) = "PKR"
    
    strAccNo = getGeneralDataValue("PRC_WithHolding_AccNo")   '45-23004
    
    If strAccNo = "" Then
        MsgBox "Withholding Tax Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "WithHolding Tax"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(3) = "PKR"
    
    strAccNo = getGeneralDataValue("PRC_ExchDiff_AccNo")   '45-23004
    
    If strAccNo = "" Then
        MsgBox "Exch. Diff. Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    strExchDiffAccNo = strAccNo
    
'    Set ITM = LV.ListItems.add(, "44-002-02002", LV.ListItems.count + 1)
'    ITM.Tag = "Debit"
'    ITM.SubItems(1) = "Any Other Charges"
'    ITM.SubItems(2) = "0"
'    ITM.SubItems(3) = "0"
'    ITM.SubItems(3) = "PKR"
    
'    Set ITM = LV.ListItems.add(, "32-06001", LV.ListItems.Count + 1)
'    ITM.Tag = ""
'    ITM.SubItems(1) = "Foreign Exchange Difference"
'    ITM.SubItems(2) = "0"
'    ITM.SubItems(3) = "0"
'    ITM.SubItems(3) = "PKR"
    
    Me.Show 1
    
End Sub

Private Sub txtAmtRealized_Change()
    CalcualteAmounts
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtEdit.Text) < 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
        If LV.SelectedItem.Tag = "" Then
            LV.SelectedItem.SubItems(2) = 0
        End If
        LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEdit)
        Call CalculateNetAmount
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtExchRate_Change()
    CalcualteAmounts
End Sub

Private Sub CalcualteAmounts()

    txtAmtInPakRs = Round(Val(txtExchRate) * Val(txtAmtRealized), 4)
    
    Dim dExchDiff As Double, dExchRate_Auth As Double
    dExchRate_Auth = Val(cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 5) & "")
    
    dAmtForED = Round(Val(dExchRate_Auth) * Val(txtAmtRealized), 4)
    
    dExchDiff = Round(Val(txtAmtInPakRs) - Val(dAmtForED), 4)
    'dExchDiff = Round(Val(txtAmtInPakRs) - Val(txtPostedAmt), 4)
    
    If dExchDiff > 0 Then   'Credit
        txtForeignExchDiff = dExchDiff & " Cr"
        txtForeignExchDiff.Tag = dExchDiff
    Else
        txtForeignExchDiff = Abs(dExchDiff) & " Dr"
        txtForeignExchDiff.Tag = Abs(dExchDiff)
    End If
    Call CalculateNetAmount
    
End Sub

Private Sub CalculateNetAmount()

    Dim dNetAmount As Double, i As Integer
    dNetAmount = Val(txtAmtInPakRs)
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).SubItems(3) <> "PKR" Then
            dNetAmount = dNetAmount - (Val(LV.ListItems(i).SubItems(2)) * Val(txtExchRate))
        Else
            dNetAmount = dNetAmount - Val(LV.ListItems(i).SubItems(2))
        End If
    Next
    txtNetAmt = dNetAmount
    
End Sub
