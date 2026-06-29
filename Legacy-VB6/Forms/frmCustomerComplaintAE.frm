VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCustomerComplaintAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Complaints"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10560
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
   ScaleHeight     =   9810
   ScaleWidth      =   10560
   StartUpPosition =   1  'CenterOwner
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
      Height          =   315
      Left            =   8865
      TabIndex        =   56
      Top             =   3360
      Width           =   1560
   End
   Begin ComboList.Usercontrol1 cmbCompany 
      Height          =   285
      Left            =   315
      TabIndex        =   50
      Tag             =   "Company_RefID"
      Top             =   315
      Width           =   10110
      _ExtentX        =   17833
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
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   315
      TabIndex        =   0
      Tag             =   "CustCode"
      Top             =   975
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
   Begin ComboList.Usercontrol1 cmbItems 
      Height          =   285
      Left            =   2400
      TabIndex        =   17
      Top             =   2745
      Width           =   3660
      _ExtentX        =   6456
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
   Begin ComboList.Usercontrol1 cmbInvoiceNo 
      Height          =   285
      Left            =   315
      TabIndex        =   16
      Top             =   2745
      Width           =   2070
      _ExtentX        =   3651
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
      Height          =   3015
      Left            =   315
      TabIndex        =   30
      Top             =   3720
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   5318
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
         Text            =   "Invoice No."
         Object.Width           =   3572
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product"
         Object.Width           =   6456
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Reason"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ref."
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Item Pic"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Reason"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1170
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Tag             =   "ActionRequested"
      Top             =   4470
      Visible         =   0   'False
      Width           =   10110
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      Height          =   1140
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Tag             =   "Reason"
      Top             =   4635
      Visible         =   0   'False
      Width           =   10110
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   1035
      Left            =   285
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Tag             =   "Description"
      Top             =   4245
      Visible         =   0   'False
      Width           =   10110
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   780
      Top             =   9165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
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
      Left            =   315
      TabIndex        =   37
      Top             =   8730
      Width           =   9690
   End
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
      Left            =   10020
      TabIndex        =   36
      Top             =   8715
      Width           =   405
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   7080
      TabIndex        =   3
      Tag             =   "DT"
      Top             =   975
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   91029507
      CurrentDate     =   40698
   End
   Begin MSComDlg.CommonDialog CD2 
      Left            =   330
      Top             =   9165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSForms.ComboBox cmbLots 
      Height          =   315
      Left            =   8865
      TabIndex        =   52
      Top             =   2730
      Width           =   1560
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2752;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Picture 2"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   12
      Left            =   1920
      TabIndex        =   55
      Top             =   6765
      Width           =   630
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1380
      Left            =   1890
      Stretch         =   -1  'True
      Top             =   7005
      Width           =   1545
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Picture 1"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   11
      Left            =   345
      TabIndex        =   54
      Top             =   6765
      Width           =   630
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1380
      Left            =   315
      Stretch         =   -1  'True
      Top             =   6990
      Width           =   1545
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   315
      Left            =   2655
      TabIndex        =   1
      Tag             =   "Country"
      Top             =   975
      Width           =   1845
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3254;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbCc 
      Height          =   315
      Left            =   5385
      TabIndex        =   15
      Tag             =   "EmailCC"
      Top             =   2160
      Width           =   5040
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8890;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbTo 
      Height          =   315
      Left            =   315
      TabIndex        =   14
      Tag             =   "EmailTo"
      Top             =   2160
      Width           =   5040
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8890;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Reason"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   10
      Left            =   7260
      TabIndex        =   53
      Top             =   2520
      Width           =   540
   End
   Begin MSForms.ComboBox cmbReason 
      Height          =   315
      Left            =   7230
      TabIndex        =   51
      Top             =   2730
      Width           =   1620
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2857;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Company"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   9
      Left            =   330
      TabIndex        =   49
      Top             =   105
      Width           =   675
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   8
      Left            =   8745
      TabIndex        =   48
      Top             =   765
      Width           =   285
   End
   Begin MSForms.TextBox txtItemReason 
      Height          =   315
      Left            =   315
      TabIndex        =   19
      Top             =   3360
      Width           =   8505
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "15002;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Reason"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   6
      Left            =   315
      TabIndex        =   47
      Top             =   3150
      Width           =   540
   End
   Begin MSForms.ComboBox cmbUnit 
      Height          =   315
      Left            =   8745
      TabIndex        =   4
      Tag             =   "Unit"
      Top             =   975
      Width           =   1650
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2910;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cc"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   9
      Left            =   5385
      TabIndex        =   46
      Top             =   1935
      Width           =   180
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   8
      Left            =   315
      TabIndex        =   45
      Top             =   1935
      Width           =   180
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   7
      Left            =   8745
      TabIndex        =   44
      Top             =   1335
      Width           =   660
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   6
      Left            =   7500
      TabIndex        =   43
      Top             =   1335
      Width           =   555
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   6270
      TabIndex        =   42
      Top             =   1335
      Width           =   270
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   6
      Left            =   8745
      TabIndex        =   10
      Tag             =   "PenaltyCurrency"
      Top             =   1560
      Width           =   1650
      VariousPropertyBits=   679495705
      BorderStyle     =   1
      Size            =   "2910;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   5
      Left            =   7500
      TabIndex        =   9
      Tag             =   "PenaltyAmount"
      Top             =   1560
      Width           =   1230
      VariousPropertyBits=   679495705
      BorderStyle     =   1
      Size            =   "2170;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   4
      Left            =   6255
      TabIndex        =   8
      Tag             =   "PenaltyQty"
      Top             =   1560
      Width           =   1230
      VariousPropertyBits=   679495705
      BorderStyle     =   1
      Size            =   "2170;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   4515
      TabIndex        =   41
      Top             =   1335
      Width           =   540
   End
   Begin MSForms.ComboBox cmbPenalty 
      Height          =   315
      Left            =   4530
      TabIndex        =   7
      Tag             =   "Penalty"
      Top             =   1560
      Width           =   1695
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "2990;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   3
      Left            =   2655
      TabIndex        =   6
      Tag             =   "Designation"
      Top             =   1560
      Width           =   1845
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "3254;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   2670
      TabIndex        =   40
      Top             =   1335
      Width           =   840
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   2
      Left            =   315
      TabIndex        =   5
      Tag             =   "ComplaintBy"
      Top             =   1560
      Width           =   2310
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4075;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complaint By"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   2
      Left            =   345
      TabIndex        =   39
      Top             =   1335
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PDF Path to Attach"
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
      Index           =   1
      Left            =   330
      TabIndex        =   38
      Top             =   8445
      Width           =   2040
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   825
      TabIndex        =   35
      Top             =   4665
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Reference"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   8895
      TabIndex        =   34
      Top             =   2520
      Width           =   750
   End
   Begin MSForms.TextBox txtEmailRef 
      Height          =   315
      Left            =   6630
      TabIndex        =   20
      Top             =   7350
      Visible         =   0   'False
      Width           =   1560
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2752;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Qty"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   2
      Left            =   6060
      TabIndex        =   33
      Top             =   2520
      Width           =   270
   End
   Begin MSForms.TextBox txtQty 
      Height          =   315
      Left            =   6075
      TabIndex        =   18
      Top             =   2730
      Width           =   1140
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2011;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   2385
      TabIndex        =   32
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Invoice No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   345
      TabIndex        =   31
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Action Requested"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Country"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   7
      Left            =   2670
      TabIndex        =   28
      Top             =   765
      Width           =   585
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   345
      TabIndex        =   27
      Top             =   765
      Width           =   690
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Complaint"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   19
      Left            =   345
      TabIndex        =   26
      Top             =   4425
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Complaint No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   73
      Left            =   4515
      TabIndex        =   25
      Top             =   765
      Width           =   1005
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   14
      Left            =   4530
      TabIndex        =   2
      Tag             =   "ComplaintNo"
      Top             =   975
      Width           =   2520
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Complaint Date"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   7125
      TabIndex        =   24
      Top             =   765
      Width           =   1095
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   8865
      TabIndex        =   23
      Top             =   9255
      Width           =   1575
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2778;741"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   420
      Left            =   7230
      TabIndex        =   22
      Top             =   9255
      Width           =   1575
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2778;741"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   420
      Left            =   5595
      TabIndex        =   21
      Top             =   9255
      Visible         =   0   'False
      Width           =   1575
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2778;741"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmCustomerComplaintAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim strDelList As String

Public Sub EditItem(p_lEntryID As Long)

    On Error GoTo err
    lEntryID = p_lEntryID
    
    Call ShowData(lEntryID)
    
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillCmbs()

    On Error GoTo 0
    
    cmbCompany.ListHeight = 6000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    cmbCompany.ID = 1
    
    cmbCust.ListHeight = 2500
    Call cmbCust.AddVals(con, "DISTINCT CustCode", "ForeignCustomers", "CustCode")
    
'    Call AddToCombo(cmbTo, "ComplaintFormEMailList", "ComplaintFormEMailList", , , "ComplaintFormEMailList")
'    Call AddToCombo(cmbCc, "ComplaintFormEMailList", "ComplaintFormEMailList", , , "ComplaintFormEMailList")
    
    cmbInvoiceNo.ListHeight = 2500
    cmbItems.ListHeight = 2500
    
    With cmbPenalty
        .AddItem "None"
        .AddItem "Free Placement"
        .AddItem "Credit Note"
        .AddItem "Investigation"
    End With
    
'    cmbUnit.AddItem "Main Unit"
'    cmbUnit.List(0, 1) = "Main Unit"
    Call AddToCombo(cmbUnit, "Departments.Name", "Departments", "", False, "Departments.Name,Departments.DeptID")
    
    Call AddToCombo(cmbReason, "Reason_QMS", "QMS_Reasons_List", " ORDER BY EntryID", False)
'    cmbUnit.AddItem "Unit 2"
'    cmbUnit.List(1, 1) = "Unit 2"
    
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT DISTINCT Country FROM ForeignCustomers "
    
    If cmbCust.ID <> "0" Then
        SQL = SQL & " WHERE CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
    
    
    cmbInvoiceNo.ClearVals
    cmbInvoiceNo.AddItem "<All Items>", "0"
    cmbInvoiceNo.AddVals con, "CustomInvoice", "CustomInvoice", "CustomInvoice", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
    cmbInvoiceNo.ID = "0"
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbInvoiceNo_matched()
    
    cmbItems.DropDownOffLine = True
    cmbItems.ClearVals
     
    If cmbInvoiceNo.ID = "0" Then
        If cmbCust.MatchFound = False Then
            cmbItems.AddVals con, "ItemID", "Items", "ItemID"
        Else
            cmbItems.AddVals con, "'{' + CompItemID + '} ' + ItemID", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
        End If
    Else
        cmbItems.AddVals con, "ItemID + ' {'+CustItemID+'}'", "VCustomInvoiceItems", "ItemID", " WHERE CustomInvoice='" & cmbInvoiceNo.Text & "'"
    End If
    
End Sub

Private Sub cmbItems_matched()

    'Fill Lots in dropdown
    Call AddToCombo(cmbLots, "DISTINCT LotNo", "VCustomPList", " WHERE CustomInvoice='" & cmbInvoiceNo.ID & "' AND ItemID='" & cmbItems.ID & "'")
    
End Sub

Private Sub cmbPenalty_Change()
    Call cmbPenalty_Click
End Sub

Private Sub cmbPenalty_Click()

    Tbox(4).Enabled = False
    Tbox(5).Enabled = False
    Tbox(6).Enabled = False
    If cmbPenalty.ListIndex = 1 Then
        Tbox(4).Enabled = True
    ElseIf cmbPenalty.ListIndex = 2 Then
        Tbox(5).Enabled = True
        Tbox(6).Enabled = True
    End If
    
End Sub

Private Sub cmbUnit_Change()
    Call cmbUnit_Click
End Sub

Private Sub cmbUnit_Click()

    
    If cmbUnit.MatchFound = False Then Exit Sub
    
    Dim strDeptID As String
    strDeptID = cmbUnit.List(cmbUnit.ListIndex, 2)
    
    Call AddToCombo(cmbTo, "Employees.Name+' {'+Employees.EmpID+'}'", "Employees", " WHERE DeptID='" & strDeptID & "' AND Active=1", False, "Employees.EmpID")
    Call AddToCombo(cmbCc, "Employees.Name+' {'+Employees.EmpID+'}'", "Employees", " WHERE DeptID='" & strDeptID & "' AND Active=1", False, "Employees.EmpID")
    
End Sub

Private Sub cmdAdd_Click()

        If cmbReason.MatchFound = False Then
            MsgBox "Please Select Reason.", vbInformation
            Exit Sub
        End If
        
        Dim ITM As ListItem
        Set ITM = LV.ListItems.add(, , cmbInvoiceNo.Text)
        ITM.ListSubItems.add(, , cmbItems.Text).Tag = cmbItems.ID
        ITM.ListSubItems.add , , Val(txtQty)
        ITM.ListSubItems.add , , cmbReason.Text
        ITM.ListSubItems.add , , cmbLots.Text
        ITM.ListSubItems.add , , ""
        ITM.ListSubItems.add , , txtItemReason
        
End Sub

Private Sub cmdBrowseBarcode_Click()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        txtFilePath = Picfile
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdNext_Click()
    If SaveItem Then
        Unload Me
        
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub

Private Sub DTPicker2_Change()
    Tbox(14) = GetNewComplaintNo(DTPicker2.Value)
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker2.Value = Date
    strDelList = ""
    Call DTPicker2_Change
    
    Call FillCmbs
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    If cmbUnit.MatchFound = False Then
        MsgBox "Please Select Unit."
        Exit Function
    End If
'    If cmbCc.MatchFound = False Then
'        MsgBox "Please Select Cc."
'        Exit Function
'    End If
'    If cmbTo.MatchFound = False Then
'        MsgBox "Please Select To."
'        Exit Function
'    End If
    Dim SQL As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        SQL = CreateInsertSQL("CustomerComplaints")
    Else
        SQL = CreateUpdateSQL("CustomerComplaints")
        SQL = SQL & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute SQL
    
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "CustomerComplaints")
    Else
        lmyEntryID = lEntryID
    End If
    If lEntryID = 0 Then
        con.Execute "UPDATE CustomerComplaints SET UserName='" & CurrentUserName & "',MachineName='" & strComputerName & "' WHERE EntryID=" & lmyEntryID
    End If
    
    con.Execute "UPDATE CustomerComplaints SET Type=0 WHERE EntryID=" & lmyEntryID
    If strDelList <> "" Then
        strDelList = Left(strDelList, Len(strDelList) - 1)
        con.Execute "DELETE FROM CustomerComplaints_Detail WHERE EntryID IN(" & strDelList & ")"
    End If
    'con.Execute "DELETE FROM CustomerComplaints_Detail WHERE RefID=" & lmyEntryID
    Dim i As Integer
    Dim lCC_D_EntryID As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
        
            If Val(.Tag) = 0 Then
                con.Execute "INSERT INTO CustomerComplaints_Detail(RefID,InvoiceNo,ItemID,Qty,ItemReason_Main,EmailReference,ItemReason) VALUES(" & lmyEntryID & _
                 ",'" & .Text & "','" & .ListSubItems(1).Tag & "'," & Val(.SubItems(2)) & ",'" & .SubItems(3) & "','" & .SubItems(4) & "','" & .SubItems(6) & "')"
                 
                lCC_D_EntryID = GetSingleLongValue("MAX(EntryID)", "CustomerComplaints_Detail")
            Else
                lCC_D_EntryID = Val(.Tag)
            End If
            
            If .SubItems(5) = "Add" Then
                Call AttachPic(lCC_D_EntryID, .ListSubItems(5).Tag)
            ElseIf .SubItems(5) = "Remove" Then
                con.Execute "UPDATE CustomerComplaints_Detail SET ItemPic=NULL WHERE EntryID=" & lCC_D_EntryID
            End If
            
        End With
    Next
    con.CommitTrans
     
    
    Call AttachPDF(lmyEntryID)
   
    
SavePic:
    
    Dim Rec As New ADODB.Recordset
    Dim rsPic As New ADODB.Stream
    If Not (Image1.Picture = 0 Or Trim(Image1.Tag) = "") Then
        Rec.Open "SELECT EntryID,CC_Pic_1 FROM CustomerComplaints WHERE EntryID=" & lmyEntryID, con, adOpenStatic, adLockOptimistic
        If Image1.Picture <> 0 And Trim(Image1.Tag) <> "" Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile Image1.Tag
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
            Rec.Close
        End If
    End If
    If Not (Image1.Picture = 0 Or Trim(Image1.Tag) = "") Then
        Rec.Open "SELECT EntryID,CC_Pic_2 FROM CustomerComplaints WHERE EntryID=" & lmyEntryID, con, adOpenStatic, adLockOptimistic
        If Image2.Picture <> 0 And Trim(Image2.Tag) <> "" Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile Image2.Tag
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
            Rec.Close
        End If
    End If
    
    lEntryID = 0
    SaveItem = True
    
    
    Exit Function
err:
    MsgBox err.Description
    
End Function



Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "INSERT INTO " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Or (TypeOf c Is VB.TextBox) Then
            If c.Tag <> "" Then
                RetString = RetString & c.Tag & ","
            End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.ComboBox) Then
                If c.Name = "cmbCountry" Then
                    RetString = RetString & "'" & c.Text & "" & "',"
                ElseIf c.Name = "cmbPenalty" Then
                    RetString = RetString & c.ListIndex & ","
                ElseIf c.Name = "cmbCc" Or c.Name = "cmbTo" Then
                    If c.MatchFound = False Then
                        RetString = RetString & "'',"
                    Else
                        RetString = RetString & "'" & c.List(c.ListIndex, 1) & "',"
                    End If
                Else
                    RetString = RetString & "'" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Then
                If c.Name = "NBox" Then
                    RetString = RetString & Val(c.Text & "") & ","
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & c.ID & "',"
            ElseIf (TypeOf c Is DTPicker) Then
                RetString = RetString & "'" & c.Value & "',"
            End If
        End If
    Next
    
    '---------------------------------------------------------------------------
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ")"
    CreateInsertSQL = RetString
    
End Function
Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.ComboBox) Then
                If c.Name = "cmbCountry" Then
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                ElseIf c.Name = "cmbPenalty" Then
                    RetString = RetString & c.Tag & "=" & c.ListIndex & ","
                ElseIf c.Name = "cmbCc" Or c.Name = "cmbTo" Then
                    If c.MatchFound = False Then
                        RetString = RetString & c.Tag & "='',"
                    Else
                        RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                    End If
                Else
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Then   '
                If c.Name = "cmbCatID" Or c.Name = "cmbItemType" Then
                    RetString = RetString & c.Tag & "=" & Val(c.List(c.ListIndex, 1)) & ","
                ElseIf c.Name = "NBox" Then
                    RetString = RetString & c.Tag & "=" & Val(c.Text) & ","
                ElseIf c.Name = "cmbSteel" Or c.Name = "cmbItemGroup" Then
                    RetString = RetString & c.Tag & "=" & c.List(c.ListIndex, 1) & ","
                ElseIf c.Name = "cmbGuage" Then
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.ID & "',"
            ElseIf (TypeOf c Is DTPicker) Then
                RetString = RetString & c.Tag & "='" & c.Value & "',"
            End If
        End If
    Next
    
    
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function



Private Sub ShowData(p_lEntryID As Long, Optional bCopy As Boolean = False)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    Dim i As Integer
    With rs
        .Open "SELECT * FROM CustomerComplaints WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    c.ID = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is DTPicker Then
                    If Not IsNull(.Fields(c.Tag)) Then
                        c.Value = .Fields(c.Tag)
                    End If
                ElseIf c.Name = "cmbPenalty" Then
                    c.ListIndex = Val(.Fields(c.Tag) & "")
                Else
                    c.Text = .Fields(c.Tag) & ""
                End If
            End If
        Next
        
        If Not IsNull(![CC_Pic_1]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "CC_Pic_1"
        End If
        If Not IsNull(![CC_Pic_2]) Then
            Set Image2.DataSource = rs
            Image2.DataField = "CC_Pic_2"
        End If
        For i = 0 To cmbTo.ListCount - 1
            If cmbTo.List(i, 1) = !EmailTo & "" Then
                cmbTo.ListIndex = i
            End If
        Next
        For i = 0 To cmbCc.ListCount - 1
            If cmbCc.List(i, 1) = !EmailCC & "" Then
                cmbCc.ListIndex = i
            End If
        Next
        .Close
        
        Dim ITM As ListItem
        Dim strPicAttach As String
        .Open "SELECT * FROM VCustomerComplaints_Detail WHERE RefID=" & lEntryID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, , !InvoiceNo & "")
            ITM.Tag = !EntryID
            ITM.ListSubItems.add(, , !ITemID & " {" & !ItemCode & "}").Tag = !ITemID & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !ItemReason_Main & ""
            ITM.ListSubItems.add , , !EmailReference & ""
            
            If IsNull(!ItemPic) Then
                strPicAttach = ""
            Else
                strPicAttach = "Attached"
            End If
            ITM.ListSubItems.add , , strPicAttach
            ITM.ListSubItems.add , , !ItemReason & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub AddNew()
    
    Tbox(14) = GetNewComplaintNo(DTPicker2.Value)
    Me.Show 1
    
End Sub

Private Function GetNewComplaintNo(myDT As Date) As String
    'CC-1/2011
    Dim lComplaintNo As Long
    lComplaintNo = GetSingleLongValue("MAX(CAST(SUBSTRING(ComplaintNo,4,CHARINDEX('/',ComplaintNo)-4) AS INT))", "CustomerComplaints", " WHERE Type=0 AND LEFT(ComplaintNo,3)='CC-'")
    lComplaintNo = lComplaintNo + 1
    
    GetNewComplaintNo = "CC-" & lComplaintNo & "/" & year(myDT)
    
End Function

Private Sub Image1_DblClick()

    On Error Resume Next
    Dim Picfile As String
    CD2.ShowSave

    If Not (err.Number = 32755) Then
        Picfile = CD2.FileName
        Image1.Picture = LoadPicture(Picfile)
        Image1.Tag = Picfile
    End If
    
End Sub

Private Sub lblPic_DblClick()

    On Error Resume Next
    Dim Picfile As String
'    CD1.ShowSave
'
'    If Not (err.Number = 32755) Then
'        Picfile = CD1.FileName
'        Image1.Picture = LoadPicture(Picfile)
'        Image1.Tag = Picfile
'        lblPic.Visible = False
'    End If
    
End Sub

Private Sub Image2_DblClick()
    On Error Resume Next
    Dim Picfile As String
    CD2.ShowSave

    If Not (err.Number = 32755) Then
        Picfile = CD2.FileName
        Image2.Picture = LoadPicture(Picfile)
        Image2.Tag = Picfile
    End If
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LV.SelectedItem.Tag) > 0 Then
                strDelList = strDelList & Val(LV.SelectedItem.Tag) & ","
            End If
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    If LV.SelectedItem.SubItems(5) = "" Or LV.SelectedItem.SubItems(5) = "Remove" Then 'Not Attached
        strRet = ShowPopUpMenu1("Attach Pic")
    Else
        strRet = ShowPopUpMenu1("View Pic", "-", "Remove Pic")
    End If
    
    If strRet = "Attach Pic" Then
        
        On Error Resume Next
        Dim strPicfile As String
        CD2.ShowSave
    
        If Not (err.Number = 32755) Then
            strPicfile = CD2.FileName
            LV.SelectedItem.SubItems(5) = "Add"
            LV.SelectedItem.ListSubItems(5).Tag = strPicfile
        End If
        
    ElseIf strRet = "Remove Pic" Then
        LV.SelectedItem.SubItems(5) = "Remove"
        LV.SelectedItem.ListSubItems(5).Tag = ""
    ElseIf strRet = "View Pic" Then
        'View Pic...
        Dim f As New frmViewPictureFSExt
        Load f
        f.ShowFromDBField "ItemPic", "CustomerComplaints_Detail", " WHERE EntryID=" & Val(LV.SelectedItem.Tag)
    End If
     
End Sub

Private Sub txtEmailRef_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        If cmbReason.MatchFound = False Then
            MsgBox "Please Select Reason.", vbInformation
            Exit Sub
        End If
        Dim ITM As ListItem
        Set ITM = LV.ListItems.add(, , cmbInvoiceNo.Text)
        ITM.ListSubItems.add(, , cmbItems.Text).Tag = cmbItems.ID
        ITM.ListSubItems.add , , Val(txtQty)
        ITM.ListSubItems.add , , cmbReason.Text
        ITM.ListSubItems.add , , txtEmailRef
        ITM.ListSubItems.add , , ""
        ITM.ListSubItems.add , , txtItemReason
    End If
End Sub


Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,AttachmentPDF FROM CustomerComplaints WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub

Private Sub AttachPic(p_lEntryID As Long, strFilePath As String)
    
    On Error Resume Next
    If strFilePath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,ItemPic FROM CustomerComplaints_Detail WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strFilePath
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub

