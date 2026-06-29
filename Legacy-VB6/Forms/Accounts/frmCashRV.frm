VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCashRV 
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   12120
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
   ScaleHeight     =   8355
   ScaleWidth      =   12120
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   8355
      Left            =   90
      TabIndex        =   16
      Top             =   0
      Width           =   12000
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   4350
         TabIndex        =   30
         Top             =   3045
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   15
         TabIndex        =   43
         Top             =   -45
         Width           =   12000
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Reciept Voucher"
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
            Index           =   8
            Left            =   75
            TabIndex        =   44
            Top             =   165
            Width           =   12525
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Reciept Voucher"
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
            Index           =   9
            Left            =   90
            TabIndex        =   45
            Top             =   180
            Width           =   12510
         End
      End
      Begin MSComctlLib.ListView LVPymt 
         Height          =   3870
         Left            =   120
         TabIndex        =   42
         Top             =   2449
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   6826
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
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
            Text            =   "Voucher #"
            Object.Width           =   3755
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7302
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "Amt"
            Text            =   "Total Amout"
            Object.Width           =   2776
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount Due"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Applied Pymt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2593
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1605
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   945
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dddd, dd/MM/yyyy"
         Format          =   120979459
         CurrentDate     =   37250
      End
      Begin VB.TextBox Tbox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   8730
         TabIndex        =   6
         Top             =   1455
         Width           =   1335
      End
      Begin VB.TextBox Tbox 
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   4
         Top             =   1455
         Width           =   7080
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   5880
         Left            =   105
         TabIndex        =   19
         Top             =   2400
         Width           =   11895
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
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   90
            TabIndex        =   41
            Top             =   5235
            Value           =   1  'Checked
            Width           =   1830
         End
         Begin VB.CheckBox ChkClear 
            Caption         =   "Clear Cheque "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   40
            ToolTipText     =   "Clear This Cheque Automatically"
            Top             =   4425
            Width           =   1470
         End
         Begin VB.Frame FClearChq 
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   60
            TabIndex        =   32
            Top             =   4485
            Width           =   5085
            Begin VB.OptionButton Opts 
               Caption         =   "Receive in Cash In Hand"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   -15
               TabIndex        =   39
               Top             =   195
               Value           =   -1  'True
               Width           =   2370
            End
            Begin VB.OptionButton Opts 
               Caption         =   "Receive in Bank Account"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   2505
               TabIndex        =   38
               Top             =   195
               Width           =   2370
            End
            Begin VB.Frame Frame3 
               Height          =   660
               Left            =   15
               TabIndex        =   33
               Top             =   360
               Visible         =   0   'False
               Width           =   5010
               Begin VB.TextBox txtSlipNo 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   3945
                  TabIndex        =   35
                  Top             =   330
                  Width           =   1080
               End
               Begin VB.CommandButton cmdToBank 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   3645
                  Picture         =   "frmCashRV.frx":0000
                  Style           =   1  'Graphical
                  TabIndex        =   34
                  Top             =   330
                  Width           =   315
               End
               Begin VB.Label Label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00BECCD8&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " Bank Account                                                             Dpst Slip #"
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
                  Height          =   225
                  Index           =   2
                  Left            =   30
                  TabIndex        =   37
                  Top             =   120
                  Width           =   4980
               End
               Begin VB.Label lblBankAcc 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   30
                  TabIndex        =   36
                  Top             =   330
                  Width           =   3630
               End
            End
         End
         Begin VB.Frame Frame6 
            Height          =   660
            Left            =   6675
            TabIndex        =   31
            Top             =   4845
            Width           =   5175
            Begin MSForms.CommandButton cmdSClose 
               Height          =   360
               Left            =   1755
               TabIndex        =   14
               Top             =   195
               Width           =   1620
               Caption         =   "Save & Close  "
               PicturePosition =   327683
               Size            =   "2857;635"
               Accelerator     =   83
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.CommandButton cmdCancel 
               Height          =   360
               Left            =   3435
               TabIndex        =   15
               Top             =   195
               Width           =   1620
               Caption         =   "Cancel            "
               PicturePosition =   327683
               Size            =   "2857;635"
               Picture         =   "frmCashRV.frx":0372
               Accelerator     =   67
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.CommandButton cmdSClear 
               Height          =   360
               Left            =   75
               TabIndex        =   13
               Top             =   195
               Width           =   1620
               Caption         =   "Save & New   "
               PicturePosition =   327683
               Size            =   "2857;635"
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
         Begin VB.Label Lbls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash In Hand :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   0
            Left            =   3390
            TabIndex        =   46
            Top             =   4020
            Width           =   1155
         End
         Begin MSForms.ComboBox cmbCashInHand 
            Height          =   300
            Left            =   4605
            TabIndex        =   47
            Top             =   3975
            Width           =   3090
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5450;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CommandButton cmd 
            Height          =   360
            Left            =   45
            TabIndex        =   12
            Top             =   4035
            Width           =   2010
            VariousPropertyBits=   25
            Caption         =   "Auto Apply         "
            PicturePosition =   327683
            Size            =   "3545;635"
            Picture         =   "frmCashRV.frx":06C4
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073750017
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Lbls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount Applied :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   2
            Left            =   8295
            TabIndex        =   23
            Top             =   4125
            Width           =   1950
         End
         Begin VB.Label Lbls 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un Applied Amount :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   3
            Left            =   8490
            TabIndex        =   22
            Top             =   4470
            Width           =   1755
         End
         Begin VB.Label lblApplied 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   10350
            TabIndex        =   21
            Top             =   4065
            Width           =   1515
         End
         Begin VB.Label lblToApply 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   10350
            TabIndex        =   20
            Top             =   4410
            Width           =   1515
         End
      End
      Begin VB.Frame FraBank 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   105
         TabIndex        =   17
         Top             =   1770
         Width           =   11835
         Begin VB.TextBox Tbox 
            Height          =   315
            Index           =   2
            Left            =   6585
            TabIndex        =   9
            Top             =   240
            Width           =   1860
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   10350
            TabIndex        =   11
            Top             =   270
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   120979459
            CurrentDate     =   37250
         End
         Begin MSForms.ComboBox cmbCheckType 
            Height          =   300
            Left            =   8475
            TabIndex        =   10
            Top             =   240
            Width           =   1860
            VariousPropertyBits=   679495707
            DisplayStyle    =   7
            Size            =   "3281;529"
            ListWidth       =   2822
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbBank 
            Height          =   315
            Left            =   15
            TabIndex        =   7
            Top             =   240
            Width           =   2565
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "4524;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbBankBranch 
            Height          =   315
            Left            =   2610
            TabIndex        =   8
            Top             =   240
            Width           =   3945
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "6959;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label2 
            Height          =   255
            Index           =   1
            Left            =   30
            TabIndex        =   18
            Top             =   15
            Width           =   11640
            ForeColor       =   4194304
            BackColor       =   12504280
            VariousPropertyBits=   8388627
            Caption         =   $"frmCashRV.frx":0ABE
            Size            =   "20532;450"
            BorderColor     =   10006208
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.Label lblBalCap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash in Hand  "
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
         Height          =   195
         Left            =   10110
         TabIndex        =   25
         Top             =   1260
         Width           =   1185
      End
      Begin VB.Label lblBal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10095
         TabIndex        =   5
         Top             =   1455
         Width           =   1830
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   10560
         TabIndex        =   28
         Top             =   945
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin MSForms.ComboBox cmbPymtmode 
         Height          =   315
         Left            =   135
         TabIndex        =   3
         Top             =   1455
         Width           =   1470
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "2593;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   6375
         TabIndex        =   2
         Top             =   945
         Width           =   4170
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "7355;556"
         ListWidth       =   8819
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   3780
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   945
         Width           =   2595
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4577;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vocuher # "
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
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   24
         Top             =   750
         Width           =   945
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   26
         Top             =   735
         Width           =   11760
         ForeColor       =   4194304
         BackColor       =   12504280
         VariousPropertyBits=   8388627
         Caption         =   $"frmCashRV.frx":0B86
         Size            =   "20743;397"
         BorderColor     =   10860487
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   29
         Top             =   1260
         Width           =   11775
         ForeColor       =   4194304
         BackColor       =   12504280
         VariousPropertyBits=   8388627
         Caption         =   $"frmCashRV.frx":0C70
         Size            =   "20770;397"
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
End
Attribute VB_Name = "frmCashRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dontShowList As Boolean
Dim PermissionDenied As Boolean

Private Sub ChkClear_Click()
    FClearChq.Visible = ChkClear.Visible And ChkClear = vbChecked
End Sub

Private Sub CmbAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode = vbKeyDown Then
       If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
       KeyCode = 0
       dontShowList = True
    ElseIf KeyCode = vbKeyUp Then
       If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
       KeyCode = 0
       dontShowList = True
    ElseIf KeyCode = 13 And List1.Visible Then
       CmbAccTitle = List1
       List1.Visible = False
       CmbAccTitle.SetFocus
    End If

End Sub

Private Sub CmbAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If dontShowList Then dontShowList = False Else ShowList List1, CmbAccTitle, CmbAccTitle
End Sub

Private Sub cmbACCNo_Change()
    If cmbAccNo.MatchFound Then Call cmbACCNo_Click Else If Not cmbAccNo = "" Then cmbAccNo.DropDown
End Sub
Private Sub cmbACCNo_Click()

    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
    lblBalance.Visible = cmbAccNo.MatchFound
    Call RefreshList
    
End Sub
Private Sub cmbAccNo_Validate(cancel As Boolean)
    If Not cmbAccNo.MatchFound Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub
Private Sub cmbACCTitle_Change()
    If CmbAccTitle.MatchFound Then Call cmbACCTitle_Click 'Else If CmbAccTitle <> "" Then If CmbAccTitle <> "" Then CmbAccTitle.DropDown
End Sub
Private Sub cmbACCTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub

Private Sub CmbAccTitle_LostFocus()
    List1.Visible = False
End Sub

Private Sub CmbAccTitle_Validate(cancel As Boolean)
If Not CmbAccTitle.MatchFound Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
End Sub

Private Sub CmbAccTitle_DropButtonClick()
   List1.Visible = False
End Sub



Private Sub cmbBank_Change()
    If cmbBank.MatchFound Then Call cmbBank_Click
End Sub

Private Sub cmbBank_Click()
    Call AddToCombo(cmbBankBranch, "Distinct Branch + '(' + AccNo + ')'", "BankList", " Where Bank='" & cmbBank & "'")
End Sub

Private Sub cmbBankBranch_Change()
If cmbBankBranch.MatchFound Then Call cmbBankBranch_Click
End Sub

Private Sub cmbBankBranch_Click()
 Dim BankIDRec As New ADODB.Recordset
    With BankIDRec
        .CursorLocation = adUseClient
        .Open "Select BankID from BankList Where (Bank='" & cmbBank & "') and (Branch='" & cmbBankBranch & "')", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            cmbBankBranch.Tag = .Fields(0) & ""
        End If
        .Close
    End With
    Set BankIDRec = Nothing
    
End Sub



Private Sub cmbBankBranch_Validate(cancel As Boolean)
On Error GoTo err
If Not cmbBank.MatchFound And cmbBank <> "" Then
   MsgBox "This Is A New Bank This Will Be Added To Bank Profile." & vbNewLine & "You Can Add More Information About This Bank And Branch From Bank Profile", vbInformation
   cmbBank.AddItem cmbBank.Text
ElseIf Not cmbBankBranch.MatchFound And cmbBankBranch <> "" Then
   MsgBox "This Is A New Branch of " & cmbBank & " This Will Be Added To Bank Profile." & vbNewLine & "You Can Add More Information About This Branch From Bank Profile", vbInformation
Else
   Exit Sub
End If

Screen.MousePointer = vbHourglass
   con.Execute "Insert into bankList (Bank,Branch) values('" & cmbBank & "','" & cmbBankBranch & "')"
   'Sleep (800)
   Call cmbBank_Click
   Call cmbBankBranch_Click
Screen.MousePointer = vbDefault

Exit Sub

err:
If err.Number = -2147467259 Then
   MsgBox "Invalid Bank Or Branch Name", vbInformation
   cancel = True
Else
   MsgBox err.Description
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub cmbCashInHand_Change()
    Call cmbCashInHand_Click
End Sub

Private Sub cmbCashInHand_Click()
    lblBal = Format(getBalance(cmbCashInHand.List(cmbCashInHand.ListIndex, 1), DTPicker1), "0.00")
End Sub

Private Sub cmbPymtmode_Change()
    Call cmbPymtmode_Click
    Call DTPicker1_Change
End Sub

Private Sub cmbPymtmode_Click()

    If cmbPymtmode.ListIndex = 0 Then

        FraBank.Visible = False

        Label1(3).Visible = True
        Label1(4).Visible = True
        Tbox(0) = "Cash Received"

        chkPrint.Visible = True
        LVPymt.Move LVPymt.Left, 1890, LVPymt.Width, 4440
        cmbCashInHand.Enabled = True
    Else
        Tbox(0) = "Cheque Received"
        FraBank.Visible = True
        Label1(3).Visible = False

        cmbCheckType.ListIndex = 0
        DTPicker2 = DTPicker1
        chkPrint.Visible = False
        LVPymt.Move LVPymt.Left, 2440, LVPymt.Width, 3870
        cmbCashInHand.ListIndex = -1
        cmbCashInHand.Enabled = False
    End If


    Call DTPicker2_Change
    Call Form_Resize
    Call ShowBals
    
End Sub

Public Sub cmd_Click()


With LVPymt
    If Trim(cmd.Caption) = "Auto Apply" Then
        For i = 1 To .ListItems.count
            
            .ListItems(i).Checked = True
            Call LVPymt_ItemCheck(.ListItems(i))
            If Val(lblToApply) = 0 Then Exit For
        Next
        If i > 1 Then cmd.Caption = "Clear Payments     "
    Else
        For i = 1 To .ListItems.count
            .ListItems(i).Checked = False
            Call LVPymt_ItemCheck(.ListItems(i))
        Next
        lblToApply = Tbox(1)
        lblApplied = 0
        cmd.Caption = "Auto Apply         "
    End If
End With


End Sub

Private Sub cmdSClear_Click()
    
    On Error GoTo err
    If Not AskForAutoApply Then Exit Sub
    If cmbPymtmode.ListIndex = 0 Then
        If Not SaveTransForCash Then Exit Sub
    Else
        If Not SaveTransForCheque Then Exit Sub
    End If
    
    Unload Me
    
    Me.Show
    
    Me.SetFocus
    
'Call ClearFields
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSClose_Click()
    
    On Error GoTo err
    If Not AskForAutoApply Then Exit Sub
    
    If cmbPymtmode.ListIndex = 0 Then
        If Not SaveTransForCash Then Exit Sub
    Else
        If Not SaveTransForCheque Then Exit Sub
    End If
    
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdToBank_Click()
    Dim BankID As Double
    lblBankAcc = SelectBankAcc(BankID)
    lblBankAcc.Tag = BankID
End Sub

Private Sub DTPicker1_Change()

    Call setDateRange(DTPicker1)
    Call DTPicker2_Change
    Label1(4).Caption = getNextVchrNo(DTPicker1.Value, IIf(cmbPymtmode.ListIndex = 0, "CRV", "BRV"))
    If cmbAccNo.MatchFound Then Call RefreshList
    
End Sub

Private Sub DTPicker1_Click()
   Call DTPicker1_Change
End Sub



Private Sub DTPicker2_Change()
    Call setDateRange(DTPicker2)
    
    ChkClear.Visible = FraBank.Visible And DTPicker1 = DTPicker2 And cmbPymtmode.ListIndex = 1
    FClearChq.Visible = Frachkclear And ChkClear = vbChecked
    Call ChkClear_Click
End Sub



Private Sub Form_Activate()
'Check For User Access
If PermissionDenied Then
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   Exit Sub
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Form_Load()

    'PermissionDenied = Not UserHasAccess("ReceiptVoucher")

    Me.KeyPreview = True
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    Call DTPicker1_Click
    Call FillAccCmbs
'    If CurrentUserName = "Administrator" Then
'        DTPicker1.Enabled = True
'    Else
'        DTPicker1.Enabled = False
'    End If
    chkPrint = Abs(bPrintWhenSaved)
End Sub

Private Sub FillAccCmbs()

        'Clear All The Combos
        cmbAccNo.Clear
        CmbAccTitle.Clear
        cmbPymtmode.Clear
        cmbBank.Clear
        cmbBankBranch.Clear
        cmbCheckType.Clear
        
        'Fill With Payment Mode
        cmbPymtmode.AddItem "Cash"
        cmbPymtmode.AddItem "Cheque"
        cmbPymtmode.ListIndex = 0
        
        'Fill With CheckType
        cmbCheckType.AddItem "Open"
        cmbCheckType.AddItem "Payees A/C Only"
        cmbCheckType.ListIndex = 0
        
'        Call AddToCombo(cmbAccNo, "AccNo", "VActiveAccounts", "Where SubAccOf like '" & ACCRECNO & "%' Order By AccNo")
'        Call AddToCombo(CmbAccTitle, "AccTitle", "VActiveAccounts", "Where SubAccOf like '" & ACCRECNO & "%' Order By AccNo")
        
        Call AddToCombo(cmbAccNo, "AccNo", "VActiveAccounts", " WHERE AccNo<>'" & CashInHand & "' Order By AccNo")
        Call AddToCombo(CmbAccTitle, "AccTitle", "VActiveAccounts", " WHERE AccNo<>'" & CashInHand & "' Order By AccNo")
        
        Call AddToCombo(cmbBank, "Distinct Bank", "BankList", "Order By Bank")
      
        Call AddToCombo(cmbCashInHand, "AccTitle", "VActiveAccounts", "WHERE SubAccOf='" & CashInHand & "'", , "AccNo")
        If cmbCashInHand.ListCount > 0 Then cmbCashInHand.ListIndex = i
    
End Sub







Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
'If UnloadMode = 0 Then
'    If LVPymt.ListItems.Count > 0 Then
'        If MsgBox("There is a Transaction active." & vbNewLine & "Save or Cancel Transaction first.", vbInformation + vbOKOnly) Then
'            Cancel = True
'        End If
'    End If
'End If
End Sub









Private Sub Form_Resize()
    FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub


Private Sub List1_Click()
If UCase(Screen.ActiveControl.Name) <> "CMBACCTITLE" Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If
End Sub

Private Sub Opts_Click(Index As Integer)
Frame3.Visible = Opts(1)


If Index = 0 Then Exit Sub

        If Not AnyBankAccountDefined Then
            MsgBox "No Bank Accounts Defined.", vbInformation
            Opts(0) = True
            Exit Sub
        End If


        'Get Default BankID
        Dim BankID As Double
         lblBankAcc = GetDefaultBankAcc(BankID)
         lblBankAcc.Tag = BankID
         

End Sub

Private Sub Tbox_Change(Index As Integer)
'    If Index = 1 Then lblToApply = Tbox(1)
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 1 Then
    KeyAscii = OnlyNumber(KeyAscii, True)
End If

End Sub
Private Sub ClearLView()

Dim iTotal As Integer
iTotal = LView.ListItems.count
    For i = 1 To iTotal
        LView.ListItems.Remove (1)
    Next i
    
Tbox(0).Text = ""
Tbox(1).Text = ""
cmdDelete.Visible = False
Call DTPicker1_Click

End Sub





Private Sub LVPymt_ItemCheck(ByVal Item As MSComctlLib.ListItem)

If Not Item.Checked Then
    lblApplied = Format(Val(lblApplied) - Val(Item.SubItems(4)), "0.00")
    lblToApply = Format(Val(lblToApply) + Val(Item.SubItems(4)), "0.00")
    Item.SubItems(4) = 0
    Item.SubItems(5) = "DUE"
Else
 
    If Val(lblToApply) = 0 Then
        MsgBox "No Amount To Apply", vbInformation
        Item.Checked = False
        Exit Sub
    End If
 
    If Item.SubItems(3) <= Val(lblToApply) Then
        Item.SubItems(4) = Format(Val(Item.SubItems(3)), "0.00")
        Item.SubItems(5) = "FULL PAID"
   
    Else
        Item.SubItems(4) = Format(Val(lblToApply), "0.00")
        Item.SubItems(5) = "PARTIAL PAID"
    End If
  
    lblApplied = Format(Val(lblApplied) + Val(Item.SubItems(4)), "0.00")
    lblToApply = Format(Val(lblToApply) - Val(Item.SubItems(4)), "0.00")
  
End If

   If lblToApply = Tbox(1) Then
      cmd.Caption = "Auto Apply      "
   Else
      cmd.Caption = "Clear Payments    "
End If

'If Val(lblapply) = 0 Then
' cmdSave.Enabled = True
'Else
' cmdSave.Enabled = False
'End If

End Sub



Private Sub RefreshList()

Screen.MousePointer = vbHourglass
Call ShowAccBalance


Dim rsList As New ADODB.Recordset
rsList.CursorLocation = adUseClient
rsList.Open "Select Sno,VchrNo,Description,TotalAmount,AmountLeft  from VAccRec where (AccNo='" & cmbAccNo & "') AND (VDate<='" & DTPicker1 & "') and ((AmountLeft>0) or (AmountLeft is null))", con, adOpenForwardOnly, adLockReadOnly
 rsList.Sort = "Sno"
With LVPymt.ListItems
    .Clear
    cmd.Caption = "Auto Apply         "
    Dim ITM As ListItem
    Do Until rsList.EOF
        Set ITM = .add(, rsList![SNo] & "?", rsList![VchrNo])
        ITM.SubItems(1) = rsList![Description] & ""
        ITM.SubItems(2) = rsList![TotalAmount] & ""
        If IsNull(rsList![AmountLeft]) Then
            ITM.SubItems(3) = rsList![TotalAmount]
        Else
            ITM.SubItems(3) = rsList![AmountLeft] & ""
        End If
        
        ITM.SubItems(4) = 0
        ITM.SubItems(5) = "DUE"
        rsList.MoveNEXT
    Loop
    If .count > 0 Then cmd.Enabled = True Else cmd.Enabled = False
End With


'Reset The Amount
Call TBox_Validate(1, False)
Call ShowBals
Screen.MousePointer = vbDefault
End Sub

Private Function SaveTransForCash() As Boolean

    If CheckData = False Then Exit Function
    Dim strCashInHand As String
    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    
    'Get Serial ' and Balances
    INextSno = getNextSno(DTPicker1.Value)
    SnoToUse = INextSno
    Bal = getBalance(cmbAccNo, DTPicker1.Value)
    Bal = Bal - Val(Tbox(1))
    'Start New Transaction
    con.BeginTrans
    'Credit Entry
    con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
        Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & cmbAccNo & _
        "','" & Tbox(0) & "',0," & Val(Tbox(1)) & "," & Bal & ",'',0)"
    
    'Con.Execute "Update Accounts Set Balance=Balance+" & -Val(Tbox(1)) & " Where AccNo='" & cmbAccNo & "'"
          
   Call adjustLedger(cmbAccNo, DTPicker1.Value, INextSno, -Val(Tbox(1)))
                    
    INextSno = INextSno + 1
    strCashInHand = cmbCashInHand.List(cmbCashInHand.ListIndex, 1)
    Bal = getBalance(strCashInHand, DTPicker1.Value)
    Bal = Bal + Val(Tbox(1))
    'Debit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & strCashInHand & _
        "','" & Tbox(0) & "'," & Val(Tbox(1)) & _
        ",0," & Bal & ",'',0)"
            
    Call adjustLedger(strCashInHand, DTPicker1.Value, INextSno, Val(Tbox(1)))
        
    With LVPymt.ListItems
        For i = 1 To .count
            If .Item(i).Checked Then con.Execute "Insert into VoucherDetail Values('" & SnoToUse & "','" & Val(.Item(i).key) & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
        Next
    End With
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    'Commit Transaction
    con.CommitTrans
    
    LVPymt.ListItems.Clear
    'Call FillAccCmbs
    Call PlaySaveSound
    
    'Print If To Be Printed
    If chkPrint = vbChecked Then Call PrintVoucher(MainForm.cr1, Label1(4), True)
    
    MsgBox "Transaction Saved Successfuly.", vbInformation
    
    'FillAccCmbs
    SaveTransForCash = True
    Exit Function
err:
    MsgBox err.Description
    
End Function


Private Function CheckData() As Boolean

On Error GoTo err

'If LVPymt.ListItems.Count = 0 Then
'    MsgBox "No Transaction to save!", vbInformation
'    Exit Function
'End If
    
If cmd.Caption = "Auto A&pply" And cmd.Enabled Then
    If MsgBox("No Amount is Applied. Do You Want To Apply Automatically.", vbQuestion + vbYesNo) = vbYes Then
        Call cmd_Click
    Else
        Exit Function
    End If
End If

If (cmbAccNo = "") Or (Not cmbAccNo.MatchFound) Then
    MsgBox "Invalid Account Selected...", vbInformation
    cmbAccNo.SetFocus
    Exit Function
ElseIf Tbox(0) = "" Then
    MsgBox "No Description Given...", vbInformation
    Tbox(0).SetFocus
    Exit Function
ElseIf Val(Tbox(1)) = 0 Then
    MsgBox "No Amount Given...", vbInformation
    Tbox(1).SetFocus
    Exit Function

End If

If cmbPymtmode.ListIndex = 1 Then
    If cmbBank = "" Or cmbBank.MatchFound = False Then
        MsgBox "No Bank Selected...", vbInformation
        cmbBank.SetFocus
        Exit Function
    ElseIf cmbBankBranch = "" Or cmbBankBranch.MatchFound = False Then
        MsgBox "No Bank Branch Selected...", vbInformation
        cmbBankBranch.SetFocus
        Exit Function
    ElseIf Tbox(2) = "" Then
        MsgBox "No Cheque # Given...", vbInformation
        Tbox(2).SetFocus
        Exit Function
    End If
Else
    If cmbCashInHand.MatchFound = False Then
        MsgBox "Please Select Cash In Hand Account.", vbInformation
        Exit Function
    End If
End If

If ChkClear.Visible And ChkClear = vbChecked And Opts(1) And txtSlipNo = "" Then
    MsgBox "Please Enter Deposit Slip No.", vbInformation
    CheckData = False
    Exit Function
End If



''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''

Dim VNo As String
VNo = getNextVchrNo(DTPicker1.Value, IIf(cmbPymtmode.ListIndex = 0, "CRV", "BRV"))
If Label1(4) <> VNo Then
   MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
       "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
       Label1(4) = VNo
End If



CheckData = True

Exit Function



err:
    MsgBox err.Description
    
End Function



Private Function SaveTransForCheque() As Boolean

    If CheckData = False Then Exit Function

    
    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    
    'Get Serial # and Balances
    INextSno = getNextSno(DTPicker1.Value, "Cheque")
    SnoToUse = INextSno
    
    'Start New Transaction
    con.BeginTrans
    
    con.Execute "INSERT INTO Cheque(SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced,BankAccNo) Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
     cmbAccNo & "','" & Tbox(0) & "'," & Tbox(1) & "," & Val(cmbBankBranch.Tag) & ",'" & _
     Tbox(2) & "',0,'" & cmbCheckType & "','" & DTPicker2 & "',0,0,0,'" & GetID(lblBankAcc) & "')"
    
    With LVPymt.ListItems
        For i = 1 To .count
            If .Item(i).Checked Then con.Execute "INSERT INTO ChequeDetail VALUES('" & SnoToUse & "','" & Val(.Item(i).key) & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
        Next
    End With
    
    
    If ChkClear.Visible And ChkClear = vbChecked Then
        If Not SaveReceiptTrans(IIf(Opts(0), CashInHand, GetID(lblBankAcc)), SnoToUse, txtSlipNo) Then
            MsgBox "Transaction Terminated Cannot Save Due To Error :" & err.Number & vbNewLine & err.Description, vbCritical
            con.RollbackTrans
        End If
    End If
    
    
    'Commit Transaction
    con.CommitTrans
    
    
    'Call FillAccCmbs
    Call PlaySaveSound
    MsgBox "Transaction Saved Successfuly.", vbInformation
    
    'FillAccCmbs
    SaveTransForCheque = True
    Exit Function
err:
    MsgBox "Transaction Terminated Cannot Save Due To Error :" & err.Number & vbNewLine & err.Description, vbCritical
    SaveTransForCheque = False
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
End Function


Private Sub ClearFields()
    Label1(3).Visible = True
    Label1(4).Visible = True
    Call FillAccCmbs
    cmbPymtmode.ListIndex = 0
    Tbox(0) = "Cash Recieved"
    Tbox(1) = ""
    Tbox(2) = ""
End Sub

Private Function AskForAutoApply() As Boolean

If Trim(cmd.Caption) = "Auto Apply" And LVPymt.ListItems.count > 0 Then
   
   If MsgBox("No Amounts Applied,Do You Want To Auto Apply", vbYesNo + vbQuestion) = vbYes Then
      Call cmd_Click
      AskForAutoApply = True
   End If
   
Else
      AskForAutoApply = True
End If

'If Val(lblToApply) <> 0 And LVPymt.ListItems.Count > 0 Then
'      MsgBox "There is Some Unapplied Amount, Cannot Save", vbInformation
'      AskForAutoApply = False
'End If
End Function



Private Sub ShowAccBalance()
'Show Previous Balance of Account
If lblBalance.Visible Then
    lblBalance.Caption = Format(getBalance(cmbAccNo, DTPicker1), "#,##0.00 DR;#,##0.00 CR;0.00 ")
    lblBalance.ToolTipText = "Balance of '" & CmbAccTitle & "' As Of " & Format(DTPicker1.Value, "dd/MM/yyyy")
    
End If

End Sub


Private Sub TBox_Validate(Index As Integer, cancel As Boolean)
If Index <> 1 Then Exit Sub
  lblApplied = 0
  lblToApply = Tbox(1)
' If Val(Tbox(1)) <> Val(lblToApply) + Val(lblApplied) And Trim(cmd.Caption) = "Clear Payments" Then
 If Trim(cmd.Caption) = "Clear Payments" Then
   Call cmd_Click
   Call cmd_Click
 End If
End Sub



Private Sub ShowBals()
    lblBalCap = "Cash In Hand  "
    If cmbCashInHand.MatchFound Then
        lblBal = Format(getBalance(cmbCashInHand.List(cmbCashInHand.ListIndex, 1), DTPicker1), "0.00")
    End If
End Sub


Private Function SaveReceiptTrans(ReceiveInAccNo As String, ChqSNo As Double, Optional SlipNo As String) As Boolean

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    
    
    Dim Dated As Date
    Dated = DTPicker2
    'Get Serial # and Balances
    NextVchrNo = getNextVchrNo(Dated, IIf(cmbPymtmode.ListIndex = 0, "CRV", "BRV"))
    INextSno = getNextSno(Dated)
    SnoToUse = INextSno
    Bal = getBalance(cmbAccNo, Dated)
    dAmount = Val(Tbox(1))
    sScNo = ChqSNo
    sCIH = ReceiveInAccNo
    Bal = Bal - dAmount
    
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & cmbAccNo & _
        "','" & Tbox(0) & "',0," & dAmount & "," & Bal & ",'" & SlipNo & "'," & sScNo & ")"
    
    
    
    Call adjustLedger(cmbAccNo, Dated, INextSno, -Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    Bal = Bal + dAmount
    'Description
    
    
       Desc = "Cheque Recieved from " & CmbAccTitle.Tag & " Chq# : " & Tbox(2)
    
        
    
    
    'Debit Entry
    
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & "'," & dAmount & ",0," & Bal & ",'" & SlipNo & "'," & sScNo & ")"
        
    
    
    Call adjustLedger(sCIH, Dated, INextSno, Val(dAmount))
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "Insert into VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    con.Execute "Update Cheque Set Posted=1 Where Sno=" & sScNo
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    SaveReceiptTrans = True
    Exit Function

err:
    
    SaveReceiptTrans = False
End Function

