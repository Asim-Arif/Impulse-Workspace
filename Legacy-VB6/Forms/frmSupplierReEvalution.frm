VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSupplierReEvalution 
   Caption         =   "Supplier Re-Evalution"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9315
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   16431
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Evalution"
      TabPicture(0)   =   "frmSupplierReEvalution.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblAproved"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEvalute"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPrepaid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblRemar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbTimeAprov"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkRecmaintain"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkDrawing"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkCalibration"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkTracebilityControl"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkClean"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkUPS"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkSampling"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkGenerator"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblSelection"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMajCust"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCapacity"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtRemarks2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtRemarks1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label6"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblCapacity"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblRemarks2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblRemarks1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtCoreCompetencies"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtProductLine"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtToolNo"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtMachineNo"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtEmpNo"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblCoreCompetencies"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label1"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblMachine"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lblNoofEmployee"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmbSelectMkrVend"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtPreparedBy"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtApprovedBy"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "CD1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "FVendorOnly"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmbEvaluatedBy"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtRemarks"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmdBrowseBarcode"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtFilePath"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Frame1"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "frmMarking"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "Re-Evalution"
      TabPicture(1)   =   "frmSupplierReEvalution.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label38"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label46"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ComboBox1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label47"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label48"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label49"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label50"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TextBox1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TextBox2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "CommonDialog1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "FQuality"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame4"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Command1"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text2"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Combo1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3840
         TabIndex        =   104
         Tag             =   "EvaluatedBy"
         Top             =   8910
         Width           =   2820
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   100
         Tag             =   "PDFFile"
         Top             =   8235
         Width           =   4590
      End
      Begin VB.CommandButton Command1 
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
         Height          =   360
         Left            =   9885
         TabIndex        =   99
         Top             =   8220
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   900
         Left            =   105
         TabIndex        =   98
         Tag             =   "Remarks"
         Top             =   7050
         Width           =   10200
      End
      Begin VB.Frame Frame4 
         Height          =   960
         Left            =   75
         TabIndex        =   90
         Top             =   4545
         Width           =   10215
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   11
            Left            =   3780
            TabIndex        =   121
            Tag             =   "QCancelledOrderComm"
            Top             =   585
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   10
            Left            =   3780
            TabIndex        =   120
            Tag             =   "QNoofOrderCompletedComm"
            Top             =   255
            Width           =   1680
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   10
            Left            =   5775
            TabIndex        =   92
            Tag             =   "EEnvironmentalRequirementsComm"
            Top             =   255
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   11
            Left            =   5775
            TabIndex        =   91
            Tag             =   "EVehiclesComm"
            Top             =   585
            Width           =   4305
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   " Environmental System "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   95
            Top             =   0
            Width           =   2010
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   " Environmental Requirements"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   94
            Top             =   255
            Width           =   2520
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   " Lighting, Ventilation and Vehicles"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   93
            Top             =   585
            Width           =   2820
         End
      End
      Begin VB.Frame frmMarking 
         Height          =   525
         Left            =   -74805
         TabIndex        =   46
         Top             =   4200
         Width           =   4920
         Begin VB.OptionButton optGood 
            Caption         =   "Good"
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
            Index           =   0
            Left            =   2070
            TabIndex        =   49
            Tag             =   "ReputationGood"
            Top             =   195
            Width           =   885
         End
         Begin VB.OptionButton optAvg 
            Caption         =   "Average"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2955
            TabIndex        =   48
            Tag             =   "ReputationAvg"
            Top             =   195
            Width           =   1320
         End
         Begin VB.OptionButton optPoor 
            Caption         =   "Poor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   4065
            TabIndex        =   47
            Tag             =   "ReputationPoor"
            Top             =   195
            Width           =   705
         End
         Begin MSForms.Label Label3 
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   195
            Width           =   1890
            VariousPropertyBits=   276824091
            Caption         =   "Marketing Reputation"
            Size            =   "3334;344"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.Frame Frame1 
         Height          =   525
         Left            =   -69780
         TabIndex        =   41
         Top             =   4200
         Width           =   4830
         Begin VB.OptionButton optPoor 
            Caption         =   "Poor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   3990
            TabIndex        =   44
            Tag             =   "CapabilityPoor"
            Top             =   195
            Width           =   705
         End
         Begin VB.OptionButton optAvg 
            Caption         =   "Average"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   2865
            TabIndex        =   43
            Tag             =   "CapabilityAvg"
            Top             =   195
            Width           =   1035
         End
         Begin VB.OptionButton optGood 
            Caption         =   "Good"
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
            Index           =   3
            Left            =   1980
            TabIndex        =   42
            Tag             =   "CapabilityGood"
            Top             =   195
            Width           =   885
         End
         Begin MSForms.Label Label4 
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   195
            Width           =   1695
            VariousPropertyBits=   276824091
            Caption         =   "Technical Capability"
            Size            =   "2990;344"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
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
         Height          =   345
         Left            =   -69810
         TabIndex        =   40
         Tag             =   "PDFFile"
         Top             =   5190
         Width           =   4455
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
         Height          =   360
         Left            =   -65340
         TabIndex        =   39
         Top             =   5175
         Width           =   405
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1470
         Left            =   -74790
         TabIndex        =   38
         Tag             =   "Remarks"
         Top             =   5895
         Width           =   9855
      End
      Begin VB.ComboBox cmbEvaluatedBy 
         Height          =   315
         Left            =   -71235
         TabIndex        =   37
         Tag             =   "EvaluatedBy"
         Top             =   8820
         Width           =   2820
      End
      Begin VB.Frame FVendorOnly 
         Height          =   765
         Left            =   -74790
         TabIndex        =   32
         Top             =   7575
         Width           =   9855
         Begin VB.CheckBox chkISO 
            Caption         =   "ISO 9000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   165
            TabIndex        =   35
            Tag             =   "ISO9000"
            Top             =   315
            Width           =   1665
         End
         Begin VB.CheckBox chkCustComplaint 
            Caption         =   "Customer Complaint Handling"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3075
            TabIndex        =   34
            Tag             =   "CustComplaintHanding"
            Top             =   315
            Width           =   2985
         End
         Begin VB.CheckBox chkStaffTrain 
            Caption         =   "Staff Training System"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7305
            TabIndex        =   33
            Tag             =   "StaffTrainingSystem"
            Top             =   315
            Width           =   2250
         End
         Begin MSForms.Label Label8 
            Height          =   195
            Left            =   15
            TabIndex        =   36
            Top             =   0
            Width           =   2820
            VariousPropertyBits=   276824091
            Caption         =   " Additional incase of Vendor Only:"
            Size            =   "4974;344"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.Frame FQuality 
         Height          =   2925
         Left            =   75
         TabIndex        =   14
         Top             =   570
         Width           =   10215
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   7
            Left            =   3750
            TabIndex        =   117
            Tag             =   "QCancelledOrderRatting"
            Top             =   2565
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   6
            Left            =   3750
            TabIndex        =   116
            Tag             =   "QNoofOrderCompletedRatting"
            Top             =   2235
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   5
            Left            =   3750
            TabIndex        =   115
            Tag             =   "QReWorkRateRatting"
            Top             =   1905
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   4
            Left            =   3750
            TabIndex        =   114
            Tag             =   "QDeliveryInTimeRatting"
            Top             =   1575
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   3
            Left            =   3750
            TabIndex        =   113
            Tag             =   "QRejectionRateRatting"
            Top             =   1245
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   2
            Left            =   3750
            TabIndex        =   112
            Tag             =   "QPackingRatting"
            Top             =   915
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   1
            Left            =   3750
            TabIndex        =   111
            Tag             =   "QPriceRatting"
            Top             =   585
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   0
            Left            =   3750
            TabIndex        =   110
            Tag             =   "QQualityoProductRatting"
            Top             =   255
            Width           =   1680
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   0
            Left            =   5775
            TabIndex        =   22
            Tag             =   "QQualityoProductComm"
            Top             =   255
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   1
            Left            =   5775
            TabIndex        =   21
            Tag             =   "QPriceComm"
            Top             =   585
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   2
            Left            =   5775
            TabIndex        =   20
            Tag             =   "QPackingComm"
            Top             =   915
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   3
            Left            =   5775
            TabIndex        =   19
            Tag             =   "QRejectionRateComm"
            Top             =   1245
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   4
            Left            =   5775
            TabIndex        =   18
            Tag             =   "QDeliveryInTimeComm"
            Top             =   1575
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   5
            Left            =   5775
            TabIndex        =   17
            Tag             =   "QReWorkRateComm"
            Top             =   1905
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   6
            Left            =   5775
            TabIndex        =   16
            Tag             =   "QNoofOrderCompletedComm"
            Top             =   2235
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   7
            Left            =   5775
            TabIndex        =   15
            Tag             =   "QCancelledOrderComm"
            Top             =   2565
            Width           =   4305
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   " Quallity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   31
            Top             =   0
            Width           =   690
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   " Quality of Product"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   30
            Top             =   255
            Width           =   1560
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   " Price"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   29
            Top             =   585
            Width           =   465
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   " Packing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   28
            Top             =   930
            Width           =   705
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   " Rejection Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   27
            Top             =   1260
            Width           =   1305
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   " Delivety In-Time"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   26
            Top             =   1590
            Width           =   1470
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   " Re-Work Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   25
            Top             =   1920
            Width           =   1245
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   " No. of Order Completed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   24
            Top             =   2265
            Width           =   1995
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   " Cancelled Orders"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   23
            Top             =   2595
            Width           =   1470
         End
      End
      Begin VB.Frame Frame2 
         Height          =   960
         Left            =   75
         TabIndex        =   8
         Top             =   3555
         Width           =   10215
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   9
            Left            =   3765
            TabIndex        =   119
            Tag             =   "PProductionCapacityRatting"
            Top             =   600
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   8
            Left            =   3765
            TabIndex        =   118
            Tag             =   "PTechnicalAbilityRatting"
            Top             =   270
            Width           =   1680
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   9
            Left            =   5775
            TabIndex        =   10
            Tag             =   "PProductionCapacityComm"
            Top             =   585
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   8
            Left            =   5775
            TabIndex        =   9
            Tag             =   "PTechnicalAbilityComm"
            Top             =   255
            Width           =   4305
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   " Production Capacity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   13
            Top             =   585
            Width           =   1740
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   " Technical Ability"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   12
            Top             =   255
            Width           =   1425
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   " Performance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   11
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1305
         Left            =   90
         TabIndex        =   1
         Top             =   5550
         Width           =   10215
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   14
            Left            =   3780
            TabIndex        =   124
            Tag             =   "OAttitudeRatting"
            Top             =   915
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   13
            Left            =   3780
            TabIndex        =   123
            Tag             =   "EVehiclesRatting"
            Top             =   585
            Width           =   1680
         End
         Begin VB.TextBox txtRatting 
            Height          =   285
            Index           =   12
            Left            =   3780
            TabIndex        =   122
            Tag             =   "EEnvironmentalRequirementsRatting"
            Top             =   255
            Width           =   1680
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   12
            Left            =   5775
            TabIndex        =   4
            Tag             =   "OAttitudeComm"
            Top             =   255
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   13
            Left            =   5775
            TabIndex        =   3
            Tag             =   "OCustomerServicesComm"
            Top             =   585
            Width           =   4305
         End
         Begin VB.TextBox txtComments 
            Height          =   285
            Index           =   14
            Left            =   5775
            TabIndex        =   2
            Tag             =   "OUrgentServoicesComm"
            Top             =   915
            Width           =   4305
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   " Urgent Servoices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   435
            TabIndex        =   96
            Top             =   915
            Width           =   1500
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   " Others"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   7
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   " Attitude"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   6
            Top             =   255
            Width           =   750
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   " Customer Services"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   5
            Top             =   585
            Width           =   1635
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   -66195
         Top             =   4995
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "PDF Files Only|*.pdf"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8865
         Top             =   8040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "PDF Files Only|*.pdf"
      End
      Begin MSForms.TextBox TextBox2 
         Height          =   270
         Left            =   7410
         TabIndex        =   109
         Tag             =   "ApprovedBy"
         Top             =   8895
         Width           =   2895
         VariousPropertyBits=   746604571
         Size            =   "5106;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   270
         Left            =   105
         TabIndex        =   108
         Tag             =   "PreparedBy"
         Top             =   8895
         Width           =   2820
         VariousPropertyBits=   746604571
         Size            =   "4974;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label50 
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   8670
         Width           =   1140
         VariousPropertyBits=   276824091
         Caption         =   " Prepared By:"
         Size            =   "2011;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label49 
         Height          =   195
         Left            =   3825
         TabIndex        =   106
         Top             =   8670
         Width           =   1200
         VariousPropertyBits=   276824091
         Caption         =   " Evaluated By:"
         Size            =   "2117;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label48 
         Height          =   195
         Left            =   7395
         TabIndex        =   105
         Top             =   8670
         Width           =   1320
         VariousPropertyBits=   276824091
         Caption         =   " Approved By:"
         Size            =   "2328;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label47 
         Height          =   195
         Left            =   150
         TabIndex        =   103
         Top             =   8295
         Width           =   1245
         VariousPropertyBits=   276824091
         Caption         =   " Approved For:"
         Size            =   "2196;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   330
         Left            =   1470
         TabIndex        =   102
         Tag             =   "ApprovalFor"
         Top             =   8235
         Width           =   3600
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "6350;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label46 
         Height          =   195
         Left            =   5280
         TabIndex        =   101
         Top             =   7995
         Width           =   1905
         VariousPropertyBits=   276824091
         Caption         =   " PDF Path to Attached:"
         Size            =   "3360;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   " Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   97
         Top             =   6840
         Width           =   810
      End
      Begin MSForms.TextBox txtApprovedBy 
         Height          =   270
         Left            =   -67680
         TabIndex        =   89
         Tag             =   "ApprovedBy"
         Top             =   8805
         Width           =   2760
         VariousPropertyBits=   746604571
         Size            =   "4868;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPreparedBy 
         Height          =   270
         Left            =   -74790
         TabIndex        =   88
         Tag             =   "PreparedBy"
         Top             =   8805
         Width           =   2820
         VariousPropertyBits=   746604571
         Size            =   "4974;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbSelectMkrVend 
         Height          =   270
         Left            =   -74805
         TabIndex        =   87
         Tag             =   "AccNo"
         Top             =   810
         Width           =   4890
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "8625;476"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblNoofEmployee 
         Height          =   195
         Left            =   -74805
         TabIndex        =   86
         Top             =   1260
         Width           =   1395
         VariousPropertyBits=   276824091
         Caption         =   " No. of Employee"
         Size            =   "2461;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblMachine 
         Height          =   195
         Left            =   -74805
         TabIndex        =   85
         Top             =   1650
         Width           =   1365
         VariousPropertyBits=   276824091
         Caption         =   " No. of Machines"
         Size            =   "2408;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Left            =   -74805
         TabIndex        =   84
         Top             =   2040
         Width           =   1020
         VariousPropertyBits=   276824091
         Caption         =   " No. of Tools"
         Size            =   "1799;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   195
         Left            =   -74805
         TabIndex        =   83
         Top             =   2415
         Width           =   1110
         VariousPropertyBits=   276824091
         Caption         =   " Product Line"
         Size            =   "1958;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblCoreCompetencies 
         Height          =   195
         Left            =   -74805
         TabIndex        =   82
         Top             =   2805
         Width           =   1740
         VariousPropertyBits=   276824091
         Caption         =   " Core. Competencies"
         Size            =   "3069;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtEmpNo 
         Height          =   270
         Left            =   -73020
         TabIndex        =   81
         Tag             =   "EmpNo"
         Top             =   1215
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtMachineNo 
         Height          =   270
         Left            =   -73020
         TabIndex        =   80
         Tag             =   "MachineNo"
         Top             =   1605
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtToolNo 
         Height          =   270
         Left            =   -73035
         TabIndex        =   79
         Tag             =   "ToolNo"
         Top             =   1995
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProductLine 
         Height          =   270
         Left            =   -73035
         TabIndex        =   78
         Tag             =   "ProductLine"
         Top             =   2385
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtCoreCompetencies 
         Height          =   270
         Left            =   -73035
         TabIndex        =   77
         Tag             =   "CoreCompetencies"
         Top             =   2790
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblRemarks1 
         Height          =   195
         Left            =   -69615
         TabIndex        =   76
         Top             =   1665
         Width           =   825
         VariousPropertyBits=   276824091
         Caption         =   " Remarks"
         Size            =   "1455;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblRemarks2 
         Height          =   195
         Left            =   -69615
         TabIndex        =   75
         Top             =   2040
         Width           =   825
         VariousPropertyBits=   276824091
         Caption         =   " Remarks"
         Size            =   "1455;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblCapacity 
         Height          =   195
         Left            =   -69615
         TabIndex        =   74
         Top             =   2430
         Width           =   795
         VariousPropertyBits=   276824091
         Caption         =   " Capacity"
         Size            =   "1402;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label6 
         Height          =   195
         Left            =   -69615
         TabIndex        =   73
         Top             =   2820
         Width           =   1515
         VariousPropertyBits=   276824091
         Caption         =   " Major Customers"
         Size            =   "2672;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtRemarks1 
         Height          =   270
         Left            =   -68040
         TabIndex        =   72
         Tag             =   "Remarks1"
         Top             =   1620
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtRemarks2 
         Height          =   270
         Left            =   -68040
         TabIndex        =   71
         Tag             =   "Remarks2"
         Top             =   2010
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtCapacity 
         Height          =   270
         Left            =   -68040
         TabIndex        =   70
         Tag             =   "Capacity"
         Top             =   2415
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtMajCust 
         Height          =   270
         Left            =   -68040
         TabIndex        =   69
         Tag             =   "MajorCustomers"
         Top             =   2805
         Width           =   3090
         VariousPropertyBits=   746604571
         Size            =   "5450;476"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSelection 
         Height          =   195
         Left            =   -74820
         TabIndex        =   68
         Top             =   540
         Width           =   1995
         VariousPropertyBits=   276824091
         Caption         =   " Select Maker / Vendor:"
         Size            =   "3519;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkGenerator 
         Height          =   345
         Left            =   -74805
         TabIndex        =   67
         Tag             =   "Generator"
         Top             =   3255
         Width           =   1275
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2249;609"
         Value           =   "0"
         Caption         =   "Generator"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkSampling 
         Height          =   345
         Left            =   -74805
         TabIndex        =   66
         Tag             =   "SamplingHandling"
         Top             =   3690
         Width           =   1965
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3466;609"
         Value           =   "0"
         Caption         =   "Sampling Handling"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkUPS 
         Height          =   345
         Left            =   -72555
         TabIndex        =   65
         Tag             =   "UPS"
         Top             =   3255
         Width           =   735
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1296;609"
         Value           =   "0"
         Caption         =   "UPS"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkClean 
         Height          =   345
         Left            =   -72555
         TabIndex        =   64
         Tag             =   "Cleaniness"
         Top             =   3690
         Width           =   1305
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2302;609"
         Value           =   "0"
         Caption         =   "Cleaniness"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkTracebilityControl 
         Height          =   345
         Left            =   -70995
         TabIndex        =   63
         Tag             =   "TracebilityControl"
         Top             =   3690
         Width           =   3330
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "5874;609"
         Value           =   "0"
         Caption         =   "Tracebility & Identification Control"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkCalibration 
         Height          =   345
         Left            =   -67530
         TabIndex        =   62
         Tag             =   "CalibrationSystem"
         Top             =   3255
         Width           =   2010
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3545;609"
         Value           =   "0"
         Caption         =   "Calibration System"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkDrawing 
         Height          =   345
         Left            =   -67530
         TabIndex        =   61
         Tag             =   "DrawingComprehension"
         Top             =   3690
         Width           =   2580
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4551;609"
         Value           =   "0"
         Caption         =   "T/Drawing Comprehension"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkRecmaintain 
         Height          =   345
         Left            =   -70980
         TabIndex        =   60
         Tag             =   "InspRecMaintain"
         Top             =   3255
         Width           =   2730
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "4815;609"
         Value           =   "0"
         Caption         =   "Inspection Record Maintain"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label5 
         Height          =   195
         Left            =   -74805
         TabIndex        =   59
         Top             =   5250
         Width           =   1575
         VariousPropertyBits=   276824091
         Caption         =   " Time for Approval "
         Size            =   "2778;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbTimeAprov 
         Height          =   330
         Left            =   -73035
         TabIndex        =   58
         Tag             =   "TimeApproval"
         Top             =   5190
         Width           =   3150
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5556;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   195
         Left            =   -69780
         TabIndex        =   57
         Top             =   4950
         Width           =   1905
         VariousPropertyBits=   276824091
         Caption         =   " PDF Path to Attached:"
         Size            =   "3360;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblRemar 
         Height          =   195
         Left            =   -74790
         TabIndex        =   56
         Top             =   5640
         Width           =   870
         VariousPropertyBits=   276824091
         Caption         =   " Remarks:"
         Size            =   "1535;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblPrepaid 
         Height          =   195
         Left            =   -74775
         TabIndex        =   55
         Top             =   8580
         Width           =   1140
         VariousPropertyBits=   276824091
         Caption         =   " Prepared By:"
         Size            =   "2011;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblEvalute 
         Height          =   195
         Left            =   -71250
         TabIndex        =   54
         Top             =   8580
         Width           =   1200
         VariousPropertyBits=   276824091
         Caption         =   " Evaluated By:"
         Size            =   "2117;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblAproved 
         Height          =   195
         Left            =   -67695
         TabIndex        =   53
         Top             =   8580
         Width           =   1185
         VariousPropertyBits=   276824091
         Caption         =   " Approved By:"
         Size            =   "2090;344"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   " Ratting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4290
         TabIndex        =   52
         Top             =   375
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   " Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7455
         TabIndex        =   51
         Top             =   375
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmSupplierReEvalution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

