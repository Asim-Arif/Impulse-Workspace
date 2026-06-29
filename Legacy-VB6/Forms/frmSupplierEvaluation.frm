VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSupplierEvaluation 
   Caption         =   "Supplier Evaluation"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
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
      Height          =   405
      Left            =   6645
      TabIndex        =   57
      Top             =   7815
      Width           =   1605
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
      Height          =   405
      Left            =   8325
      TabIndex        =   56
      Top             =   7815
      Width           =   1605
   End
   Begin VB.Frame FVendorOnly 
      Height          =   765
      Left            =   75
      TabIndex        =   43
      Top             =   6285
      Width           =   9855
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
         TabIndex        =   47
         Tag             =   "StaffTrainingSystem"
         Top             =   315
         Width           =   2250
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
         TabIndex        =   46
         Tag             =   "CustComplaintHanding"
         Top             =   315
         Width           =   2985
      End
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
         TabIndex        =   44
         Tag             =   "ISO9000"
         Top             =   315
         Width           =   1665
      End
      Begin MSForms.Label Label8 
         Height          =   195
         Left            =   15
         TabIndex        =   45
         Top             =   -15
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
   Begin VB.ComboBox cmbEvaluatedBy 
      Height          =   315
      Left            =   3630
      TabIndex        =   42
      Tag             =   "EvaluatedBy"
      Top             =   7380
      Width           =   2820
   End
   Begin VB.TextBox txtRemarks 
      Height          =   1185
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   38
      Tag             =   "Remarks"
      Top             =   5055
      Width           =   9855
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
      Left            =   9525
      TabIndex        =   35
      Top             =   4500
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
      Height          =   345
      Left            =   5070
      TabIndex        =   34
      Top             =   4515
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   5085
      TabIndex        =   30
      Tag             =   "-1"
      Top             =   3720
      Width           =   4830
      Begin VB.OptionButton OptTech 
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
         Height          =   270
         Index           =   2
         Left            =   3795
         TabIndex        =   55
         Top             =   165
         Width           =   795
      End
      Begin VB.OptionButton OptTech 
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   2940
         TabIndex        =   54
         Top             =   165
         Width           =   795
      End
      Begin VB.OptionButton OptTech 
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
         Height          =   270
         Index           =   0
         Left            =   2070
         TabIndex        =   53
         Tag             =   "TechnicalCapability"
         Top             =   165
         Width           =   795
      End
      Begin MSForms.Label lblTechnicalCapability 
         Height          =   195
         Left            =   120
         TabIndex        =   31
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
   Begin VB.Frame frmMarking 
      Height          =   525
      Left            =   60
      TabIndex        =   28
      Top             =   3720
      Width           =   4920
      Begin VB.OptionButton OptMarketting 
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
         Height          =   270
         Index           =   2
         Left            =   4020
         TabIndex        =   52
         Top             =   165
         Width           =   795
      End
      Begin VB.OptionButton OptMarketting 
         Caption         =   "Avg."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   3150
         TabIndex        =   51
         Top             =   165
         Width           =   795
      End
      Begin VB.OptionButton OptMarketting 
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
         Height          =   270
         Index           =   0
         Left            =   2265
         TabIndex        =   50
         Tag             =   "MarketingReputation"
         Top             =   165
         Width           =   795
      End
      Begin MSForms.Label lblMarketingReputation 
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Tag             =   "-1"
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
   Begin MSComDlg.CommonDialog CD1 
      Left            =   9450
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
   End
   Begin MSForms.TextBox txtApprovedBy 
      Height          =   315
      Left            =   7185
      TabIndex        =   49
      Tag             =   "ApprovedBy"
      Top             =   7365
      Width           =   2760
      VariousPropertyBits=   746604571
      Size            =   "4868;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPreparedBy 
      Height          =   315
      Left            =   75
      TabIndex        =   48
      Tag             =   "PreparedBy"
      Top             =   7365
      Width           =   2820
      VariousPropertyBits=   746604571
      Size            =   "4974;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblAproved 
      Height          =   195
      Left            =   7170
      TabIndex        =   41
      Top             =   7140
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
   Begin MSForms.Label lblEvalute 
      Height          =   195
      Left            =   3615
      TabIndex        =   40
      Top             =   7140
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
   Begin MSForms.Label lblPrepaid 
      Height          =   270
      Left            =   90
      TabIndex        =   39
      Top             =   7140
      Width           =   1230
      VariousPropertyBits=   276824091
      Caption         =   " Prepared By:"
      Size            =   "2170;476"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblRemar 
      Height          =   195
      Left            =   45
      TabIndex        =   37
      Top             =   4860
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
   Begin MSForms.Label Label7 
      Height          =   195
      Left            =   5085
      TabIndex        =   36
      Top             =   4275
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
   Begin MSForms.ComboBox cmbTimeApprove 
      Height          =   330
      Left            =   60
      TabIndex        =   33
      Tag             =   "TimeApproval"
      Top             =   4515
      Width           =   4920
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "8678;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   195
      Left            =   60
      TabIndex        =   32
      Top             =   4275
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
   Begin MSForms.CheckBox chkRecmaintain 
      Height          =   345
      Left            =   3885
      TabIndex        =   24
      Tag             =   "InspRecMaintain"
      Top             =   2775
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
   Begin MSForms.CheckBox chkDrawing 
      Height          =   345
      Left            =   7335
      TabIndex        =   27
      Tag             =   "DrawingComprehension"
      Top             =   3210
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
   Begin MSForms.CheckBox chkCalibration 
      Height          =   345
      Left            =   7335
      TabIndex        =   26
      Tag             =   "CalibrationSystem"
      Top             =   2775
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
   Begin MSForms.CheckBox chkTracebilityControl 
      Height          =   345
      Left            =   3870
      TabIndex        =   25
      Tag             =   "TracebilityControl"
      Top             =   3210
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
   Begin MSForms.CheckBox chkClean 
      Height          =   345
      Left            =   2310
      TabIndex        =   23
      Tag             =   "Cleaniness"
      Top             =   3210
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
   Begin MSForms.CheckBox chkUPS 
      Height          =   345
      Left            =   2310
      TabIndex        =   22
      Tag             =   "UPS"
      Top             =   2775
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
   Begin MSForms.CheckBox chkSampling 
      Height          =   345
      Left            =   60
      TabIndex        =   21
      Tag             =   "SamplingHandling"
      Top             =   3210
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
   Begin MSForms.CheckBox chkGenerator 
      Height          =   345
      Left            =   60
      TabIndex        =   20
      Tag             =   "Generator"
      Top             =   2775
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
   Begin MSForms.Label lblSelection 
      Height          =   195
      Left            =   45
      TabIndex        =   19
      Top             =   60
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
   Begin MSForms.TextBox txtMajCust 
      Height          =   315
      Left            =   6825
      TabIndex        =   18
      Tag             =   "MajorCustomers"
      Top             =   2325
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtCapacity 
      Height          =   315
      Left            =   6825
      TabIndex        =   17
      Tag             =   "Capacity"
      Top             =   1935
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtRemarks2 
      Height          =   315
      Left            =   6825
      TabIndex        =   16
      Tag             =   "Remarks2"
      Top             =   1530
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtRemarks1 
      Height          =   315
      Left            =   6825
      TabIndex        =   15
      Tag             =   "Remarks1"
      Top             =   1140
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   195
      Left            =   5250
      TabIndex        =   14
      Top             =   2340
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
   Begin MSForms.Label lblCapacity 
      Height          =   195
      Left            =   5250
      TabIndex        =   13
      Top             =   1950
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
   Begin MSForms.Label lblRemarks2 
      Height          =   195
      Left            =   5250
      TabIndex        =   12
      Top             =   1560
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
   Begin MSForms.Label lblRemarks1 
      Height          =   195
      Left            =   5250
      TabIndex        =   11
      Top             =   1185
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
   Begin MSForms.TextBox txtCoreCompetencies 
      Height          =   315
      Left            =   1845
      TabIndex        =   10
      Tag             =   "CoreCompetencies"
      Top             =   2325
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtProductLine 
      Height          =   315
      Left            =   1830
      TabIndex        =   9
      Tag             =   "ProductLine"
      Top             =   1905
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtToolNo 
      Height          =   315
      Left            =   1830
      TabIndex        =   8
      Tag             =   "ToolNo"
      Top             =   1515
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtMachineNo 
      Height          =   315
      Left            =   1845
      TabIndex        =   7
      Tag             =   "MachineNo"
      Top             =   1125
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtEmpNo 
      Height          =   315
      Left            =   1845
      TabIndex        =   6
      Tag             =   "EmpNo"
      Top             =   735
      Width           =   3090
      VariousPropertyBits=   746604571
      Size            =   "5450;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblCoreCompetencies 
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   2325
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
   Begin MSForms.Label Label2 
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1935
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
   Begin MSForms.Label Label1 
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1560
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
   Begin MSForms.Label lblMachine 
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1170
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
   Begin MSForms.Label lblNoofEmployee 
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   780
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
   Begin MSForms.ComboBox cmbSupplier 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Tag             =   "AccNo"
      Top             =   330
      Width           =   4890
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "8625;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmSupplierEvaluation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileNameTag As String, strFileName As String

Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "INSERT INTO " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is VB.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is VB.CheckBox) Or (TypeOf c Is VB.OptionButton) Then
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
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is VB.TextBox) Or (TypeOf c Is VB.ComboBox) Then
            If c.Name = "cmbSupplier" Then
                RetString = RetString & "'" & c.List(c.ListIndex, 1) & "',"
            Else
                RetString = RetString & "'" & c.Text & "',"
            End If
                
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & Abs(c.Value) & ","
            ElseIf (TypeOf c Is VB.CheckBox) Then
                RetString = RetString & c.Value & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & c.ID & "',"
            ElseIf (TypeOf c Is VB.OptionButton) Then
                If c.Name = "OptMarketting" Then
                    RetString = RetString & Val(lblMarketingReputation.Tag) & ","
                ElseIf c.Name = "OptTech" Then
                    RetString = RetString & Val(lblTechnicalCapability.Tag) & ","
                End If
            End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ")"
    CreateInsertSQL = RetString
    
End Function

Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "UPDATE " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
    
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is VB.ComboBox) Then '
                If c.Name = "cmbSupplier" Then
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is VB.CheckBox) Then
                RetString = RetString & c.Tag & "=" & c.Value & ","
            ElseIf (TypeOf c Is VB.OptionButton) Then
                If c.Name = "OptMarketting" Then
                    RetString = RetString & c.Tag & "=" & Val(lblMarketingReputation.Tag) & ","
                ElseIf c.Name = "OptTech" Then
                    RetString = RetString & c.Tag & "=" & Val(lblTechnicalCapability.Tag) & ","
                End If
            End If
        End If
    Next
     
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function

Private Sub cmbSupplier_Change()
    Call cmbSupplier_Click
End Sub

Private Sub cmbSupplier_Click()
    If cmbSupplier.MatchFound Then
        Call AddEvaluation(cmbSupplier.List(cmbSupplier.ListIndex, 1))
    Else
        Call ClearAll
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If cmbSupplier.MatchFound = -False Then
        MsgBox "Please Select Supplier.", vbInformation
        Exit Sub
    ElseIf lblMarketingReputation.Tag = "-1" Then
        MsgBox "Please Select Marketing Reputation.", vbInformation
        Exit Sub
    ElseIf lblTechnicalCapability.Tag = "-1" Then
        MsgBox "Please Select Technical Capability.", vbInformation
        Exit Sub
    End If
    
    Dim lRecAffect As Long
    
    Dim SQL As String
    SQL = CreateUpdateSQL("SupplierEvaluation")
    SQL = SQL & " WHERE AccNo='" & cmbSupplier.List(cmbSupplier.ListIndex, 1) & "'"
    con.Execute SQL, lRecAffect
    
    If lRecAffect = 0 Then
        SQL = CreateInsertSQL("SupplierEvaluation")
        con.Execute SQL
    End If
    
    Dim lmyEntryID As Long
    lmyEntryID = GetSingleLongValue("EntryID", "SupplierEvaluation", " WHERE AccNo='" & cmbSupplier.List(cmbSupplier.ListIndex, 1) & "'")
    
    Call AttachPDF(lmyEntryID)
    
    MsgBox "Saved Successfully.", vbInformation
    
    
End Sub

Private Sub Form_Load()

    With cmbTimeApprove
        .AddItem "1 Year"
        .AddItem "2 Year"
        .AddItem "3 Year"
        .AddItem "Forever"
    End With
    
    Call AddToCombo(cmbSupplier, "AccTitle", "VVendersAndMakers", , , "AccNo")
    
End Sub


Public Sub ShowData(p_strAccNo As String)

    Dim rs As New ADODB.Recordset
    Dim i As Integer, c As Control, strObject As String
    With rs
        .Open "SELECT * FROM SupplierEvaluation WHERE AccNo='" & p_strAccNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" And Not (TypeOf c Is MSForms.Label) Then
                strObject = c.Name
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is VB.CheckBox Then
                    c.Value = Abs(.Fields(c.Tag))
                ElseIf TypeOf c Is VB.OptionButton Then
                    If c.Name = "OptMarketting" Then
                        OptMarketting(Val(.Fields(c.Tag))).Value = True
                        lblMarketingReputation.Tag = .Fields(c.Tag)
                    ElseIf c.Name = "OptTech" Then
                        OptTech(Val(.Fields(c.Tag))).Value = True
                        lblTechnicalCapability.Tag = .Fields(c.Tag)
                    End If
                ElseIf (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is VB.ComboBox) Then '
                    If c.Name = "cmbSupplier" Then
                        For i = 0 To cmbSupplier.ListCount - 1
                            If .Fields(c.Tag) = cmbSupplier.List(i, 1) Then
                                c.ListIndex = i
                                Exit For
                            End If
                        Next
                    Else
                        c.Text = .Fields(c.Tag)
                    End If
                End If
            End If
        Next
        
        txtFilePath = !PDFFileName & ""
        strFileNameTag = !PDFFileName & ""
        
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub AddEvaluation(p_strAccNo As String)

    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "SupplierEvaluation", " WHERE AccNo='" & p_strAccNo & "'")
    If lCount > 0 Then
        Call ShowData(p_strAccNo)
    Else
        Call ClearAll
    End If
     
End Sub

Private Sub OptMarketting_Click(Index As Integer)
    lblMarketingReputation.Tag = Index
End Sub

Private Sub OptTech_Click(Index As Integer)
    lblTechnicalCapability.Tag = Index
End Sub

Private Sub cmdBrowseBarcode_Click()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        txtFilePath = Picfile
        strFileNameTag = ""
        strFileName = CD1.FileTitle
    End If

End Sub

Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" And strFileNameTag = "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,PDFFile,PDFFileName FROM SupplierEvaluation WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
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


Private Sub ClearAll()

    Dim c As Control
    For Each c In Controls
        If TypeOf c Is MSForms.TextBox Or TypeOf c Is VB.TextBox Or TypeOf c Is MSForms.ComboBox Or TypeOf c Is VB.ComboBox Then
            If c.Name <> "cmbSupplier" Then
                c.Text = ""
            End If
        ElseIf TypeOf c Is MSForms.CheckBox Then
            c.Value = False
        ElseIf TypeOf c Is VB.CheckBox Then
            c.Value = vbUnchecked
        ElseIf TypeOf c Is VB.OptionButton Then
            c.Value = False
        End If
    Next
    
    txtFilePath = ""
    strFileName = ""
    strFileNameTag = ""
    lblTechnicalCapability.Tag = "-1"
    lblMarketingReputation.Tag = "-1"
    
End Sub
