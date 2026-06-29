VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form NewEmployee1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Info..."
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
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
   ScaleHeight     =   5190
   ScaleWidth      =   9120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Employee Information :"
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
      Height          =   4515
      Left            =   105
      TabIndex        =   31
      Top             =   15
      Width           =   8940
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   4080
         TabIndex        =   36
         Top             =   780
         Visible         =   0   'False
         Width           =   4320
      End
      Begin MSComCtl2.DTPicker JoinDate 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   3450
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   22740995
         CurrentDate     =   37383
      End
      Begin MSComCtl2.DTPicker RetireDate 
         Height          =   315
         Left            =   1740
         TabIndex        =   20
         Top             =   3450
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   22740995
         CurrentDate     =   37383
      End
      Begin MSComCtl2.DTPicker DOB 
         Height          =   315
         Left            =   5070
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   22740995
         CurrentDate     =   37383
      End
      Begin VB.Frame fCat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   390
         TabIndex        =   33
         Top             =   180
         Width           =   8070
         Begin VB.Label Label1 
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
            Left            =   90
            TabIndex        =   35
            Top             =   270
            Width           =   1365
         End
         Begin MSForms.ComboBox cmbDeptID 
            Height          =   300
            Left            =   1500
            TabIndex        =   0
            Top             =   210
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
         Begin MSForms.ComboBox cmbDept 
            Height          =   300
            Left            =   4860
            TabIndex        =   1
            Top             =   210
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
         Begin VB.Label Label1 
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
            Left            =   3165
            TabIndex        =   34
            Top             =   255
            Width           =   1650
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P Profit"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   24
         Left            =   4155
         TabIndex        =   59
         Top             =   3255
         Width           =   525
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   17
         Left            =   4140
         TabIndex        =   22
         Top             =   3450
         Width           =   885
         VariousPropertyBits=   679495707
         Size            =   "1561;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkOldAge 
         Height          =   345
         Left            =   3090
         TabIndex        =   25
         Top             =   4080
         Width           =   1980
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3492;609"
         Value           =   "0"
         Caption         =   "Old Age Benifit Holder"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkSocial 
         Height          =   345
         Left            =   1410
         TabIndex        =   24
         Top             =   4080
         Width           =   1410
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2487;609"
         Value           =   "0"
         Caption         =   "Social Security"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkTaxable 
         Height          =   345
         Left            =   240
         TabIndex        =   23
         Top             =   4080
         Width           =   915
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1614;609"
         Value           =   "0"
         Caption         =   "Taxable"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   915
         Index           =   16
         Left            =   5100
         TabIndex        =   26
         Top             =   3450
         Width           =   3705
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "6535;1614"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   23
         Left            =   5115
         TabIndex        =   57
         Top             =   3240
         Width           =   615
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   15
         Left            =   3210
         TabIndex        =   21
         Top             =   3450
         Width           =   885
         VariousPropertyBits=   679495707
         Size            =   "1561;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Salary"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   22
         Left            =   3255
         TabIndex        =   56
         Top             =   3240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Retirement Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   21
         Left            =   1770
         TabIndex        =   55
         Top             =   3210
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Join Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   20
         Left            =   270
         TabIndex        =   54
         Top             =   3210
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reason"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   19
         Left            =   7035
         TabIndex        =   53
         Top             =   2640
         Width           =   540
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   14
         Left            =   6990
         TabIndex        =   18
         Top             =   2850
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   5085
         TabIndex        =   52
         Top             =   2640
         Width           =   1185
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   13
         Left            =   5070
         TabIndex        =   17
         Top             =   2850
         Width           =   1875
         VariousPropertyBits=   679495707
         Size            =   "3307;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Firm"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   3255
         TabIndex        =   51
         Top             =   2640
         Width           =   645
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   12
         Left            =   3210
         TabIndex        =   16
         Top             =   2850
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Identity Mark"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   7035
         TabIndex        =   50
         Top             =   1980
         Width           =   960
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   11
         Left            =   6990
         TabIndex        =   15
         Top             =   2190
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reference"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   5085
         TabIndex        =   49
         Top             =   1980
         Width           =   750
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   10
         Left            =   5070
         TabIndex        =   14
         Top             =   2190
         Width           =   1875
         VariousPropertyBits=   679495707
         Size            =   "3307;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   3225
         TabIndex        =   48
         Top             =   1980
         Width           =   840
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   9
         Left            =   3210
         TabIndex        =   13
         Top             =   2190
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   47
         Top             =   1980
         Width           =   585
      End
      Begin MSForms.TextBox TBox 
         Height          =   975
         Index           =   8
         Left            =   225
         TabIndex        =   12
         Top             =   2190
         Width           =   2925
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "5159;1720"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Phone 2"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   7035
         TabIndex        =   46
         Top             =   1440
         Width           =   585
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   7
         Left            =   6990
         TabIndex        =   11
         Top             =   1650
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Phone 1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   5085
         TabIndex        =   45
         Top             =   1440
         Width           =   585
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   6
         Left            =   5070
         TabIndex        =   10
         Top             =   1650
         Width           =   1875
         VariousPropertyBits=   679495707
         Size            =   "3307;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NTN"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   3225
         TabIndex        =   44
         Top             =   1440
         Width           =   300
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   5
         Left            =   3210
         TabIndex        =   9
         Top             =   1650
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NIC #"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   1365
         TabIndex        =   43
         Top             =   1440
         Width           =   435
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   4
         Left            =   1350
         TabIndex        =   8
         Top             =   1650
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbSex 
         Height          =   315
         Left            =   225
         TabIndex        =   7
         Top             =   1650
         Width           =   1065
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1879;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   42
         Top             =   1440
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Birth"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   5100
         TabIndex        =   41
         Top             =   840
         Width           =   945
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   3
         Left            =   8220
         TabIndex        =   6
         Top             =   1080
         Width           =   585
         VariousPropertyBits=   679495707
         Size            =   "1032;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   8235
         TabIndex        =   40
         Top             =   870
         Width           =   285
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   2
         Left            =   6570
         TabIndex        =   5
         Top             =   1080
         Width           =   1605
         VariousPropertyBits=   679495707
         Size            =   "2831;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Religion"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   6585
         TabIndex        =   39
         Top             =   870
         Width           =   555
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   1
         Left            =   3210
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Father Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   3225
         TabIndex        =   38
         Top             =   870
         Width           =   930
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   0
         Left            =   1350
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
         VariousPropertyBits=   679495707
         Size            =   "3201;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1365
         TabIndex        =   37
         Top             =   870
         Width           =   405
      End
      Begin VB.Label lblEmpID 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   225
         TabIndex        =   30
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Employee ID"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   32
         Top             =   870
         Width           =   900
      End
   End
   Begin MSForms.CheckBox chkSpecial 
      Height          =   345
      Left            =   330
      TabIndex        =   58
      Top             =   4650
      Visible         =   0   'False
      Width           =   1305
      VariousPropertyBits=   1015023643
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2302;609"
      Value           =   "0"
      Caption         =   "Special Case"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   405
      Left            =   7395
      TabIndex        =   29
      Top             =   4680
      Width           =   1620
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "NewEmployee1.frx":0000
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   405
      Left            =   4035
      TabIndex        =   27
      Top             =   4680
      Width           =   1620
      Caption         =   "Save & New"
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "NewEmployee1.frx":0112
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   405
      Left            =   5715
      TabIndex        =   28
      Top             =   4680
      Width           =   1620
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "NewEmployee1.frx":0224
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "NewEmployee1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean
Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Dim dontShowList As Boolean
Public SqlSettings As String

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
    Call AddToCombo(cmbDeptID, "deptID", "Departments", "Order By DeptID")
    Call AddToCombo(cmbDept, "Name", "Departments", " Order By DeptID")
    'Call AddToCombo(cmbSaleAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '31%' AND Not Parent Order By ACCTitle")
    'Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '41%' AND Not Parent Order By ACCTitle")
    'Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    
End Sub
Private Sub chkSpecial_Click()
    If Not Loaded Then Exit Sub
    If lblEmpID = "" Then chkSpecial.value = False
    If chkSpecial Then
        If SqlSettings = "" Then
            
            Load frmSettings
            frmSettings.lblEmpID = lblEmpID
            frmSettings.IsNew = True
            
            Dim rs As New ADODB.Recordset
            rs.Open "Select ID from Generalsettings Where ID='" & lblEmpID & "'", Con, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount > 0 Then
                frmSettings.Add = False
                Call frmSettings.Editrec
            Else
                frmSettings.Add = True
            End If
            rs.Close
            Set rs = Nothing
            
            frmSettings.Show 1
        Else
            frmSettings.Show 1
        End If
    End If
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
    If dontShowList Then dontShowList = False Else ShowList List1, cmbDept, cmbDept, fCat.Left, fCat.Top
End Sub

Private Sub cmbdept_LostFocus()
    List1.Visible = False
End Sub


Private Sub cmbdeptID_Validate(Cancel As Boolean)
    If Not cmbDeptID.MatchFound Then cmbDeptID.ListIndex = cmbDept.ListIndex
End Sub
Private Sub cmbdept_Change()
    If cmbDept.MatchFound Then Call cmbdept_Click 'Else If cmbdept <> "" Then cmbdept.DropDown
End Sub
Private Sub cmbdept_Click()
    If cmbDeptID.ListIndex <> cmbDept.ListIndex Then cmbDeptID.ListIndex = cmbDept.ListIndex
End Sub
Private Sub cmbdept_Validate(Cancel As Boolean)
   If Not cmbDept.MatchFound Then cmbDept.ListIndex = cmbDeptID.ListIndex
End Sub
Private Sub cmbsaleAcc_Validate(Cancel As Boolean)
    If Not cmbSaleAcc.MatchFound Then cmbSaleAcc = ""
End Sub

Private Sub cmbPchAcc_Validate(Cancel As Boolean)
    If Not cmbPchAcc.MatchFound Then cmbPchAcc = ""
End Sub

Private Sub cmbDeptID_Change()
    If cmbDeptID.MatchFound Then Call cmbdeptid_Click Else If Not cmbDeptID = "" Then cmbDeptID.DropDown
End Sub
Private Sub cmbdeptid_Click()
    If cmbDept.ListIndex <> cmbDeptID.ListIndex Then cmbDept.ListIndex = cmbDeptID.ListIndex
    If Add Then lblEmpID = GetNextEmpID(cmbDeptID)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    If SqlSettings <> "" Then
        Unload frmSettings
    End If
End Sub

Private Sub cmdNext_Click()
    If saveEmp Then
        Unload Me
        Me.SqlSettings = ""
        Me.Add = True
        Me.Edit = False
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If saveEmp Then
        Unload Me
        If SqlSettings <> "" Then Unload frmSettings
    End If
End Sub
'Private Sub Form_Activate()
'   If Edit And Not Loaded Then Call showdata
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   Me.KeyPreview = True
'    DTPicker1.value = Date
    Call FillCmbs
    SqlSettings = ""
End Sub

Private Function saveEmp() As Boolean

On Error GoTo err

Dim OpBalance As Double
Dim rec As New ADODB.Recordset




If Not cmbDeptID.MatchFound Then
    MsgBox "Invalid Department Name.", vbInformation
    saveEmp = False
    Exit Function
ElseIf Trim(TBox(0)) = "" Then
    MsgBox "Invalid Employee Name.", vbInformation
    saveEmp = False
    Exit Function
'ElseIf Not cmbPchAcc.MatchFound Then
'    MsgBox "No Purchase Account Selected", vbInformation
'    saveEmp = False
'    Exit Function
'ElseIf Not cmbSaleAcc.MatchFound Then
'    MsgBox "No Sales Account Selected", vbInformation
'    saveEmp = False
'    Exit Function
'ElseIf Val(TBox(4)) = 0 Then
'    MsgBox "Invalid Purchase Price", vbInformation
'    saveEmp = False
'    Exit Function
'ElseIf Val(TBox(5)) = 0 Then
'    MsgBox "Invalid Sale Price", vbInformation
'    saveEmp = False
'    Exit Function
'ElseIf Not Val(TBox(4)) >= Val(TBox(5)) = 0 Then
'    If MsgBox("Sale Price Should Be More Than Purchase Price." & vbNewLine & "Proceed Anyway ?", vbInformation + vbYesNo) = vbNo Then
'        saveEmp = False
'        Exit Function
'    End If
End If


'pDim SaleAcc As String, PchAcc As String

'SaleAcc = getAccountNo(cmbSaleAcc)
'PchAcc = getAccountNo(cmbPchAcc)
Dim OldAge As Boolean, Social As Boolean, TaxPayee As Boolean
Dim Special As String
'If chkOldAge.value = vbChecked Then OldAge = True Else OldAge = False
'If chkSocial.value = vbChecked Then Social = True Else Social = False
'If chkTaxable.value = vbChecked Then TaxPayee = True Else TaxPayee = False
OldAge = chkOldAge.value
Social = chkSocial.value
TaxPayee = chkTaxable.value
If chkSpecial.value Then Special = lblEmpID Else Special = "GENERAL"
Con.BeginTrans
If Special <> "GENERAL" Then Con.Execute SqlSettings
If Add Then
    Con.Execute "Insert into Employees(DeptID,EmpID,Name,fname," & _
     "dob,religion,age,sex,address,NIC,NTN,Phone1,Phone2,designation" & _
     ",refName,IdentityMark,lastEmpFirm,lastEmpDesig,lastEmpReason" & _
     ",JoinDate,Retirement,Salary,PProfit,StartingSalary,oldage,Social," & _
     "taxpayee,remarks,GInfo) Values('" & cmbDeptID & "','" & lblEmpID & _
     "','" & TBox(0) & "','" & TBox(1) & "',#" & DOB & "#,'" & TBox(2) & _
     "'," & Val(TBox(3)) & ",'" & cmbSex & "','" & TBox(8) & "','" & TBox(4) & _
     "','" & TBox(5) & "','" & TBox(6) & "','" & TBox(7) & "','" & _
     TBox(9) & "','" & TBox(10) & "','" & TBox(11) & "','" & TBox(12) & _
     "','" & TBox(13) & "','" & TBox(14) & "',#" & JoinDate & "#,#" & _
     RetireDate & "#," & TBox(15) & "," & TBox(17) & "," & _
     Val(TBox(15)) + Val(TBox(17)) & "," & OldAge & "," & Social & "," & _
     TaxPayee & ",'" & TBox(16) & "','" & Special & "')"
Else
    Con.Execute "Update Employees Set Name='" & TBox(0) & "',fName='" & _
     TBox(1) & "',DOB=#" & DOB & "#,Religion='" & TBox(2) & "',age=" & _
     TBox(3) & ",Sex='" & cmbSex & "',Address='" & TBox(8) & "',NIC='" & _
     TBox(4) & "',NTN='" & TBox(5) & "',Phone1='" & TBox(6) & "',Phone2='" & _
     TBox(7) & "',designation='" & TBox(9) & "',RefName='" & TBox(10) & _
     "',IdentityMark='" & TBox(11) & "',lastEmpFirm='" & TBox(12) & _
     "',lastEmpDesig='" & TBox(13) & "',lastEmpReason='" & TBox(14) & _
     "',JoinDate=#" & JoinDate & "#,Retirement=#" & RetireDate & "#," & _
     "StartingSalary=" & Val(TBox(15)) + Val(TBox(17)) & ",PProfit=" & TBox(17) & "," & _
     "Salary=" & TBox(15) & ",OldAge=" & OldAge & ",Social=" & _
     Social & ",TaxPayee=" & TaxPayee & ",Remarks='" & TBox(16) & _
     "',GInfo='" & Special & "' Where EmpID='" & Me.Tag & "'"
End If
Con.CommitTrans
'If Special = "GENERAL" Then Con.Execute "delete from GeneralSettings Where ID='" & lblEmpID & "'"
    
 saveEmp = True
Exit Function
err:
    MsgBox err.Description
End Function


Private Sub List1_Click()

    If UCase(Screen.ActiveControl.Name) <> "cmbdept" Then
        cmbDept = List1
        List1.Visible = False
        cmbDept.SetFocus
    End If

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
 NextVal = GetMax("val(Right(ItemID,2))", "Items", "Where CatID ='" & CatID & "'")
 GetNextItemID = CatID & "-" & Format(Val(NextVal) + 1, "000")
End Function

Private Sub TBox_Validate(Index As Integer, Cancel As Boolean)
    'If Index = 4 Or Index = 5 Then
    '    TBox(Index) = Format(Val(TBox(Index)), "#0.00")
    'ElseIf Index = 2 Or Index = 3 Then
    '    TBox(Index) = Format(Val(TBox(Index)), "#0")
    'End If
End Sub


Public Sub Editrec()
    Loaded = False
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From Employees Where EmpID='" & Me.Tag & "'", Con, adOpenForwardOnly, adLockReadOnly
        lblEmpID = ![EmpID] & ""
        cmbDeptID = ![DeptID] & ""
        TBox(0) = ![Name] & ""
        TBox(1) = ![FName] & ""
        DOB = ![DOB] & ""
        TBox(2) = ![Religion] & ""
        TBox(3) = ![age] & ""
        cmbSex = ![Sex] & ""
        TBox(4) = ![NIC] & ""
        TBox(5) = ![NTN] & ""
        TBox(6) = ![Phone1] & ""
        TBox(7) = ![Phone2] & ""
        TBox(8) = ![Address] & ""
        TBox(9) = ![Designation] & ""
        TBox(10) = ![RefName] & ""
        TBox(11) = ![IdentityMark] & ""
        TBox(12) = ![lastEmpFirm] & ""
        TBox(13) = ![LastEmpDesig] & ""
        TBox(14) = ![LastEmpReason] & ""
        TBox(15) = ![Salary] & ""
        TBox(17) = ![PProfit] & ""
        TBox(16) = ![Remarks] & ""
        JoinDate = ![JoinDate] & ""
        RetireDate = ![Retirement] & ""
        chkOldAge.value = ![OldAge] & ""
        chkSocial.value = ![Social] & ""
        chkTaxable.value = ![TaxPayee] & ""
        
        If ![GInfo] = "GENERAL" Then
            chkSpecial.value = False
        Else
            chkSpecial.value = True
        End If
        .Close
    End With

    Set rs = Nothing

    fCat.Enabled = False
    Loaded = True
    
    'TBox(2).Enabled = False
    'Loaded = True
    
End Sub
