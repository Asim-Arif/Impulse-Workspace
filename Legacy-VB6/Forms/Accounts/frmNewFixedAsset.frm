VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmNewFixedAsset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Asset"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9690
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
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   646
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Fixed Asset :"
      ClipControls    =   0   'False
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
      Height          =   7860
      Left            =   90
      TabIndex        =   3
      Top             =   15
      Width           =   9480
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   9120
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   7665
            TabIndex        =   25
            Tag             =   "AcquiredDT"
            Top             =   915
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   120455171
            CurrentDate     =   40698
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Master :"
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
            Index           =   14
            Left            =   60
            TabIndex        =   70
            Top             =   240
            Width           =   690
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   315
            Left            =   960
            TabIndex        =   69
            Tag             =   "MasterFA_RefID"
            Top             =   180
            Width           =   5385
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "9499;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   11
            Left            =   2160
            TabIndex        =   27
            Tag             =   "Description"
            Top             =   915
            Width           =   5490
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "9684;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Date Acquired"
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
            Index           =   8
            Left            =   7695
            TabIndex        =   26
            Top             =   660
            Width           =   1200
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Description"
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
            Index           =   7
            Left            =   2160
            TabIndex        =   24
            Top             =   660
            Width           =   960
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Asset ID"
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
            Index           =   4
            Left            =   90
            TabIndex        =   15
            Top             =   660
            Width           =   720
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   105
            TabIndex        =   14
            Tag             =   "AssetID"
            Top             =   915
            Width           =   2040
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "3598;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5625
         Left            =   90
         TabIndex        =   4
         Top             =   2085
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   9922
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   609
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General Info."
         TabPicture(0)   =   "frmNewFixedAsset.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmbCalibrationFrequency"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkRequiresCalibration"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Depreciation"
         TabPicture(1)   =   "frmNewFixedAsset.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "label1(10)"
         Tab(1).Control(1)=   "label1(11)"
         Tab(1).Control(2)=   "label1(15)"
         Tab(1).Control(3)=   "cmbDepMethod"
         Tab(1).Control(4)=   "label1(16)"
         Tab(1).Control(5)=   "label1(17)"
         Tab(1).Control(6)=   "label1(18)"
         Tab(1).Control(7)=   "cmbConvention"
         Tab(1).Control(8)=   "NBox(0)"
         Tab(1).Control(9)=   "label1(19)"
         Tab(1).Control(10)=   "NBox(1)"
         Tab(1).Control(11)=   "label1(20)"
         Tab(1).Control(12)=   "NBox(9)"
         Tab(1).Control(13)=   "DTDepStartMonth"
         Tab(1).Control(14)=   "Frame3"
         Tab(1).Control(15)=   "cmbDepAcc"
         Tab(1).Control(16)=   "cmbAccumAcc"
         Tab(1).ControlCount=   17
         TabCaption(2)   =   "Picture"
         TabPicture(2)   =   "frmNewFixedAsset.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "CD1"
         Tab(2).Control(1)=   "Image1"
         Tab(2).Control(2)=   "label1(32)"
         Tab(2).Control(3)=   "lblPic"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Repair Log"
         TabPicture(3)   =   "frmNewFixedAsset.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "LV"
         Tab(3).ControlCount=   1
         Begin VB.CheckBox chkRequiresCalibration 
            Caption         =   "Requires Calibration"
            Height          =   315
            Left            =   210
            TabIndex        =   74
            Tag             =   "RequiresCalibration"
            Top             =   5160
            Width           =   1935
         End
         Begin ComboList.Usercontrol1 cmbAccumAcc 
            Height          =   285
            Left            =   -74865
            TabIndex        =   38
            Tag             =   "AccumAccNo"
            Top             =   720
            Width           =   8805
            _ExtentX        =   15531
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
         Begin ComboList.Usercontrol1 cmbDepAcc 
            Height          =   285
            Left            =   -74865
            TabIndex        =   39
            Tag             =   "DepriciationAccNo"
            Top             =   1275
            Width           =   8805
            _ExtentX        =   15531
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
         Begin VB.Frame Frame3 
            Caption         =   "Depreciation Status"
            Height          =   1740
            Left            =   -74265
            TabIndex        =   53
            Top             =   3555
            Width           =   7620
            Begin MSComCtl2.DTPicker DTLastDep 
               Height          =   315
               Left            =   1815
               TabIndex        =   55
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   120455171
               CurrentDate     =   40698
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   6
               Left            =   5160
               TabIndex        =   68
               Top             =   605
               Width           =   1590
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2805;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   8
               Left            =   5160
               TabIndex        =   66
               Top             =   1335
               Width           =   1590
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2805;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "As of"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   28
               Left            =   3480
               TabIndex        =   65
               Top             =   1410
               Width           =   375
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   7
               Left            =   5160
               TabIndex        =   64
               Top             =   970
               Width           =   1590
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2805;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "LTD Depreciation"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   26
               Left            =   3510
               TabIndex        =   63
               Top             =   1050
               Width           =   1215
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "YTD Depreciation"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   25
               Left            =   3525
               TabIndex        =   62
               Top             =   660
               Width           =   1230
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   5
               Left            =   5160
               TabIndex        =   61
               Top             =   240
               Width           =   1590
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2805;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Last Depr. Run"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   24
               Left            =   3495
               TabIndex        =   60
               Top             =   285
               Width           =   1080
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Current Book Value"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   23
               Left            =   90
               TabIndex        =   59
               Top             =   1065
               Width           =   1380
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   3
               Left            =   1815
               TabIndex        =   58
               Top             =   975
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   2
               Left            =   1815
               TabIndex        =   57
               Top             =   600
               Width           =   1335
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2355;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Depreciations Done"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   22
               Left            =   120
               TabIndex        =   56
               Top             =   630
               Width           =   1395
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Last Depr. Date"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   21
               Left            =   105
               TabIndex        =   54
               Top             =   285
               Width           =   1140
            End
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4155
            Left            =   210
            TabIndex        =   5
            Top             =   915
            Width           =   8700
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   330
               Left            =   4110
               TabIndex        =   32
               Tag             =   "PlacedInServiceDT"
               Top             =   3585
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   582
               _Version        =   393216
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   120455171
               CurrentDate     =   40698
            End
            Begin MSComCtl2.DTPicker DTRetirement 
               Height          =   315
               Left            =   7020
               TabIndex        =   36
               Tag             =   "RetirementDT"
               Top             =   3570
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   120455171
               CurrentDate     =   40698
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Retirement Date"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   9
               Left            =   7065
               TabIndex        =   37
               Top             =   3360
               Width           =   1185
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Acquisition Cost"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   6
               Left            =   5790
               TabIndex        =   35
               Top             =   3345
               Width           =   1140
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   4
               Left            =   5775
               TabIndex        =   34
               Tag             =   "AcquisitionCost"
               Top             =   3570
               Width           =   1230
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2170;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Date Placed in Service"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   5
               Left            =   4155
               TabIndex        =   33
               Top             =   3375
               Width           =   1590
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Vendor"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   31
               Top             =   3375
               Width           =   510
            End
            Begin MSForms.ComboBox cmbVender 
               Height          =   315
               Left            =   105
               TabIndex        =   30
               Tag             =   "VendAccNo"
               Top             =   3585
               Width           =   3975
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "7011;556"
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbNewUsed 
               Height          =   315
               Left            =   6060
               TabIndex        =   29
               Tag             =   "Status"
               Top             =   2895
               Width           =   2520
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "4445;556"
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "New / Used"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   2
               Left            =   6060
               TabIndex        =   28
               Top             =   2670
               Width           =   825
            End
            Begin MSForms.TextBox TBox 
               Height          =   1065
               Index           =   15
               Left            =   105
               TabIndex        =   21
               Tag             =   "Notes"
               Top             =   1530
               Width           =   8475
               VariousPropertyBits=   -1467987941
               BorderStyle     =   1
               Size            =   "14949;1879"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   14
               Left            =   6060
               TabIndex        =   10
               Tag             =   "ModelNo"
               Top             =   390
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
               Caption         =   "Model No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   73
               Left            =   6060
               TabIndex        =   20
               Top             =   180
               Width           =   720
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Notes"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   19
               Left            =   150
               TabIndex        =   19
               Top             =   1305
               Width           =   420
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   9
               Left            =   105
               TabIndex        =   9
               Tag             =   "Manufacturer"
               Top             =   960
               Width           =   5925
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "10451;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Manufacturer"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   36
               Left            =   135
               TabIndex        =   18
               Top             =   750
               Width           =   975
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Department"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   27
               Left            =   150
               TabIndex        =   17
               Top             =   2670
               Width           =   855
            End
            Begin MSForms.ComboBox cmbDepartment 
               Height          =   315
               Left            =   105
               TabIndex        =   12
               Tag             =   "DeptID"
               Top             =   2895
               Width           =   5925
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "10451;556"
               MatchEntry      =   1
               ListStyle       =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Location"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   8
               Top             =   180
               Width           =   600
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   0
               Left            =   105
               TabIndex        =   7
               Tag             =   "Location"
               Top             =   390
               Width           =   5925
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "10451;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   5
               Left            =   6060
               TabIndex        =   11
               Tag             =   "SerialNo"
               Top             =   960
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
               Caption         =   "Serial No."
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   0
               Left            =   6060
               TabIndex        =   6
               Top             =   765
               Width           =   690
            End
         End
         Begin MSComCtl2.DTPicker DTDepStartMonth 
            Height          =   315
            Left            =   -74865
            TabIndex        =   45
            Tag             =   "DepriciationBeginMonth"
            Top             =   2820
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMMM"
            Format          =   120389635
            CurrentDate     =   40698
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   -73845
            Top             =   1680
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "Please Select Employee's Picture"
            Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
         End
         Begin MSComctlLib.ListView LV 
            Height          =   5055
            Left            =   -74910
            TabIndex        =   73
            Top             =   480
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   952
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Remarks"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Reading"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Sch. Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Cost"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Bill Ref."
               Object.Width           =   1764
            EndProperty
         End
         Begin MSForms.ComboBox cmbCalibrationFrequency 
            Height          =   315
            Left            =   2235
            TabIndex        =   75
            Tag             =   "CalibrationFrequency"
            Top             =   5160
            Visible         =   0   'False
            Width           =   2265
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   3975
            Left            =   -73050
            Stretch         =   -1  'True
            Top             =   1020
            Width           =   4830
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Picture"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   -73050
            TabIndex        =   72
            Top             =   810
            Width           =   495
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double Click To Add Picture."
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -71205
            TabIndex        =   71
            Top             =   2820
            Width           =   1095
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   9
            Left            =   -72330
            TabIndex        =   67
            Tag             =   "LifeInYears"
            Top             =   2820
            Width           =   2520
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "4445;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Salvage Value"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   20
            Left            =   -68535
            TabIndex        =   52
            Top             =   1830
            Width           =   1005
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   1
            Left            =   -68520
            TabIndex        =   51
            Tag             =   "SalvageValue"
            Top             =   2055
            Width           =   1230
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2170;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Adjusted Cost"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   19
            Left            =   -69765
            TabIndex        =   50
            Top             =   1830
            Width           =   1020
         End
         Begin MSForms.TextBox NBox 
            Height          =   315
            Index           =   0
            Left            =   -69780
            TabIndex        =   49
            Tag             =   "AdjustedCost"
            Top             =   2055
            Width           =   1230
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2170;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox cmbConvention 
            Height          =   315
            Left            =   -72330
            TabIndex        =   48
            Tag             =   "Convention"
            Top             =   2055
            Width           =   2520
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4445;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Convention"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   18
            Left            =   -72330
            TabIndex        =   47
            Top             =   1830
            Width           =   825
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Life in Years"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   17
            Left            =   -72300
            TabIndex        =   46
            Top             =   2580
            Width           =   870
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Depreciation Begin Month"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   16
            Left            =   -74820
            TabIndex        =   44
            Top             =   2580
            Width           =   1830
         End
         Begin MSForms.ComboBox cmbDepMethod 
            Height          =   315
            Left            =   -74865
            TabIndex        =   43
            Tag             =   "DepriciationMethodID"
            Top             =   2055
            Width           =   2520
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4445;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Depreciation Method"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   15
            Left            =   -74820
            TabIndex        =   42
            Top             =   1830
            Width           =   1485
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Accum Dep'n Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   11
            Left            =   -74820
            TabIndex        =   41
            Top             =   525
            Width           =   1545
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Depr Expense Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   10
            Left            =   -74820
            TabIndex        =   40
            Top             =   1065
            Width           =   1635
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sale Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   105
            TabIndex        =   23
            Top             =   -1035
            Width           =   930
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sale Return Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   105
            TabIndex        =   22
            Top             =   -495
            Width           =   1470
         End
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   8010
      TabIndex        =   16
      Top             =   8160
      Width           =   1575
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2778;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CheckBox chkInactive 
      Height          =   345
      Left            =   105
      TabIndex        =   2
      Tag             =   "InActive"
      Top             =   8160
      Width           =   1170
      BackColor       =   -2147483633
      ForeColor       =   8388608
      DisplayStyle    =   4
      Size            =   "2064;609"
      Value           =   "0"
      Caption         =   "Inactive"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   6405
      TabIndex        =   1
      Top             =   8160
      Width           =   1575
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2778;661"
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
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   8160
      Width           =   1575
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2778;661"
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
Attribute VB_Name = "frmNewFixedAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim lEntryIDCopyFrom As Long

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
    
    Call cmbAccumAcc.AddVals(con, "AccTitle + ' {' + AccNo + '}'", "VActiveAccounts", "AccNo", " Order By ACCTitle")
    
    Call cmbDepAcc.AddVals(con, "AccTitle + ' {' + AccNo + '}'", "VActiveAccounts", "AccNo", " Order By ACCTitle")
    
    Call AddToCombo(cmbDepartment, "Name", "Departments", "", False, "DeptID")
    Call AddToCombo(cmbVender, "AccTitle", "VVenders", "", False, "AccNo")
    Call AddToCombo(cmbDepMethod, "DepriciationMethod", "DepriciationMethods", "", False, "EntryID")
    Call AddToCombo(cmbConvention, "Convention", "FixedAssets", , , "Convention")
    
    With cmbNewUsed
        .AddItem "New"
        .List(0, 1) = "0"
        .AddItem "Used"
        .List(1, 1) = "1"
    End With
    
    Call AddToCombo(cmbCatID, "Description", "FixedAssets_Master", "", False, "EntryID")
        
    cmbCalibrationFrequency.AddItem "Monthly"
    cmbCalibrationFrequency.AddItem "Yearly"
     
End Sub

Private Sub chkRequiresCalibration_Click()

    If chkRequiresCalibration.Value = vbChecked Then
        cmbCalibrationFrequency.Visible = True
    Else
        cmbCalibrationFrequency.Visible = False
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



Private Sub Form_Load()
    
    SSTab1.Tab = 0
    
    Me.KeyPreview = True
    
    cmbAccumAcc.ListHeight = 3000
    cmbDepAcc.ListHeight = 3000
    
    Call FillCmbs
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    If Trim(TBox(3)) = "" Then
        MsgBox "Invalid Asset ID", vbInformation
        SaveItem = False
        Exit Function
    ElseIf Trim(TBox(11)) = "" Then
        MsgBox "Invalid Asset Description", vbInformation
        SaveItem = False
        Exit Function
    ElseIf cmbDepMethod.MatchFound = False Then
        MsgBox "Please Select Dept. Method.", vbInformation
        Exit Function
    ElseIf cmbDepartment.MatchFound = False Then
        MsgBox "Please Select Department.", vbInformation
        Exit Function
    ElseIf cmbConvention.Text = "" Then
        MsgBox "Please Select Convention.", vbInformation
        Exit Function
    ElseIf cmbVender.MatchFound = False Then
        MsgBox "Please Select Vendor.", vbInformation
        Exit Function
    ElseIf cmbNewUsed.MatchFound = False Then
        MsgBox "Please Select New / Used.", vbInformation
        Exit Function
    ElseIf cmbCatID.MatchFound = False Then
        MsgBox "Please Select Master.", vbInformation
        Exit Function
    ElseIf chkRequiresCalibration.Value = vbChecked Then
        If cmbCalibrationFrequency.MatchFound = False Then
            MsgBox "Please Select Calibration Frequency.", vbInformation
            Exit Function
        End If
    End If
    
    Dim SQL As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        SQL = CreateInsertSQL("FixedAssets")
    Else
        SQL = CreateUpdateSQL("FixedAssets")
        SQL = SQL & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute SQL
    
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "FixedAssets")
    Else
        lmyEntryID = lEntryID
    End If
    
    con.CommitTrans
    
    
SavePic:

    If (Image1.Picture = 0 Or Trim(Image1.Tag) = "") Then GoTo TheEnd

    Dim Rec As New ADODB.Recordset
    Dim rsPic As New ADODB.Stream
    Dim rsThumb As New ADODB.Stream
        
    Rec.Open "SELECT EntryID,AssetPic FROM FixedAssets WHERE EntryID=" & lmyEntryID, con, adOpenStatic, adLockOptimistic
    If Image1.Picture <> 0 And Trim(Image1.Tag) <> "" Then
        rsPic.Open
        rsPic.Type = adTypeBinary
        rsPic.LoadFromFile Image1.Tag
        Rec.Fields(1) = rsPic.Read
        rsPic.Close
        Rec.Update
    End If

TheEnd:
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
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox Or TypeOf c Is VB.CheckBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Then
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
                If c.Name = "cmbConvention" Then
                    RetString = RetString & "'" & c.Text & "" & "',"
                ElseIf c.Name = "cmbCalibrationFrequency" Then
                    RetString = RetString & c.ListIndex & ","
                Else
                    RetString = RetString & "'" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Then
                If c.Name = "NBox" Then
                    RetString = RetString & Val(c.Text & "") & ","
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & Abs(c.Value) & ","
            ElseIf (TypeOf c Is VB.CheckBox) Then
                RetString = RetString & c.Value & ","
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
                If c.Name = "cmbConvention" Then
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                ElseIf c.Name = "cmbCalibrationFrequency" Then
                    RetString = RetString & c.Tag & "=" & c.ListIndex & ","
                Else
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Then   '
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
            ElseIf (TypeOf c Is VB.CheckBox) Then
                RetString = RetString & c.Tag & "=" & c.Value & ","
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


Public Sub CopyItem(p_lEntryID As Long)
       
    lEntryID = 0
    
    lEntryIDCopyFrom = p_lEntryID
    'CopyItemCode = strItemCode
    Call ShowData(lEntryIDCopyFrom, True)
    Call GetNewFixedAssetID
    
    Me.Show 1
    
End Sub

Private Sub ShowData(p_lEntryID As Long, Optional bCopy As Boolean = False)

    On Error GoTo err
    Dim ITM As ListItem
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
        .Open "SELECT * FROM VFixedAssets WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
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
                ElseIf TypeOf c Is VB.CheckBox Then
                    c.Value = Abs(!RequiresCalibration)
                ElseIf c.Name = "cmbCalibrationFrequency" Then
                    c.ListIndex = Val(!CalibrationFrequency & "")
                Else
                    If c.Name = "cmbCatID" Then
                        c.Text = .Fields("FixedAssets_MasterDescription")
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        
        
        'TBox(8) = ![CatID] & "-"
        TBox(3) = ![AssetID] & ""
        
        cmbDepartment = !Name & ""
        cmbNewUsed.ListIndex = Val(!Status & "")
        cmbDepMethod.Text = !Depriciationmethod & ""
        cmbVender.Text = !VenderAccTitle & ""
        
        If Not IsNull(![AssetPic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "AssetPic"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
        
        .Close
        
        .Open "SELECT * FROM FixedAssets_RepairLog WHERE AssetID='" & TBox(3) & "' ORDER BY EntryID DESC", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !Remarks & ""
            ITM.ListSubItems.add , , Val(!MeterReading & "")
            ITM.ListSubItems.add , , Format(!ScheduleDT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Val(!CostOfRepair & "")
            ITM.ListSubItems.add , , !BillReference & ""
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
    
    Call GetNewFixedAssetID
    Me.Show 1
    
End Sub

Private Sub GetNewFixedAssetID()
    'Get Asset ID...
    TBox(3) = GetSingleLongValue("MAX(CAST(AssetID AS INT))", "FixedAssets") + 1
    TBox(3) = Format(TBox(3), "0000")
End Sub

Private Sub Image1_DblClick()
    Call lblPic_DblClick
End Sub

Private Sub lblPic_DblClick()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        Image1.Picture = LoadPicture(Picfile)
        Image1.Tag = Picfile
        lblPic.Visible = False
    End If
    
End Sub
