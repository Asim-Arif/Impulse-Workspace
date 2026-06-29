VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditItem_VC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMR (Device Master Record)"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8430
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
   ScaleHeight     =   8220
   ScaleWidth      =   8430
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   7050
      Left            =   105
      TabIndex        =   13
      Top             =   570
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   12435
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Device Info."
      TabPicture(0)   =   "frmEditItem_VC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Other Info."
      TabPicture(1)   =   "frmEditItem_VC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Barcode"
      TabPicture(2)   =   "frmEditItem_VC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdBrowsePDF"
      Tab(2).Control(1)=   "txtPDFPath"
      Tab(2).Control(2)=   "txtProgramPath"
      Tab(2).Control(3)=   "cmdBrowseBarcode(0)"
      Tab(2).Control(4)=   "txtFilePath"
      Tab(2).Control(5)=   "cmdBrowseBarcode(1)"
      Tab(2).Control(6)=   "cmdOpenBarcode"
      Tab(2).Control(7)=   "CDBarcode"
      Tab(2).Control(8)=   "CD2"
      Tab(2).Control(9)=   "Label2(2)"
      Tab(2).Control(10)=   "Label2(0)"
      Tab(2).Control(11)=   "Label2(1)"
      Tab(2).ControlCount=   12
      Begin VB.CommandButton cmdBrowsePDF 
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
         Left            =   -67335
         TabIndex        =   94
         Top             =   4965
         Width           =   405
      End
      Begin VB.TextBox txtPDFPath 
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
         Left            =   -74925
         TabIndex        =   93
         Top             =   4980
         Width           =   7560
      End
      Begin VB.TextBox txtProgramPath 
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
         Left            =   -74115
         TabIndex        =   42
         Top             =   1095
         Width           =   6015
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
         Index           =   0
         Left            =   -68070
         TabIndex        =   41
         Top             =   1095
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
         Height          =   435
         Left            =   -74115
         TabIndex        =   40
         Top             =   1995
         Width           =   6015
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
         Index           =   1
         Left            =   -68070
         TabIndex        =   39
         Top             =   1980
         Width           =   405
      End
      Begin VB.CommandButton cmdOpenBarcode 
         Caption         =   "Open Barcode"
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
         Left            =   -74115
         TabIndex        =   38
         Top             =   2565
         Width           =   1845
      End
      Begin VB.Frame Frame2 
         Height          =   4305
         Left            =   -74100
         TabIndex        =   33
         Top             =   1095
         Width           =   6315
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Finishing"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   35
            Left            =   2160
            TabIndex        =   92
            Top             =   225
            Width           =   615
         End
         Begin MSForms.ComboBox cmbFinQuality 
            Height          =   315
            Left            =   2145
            TabIndex        =   91
            Tag             =   "FinQuality"
            Top             =   435
            Width           =   2415
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4260;556"
            ListWidth       =   3527
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
            Height          =   795
            Index           =   11
            Left            =   135
            TabIndex        =   7
            Tag             =   "PackingInstructions"
            Top             =   3195
            Width           =   6030
            VariousPropertyBits=   -1467987941
            BorderStyle     =   1
            Size            =   "10636;1402"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Packing Instructions"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   165
            TabIndex        =   37
            Top             =   2970
            Width           =   1440
         End
         Begin MSForms.TextBox TBox 
            Height          =   795
            Index           =   10
            Left            =   135
            TabIndex        =   6
            Tag             =   "StampInstructions"
            Top             =   2130
            Width           =   6030
            VariousPropertyBits=   -1467987941
            BorderStyle     =   1
            Size            =   "10636;1402"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Stamp Instructions"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   165
            TabIndex        =   36
            Top             =   1905
            Width           =   1350
         End
         Begin MSForms.TextBox TBox 
            Height          =   795
            Index           =   9
            Left            =   135
            TabIndex        =   5
            Tag             =   "SpecialInstructions"
            Top             =   1035
            Width           =   6030
            VariousPropertyBits=   -1467987941
            BorderStyle     =   1
            Size            =   "10636;1402"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Special Instructions"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   9
            Left            =   165
            TabIndex        =   35
            Top             =   810
            Width           =   1395
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   8
            Left            =   135
            TabIndex        =   4
            Tag             =   "PageNo"
            Top             =   435
            Width           =   1980
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "3492;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Customer Catalog Page #"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   7
            Left            =   165
            TabIndex        =   34
            Top             =   210
            Width           =   1860
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   6465
         Left            =   150
         TabIndex        =   14
         Top             =   495
         Width           =   7890
         Begin VB.Frame Frame3 
            Caption         =   "Revisions :"
            ForeColor       =   &H00800000&
            Height          =   3660
            Left            =   4620
            TabIndex        =   16
            Top             =   2505
            Width           =   3180
            Begin VB.ListBox lstRev 
               Appearance      =   0  'Flat
               ForeColor       =   &H00400000&
               Height          =   3150
               Left            =   75
               TabIndex        =   17
               Top             =   375
               Width           =   3045
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Pricing :"
            ForeColor       =   &H00800000&
            Height          =   3660
            Left            =   150
            TabIndex        =   15
            Top             =   2505
            Width           =   4425
            Begin VB.Frame Frame7 
               Caption         =   "Top :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1350
               Left            =   90
               TabIndex        =   71
               Top             =   1605
               Width           =   4230
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FOB Karachi"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   33
                  Left            =   120
                  TabIndex        =   87
                  Top             =   225
                  Width           =   870
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   20
                  Left            =   1095
                  TabIndex        =   86
                  Tag             =   "ExWorksTop"
                  Top             =   420
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   19
                  Left            =   2115
                  TabIndex        =   85
                  Tag             =   "CnISeaTop"
                  Top             =   420
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FOB Sialkot"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   32
                  Left            =   1095
                  TabIndex        =   84
                  Top             =   225
                  Width           =   810
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C && I (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   31
                  Left            =   2115
                  TabIndex        =   83
                  Top             =   225
                  Width           =   795
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   18
                  Left            =   3135
                  TabIndex        =   82
                  Tag             =   "CnIAirTop"
                  Top             =   420
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C && I (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   30
                  Left            =   3135
                  TabIndex        =   81
                  Top             =   210
                  Width           =   720
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "CIF (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   29
                  Left            =   1095
                  TabIndex        =   80
                  Top             =   735
                  Width           =   615
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "CIF (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   28
                  Left            =   120
                  TabIndex        =   79
                  Top             =   735
                  Width           =   690
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C & F (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   27
                  Left            =   3135
                  TabIndex        =   78
                  Top             =   735
                  UseMnemonic     =   0   'False
                  Width           =   750
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C & F (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   26
                  Left            =   2115
                  TabIndex        =   77
                  Top             =   735
                  UseMnemonic     =   0   'False
                  Width           =   825
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   17
                  Left            =   75
                  TabIndex        =   76
                  Tag             =   "FOBTop"
                  Top             =   420
                  Width           =   1000
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   16
                  Left            =   1095
                  TabIndex        =   75
                  Tag             =   "CIFAirTop"
                  Top             =   930
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   15
                  Left            =   75
                  TabIndex        =   74
                  Tag             =   "CIFSeaTop"
                  Top             =   930
                  Width           =   1000
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   5
                  Left            =   3135
                  TabIndex        =   73
                  Tag             =   "CnFAirTop"
                  Top             =   930
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   4
                  Left            =   2115
                  TabIndex        =   72
                  Tag             =   "CnFSeaTop"
                  Top             =   930
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Regular :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1395
               Left            =   90
               TabIndex        =   54
               Top             =   195
               Width           =   4230
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   10
                  Left            =   2115
                  TabIndex        =   70
                  Tag             =   "CnFSea"
                  Top             =   930
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   11
                  Left            =   3135
                  TabIndex        =   69
                  Tag             =   "CnFAir"
                  Top             =   930
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
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
                  Left            =   75
                  TabIndex        =   68
                  Tag             =   "CIFSea"
                  Top             =   930
                  Width           =   1000
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   9
                  Left            =   1095
                  TabIndex        =   67
                  Tag             =   "CIFAir"
                  Top             =   930
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   0
                  Left            =   75
                  TabIndex        =   66
                  Tag             =   "FOB"
                  Top             =   420
                  Width           =   1000
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C & F (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   22
                  Left            =   2115
                  TabIndex        =   65
                  Top             =   735
                  UseMnemonic     =   0   'False
                  Width           =   825
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C & F (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   23
                  Left            =   3135
                  TabIndex        =   64
                  Top             =   735
                  UseMnemonic     =   0   'False
                  Width           =   750
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "CIF (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   20
                  Left            =   120
                  TabIndex        =   63
                  Top             =   735
                  Width           =   690
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "CIF (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   21
                  Left            =   1125
                  TabIndex        =   62
                  Top             =   735
                  Width           =   615
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C && I (Air)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   6
                  Left            =   3135
                  TabIndex        =   61
                  Top             =   210
                  Width           =   720
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   3
                  Left            =   3135
                  TabIndex        =   60
                  Tag             =   "CnIAir"
                  Top             =   420
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "C && I (Sea)"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   5
                  Left            =   2115
                  TabIndex        =   59
                  Top             =   225
                  Width           =   795
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FOB Sialkot"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   3
                  Left            =   1125
                  TabIndex        =   58
                  Top             =   225
                  Width           =   810
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   2
                  Left            =   2115
                  TabIndex        =   57
                  Tag             =   "CnISea"
                  Top             =   420
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox NBox 
                  Height          =   315
                  Index           =   1
                  Left            =   1095
                  TabIndex        =   56
                  Tag             =   "ExWorks"
                  Top             =   420
                  Width           =   1005
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1764;556"
                  SpecialEffect   =   0
                  FontName        =   "Tahoma"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin VB.Label label1 
                  AutoSize        =   -1  'True
                  Caption         =   "FOB Karachi"
                  ForeColor       =   &H00400000&
                  Height          =   195
                  Index           =   17
                  Left            =   120
                  TabIndex        =   55
                  Top             =   225
                  Width           =   870
               End
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   22
               Left            =   1665
               TabIndex        =   99
               Tag             =   "First_Inspection_Rate"
               Top             =   3195
               Width           =   1545
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2725;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "First Inspection Rate"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   37
               Left            =   1695
               TabIndex        =   98
               Top             =   3000
               Width           =   1500
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   21
               Left            =   105
               TabIndex        =   97
               Tag             =   "Temper_Rate"
               Top             =   3195
               Width           =   1545
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2725;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Temper Rate"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   36
               Left            =   135
               TabIndex        =   96
               Top             =   3000
               Width           =   930
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Custom Prices :"
            ForeColor       =   &H00800000&
            Height          =   930
            Left            =   105
            TabIndex        =   18
            Top             =   2955
            Visible         =   0   'False
            Width           =   6030
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   6
               Left            =   135
               TabIndex        =   28
               Tag             =   "CustomFOB"
               Top             =   465
               Width           =   1125
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1984;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   7
               Left            =   1290
               TabIndex        =   27
               Tag             =   "CustomCIFSea"
               Top             =   465
               Width           =   1125
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1984;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   12
               Left            =   2445
               TabIndex        =   26
               Tag             =   "CustomCIFAir"
               Top             =   465
               Width           =   1125
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1984;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   13
               Left            =   3600
               TabIndex        =   25
               Tag             =   "CustomCnFSea"
               Top             =   465
               Width           =   1125
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1984;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   14
               Left            =   4755
               TabIndex        =   24
               Tag             =   "CustomCnFAir"
               Top             =   465
               Width           =   1125
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1984;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "F.O.B"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   14
               Left            =   150
               TabIndex        =   23
               Top             =   270
               Width           =   420
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "CIF (Sea)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   15
               Left            =   1350
               TabIndex        =   22
               Top             =   270
               Width           =   690
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "CIF (Air)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   16
               Left            =   2505
               TabIndex        =   21
               Top             =   255
               Width           =   615
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "C & F (Sea)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   18
               Left            =   3630
               TabIndex        =   20
               Top             =   270
               UseMnemonic     =   0   'False
               Width           =   825
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "C & F (Air)"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   19
               Left            =   4815
               TabIndex        =   19
               Top             =   270
               UseMnemonic     =   0   'False
               Width           =   750
            End
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   13
            Left            =   5235
            TabIndex        =   49
            Tag             =   "OCR2"
            Top             =   2175
            Width           =   2535
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "4471;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "OCR 2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   34
            Left            =   5235
            TabIndex        =   90
            Top             =   1965
            Width           =   465
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "OCR 1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   80
            Left            =   2685
            TabIndex        =   89
            Top             =   1965
            Width           =   465
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   22
            Left            =   2685
            TabIndex        =   48
            Tag             =   "OCR"
            Top             =   2175
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
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   21
            Left            =   105
            TabIndex        =   47
            Tag             =   "ILO"
            Top             =   2175
            Width           =   2550
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "4498;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Imp. License #"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   79
            Left            =   150
            TabIndex        =   88
            Top             =   1965
            Width           =   1065
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Packing Mode"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   11
            Left            =   150
            TabIndex        =   53
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   975
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   105
            TabIndex        =   52
            Tag             =   "PackingMode"
            Top             =   1635
            Width           =   1995
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "3519;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Device Description (Customer)"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   51
            Top             =   855
            Width           =   2175
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   105
            TabIndex        =   50
            Tag             =   "Description"
            Top             =   1065
            Width           =   7680
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "13547;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   12
            Left            =   2130
            TabIndex        =   46
            Tag             =   "ItemColor"
            Top             =   1635
            Width           =   5670
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "10001;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Color"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   25
            Left            =   2130
            TabIndex        =   45
            Top             =   1440
            Width           =   375
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   105
            TabIndex        =   1
            TabStop         =   0   'False
            Tag             =   "ItemID"
            Top             =   465
            Width           =   1995
            VariousPropertyBits=   679495711
            MaxLength       =   9
            BorderStyle     =   1
            Size            =   "3519;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Company Code"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin MSForms.ComboBox cmbUnit 
            Height          =   315
            Left            =   4155
            TabIndex        =   2
            Tag             =   "Unit"
            Top             =   465
            Width           =   1200
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2117;556"
            ListWidth       =   3527
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
            Caption         =   "Unit"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   24
            Left            =   4155
            TabIndex        =   31
            Top             =   240
            Width           =   285
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   1
            Left            =   2130
            TabIndex        =   0
            Tag             =   "ItemID_Vendor"
            Top             =   465
            Width           =   1995
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "3519;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Customer Ref. Code"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   8
            Left            =   2130
            TabIndex        =   30
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Barcode #"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   2
            Left            =   5385
            TabIndex        =   29
            Top             =   240
            Width           =   750
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   7
            Left            =   5385
            TabIndex        =   3
            Tag             =   "BarCodeNo"
            Top             =   465
            Width           =   2400
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "4233;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSComDlg.CommonDialog CDBarcode 
         Left            =   -74910
         Top             =   510
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin MSComDlg.CommonDialog CD2 
         Left            =   -75000
         Top             =   1140
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "PDF Files Only|*.pdf"
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
         Index           =   2
         Left            =   -74910
         TabIndex        =   95
         Top             =   4695
         Width           =   2040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Barcode Program Exe File Path"
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
         Index           =   0
         Left            =   -74085
         TabIndex        =   44
         Top             =   825
         Width           =   3360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Barcode File Path"
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
         Left            =   -74085
         TabIndex        =   43
         Top             =   1725
         Width           =   1905
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4710
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   3
      Left            =   450
      TabIndex        =   12
      Tag             =   "CustomDescription"
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "3625;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer Code"
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
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   2
      Left            =   105
      TabIndex        =   10
      Tag             =   "Vend_AccNo"
      Top             =   225
      Visible         =   0   'False
      Width           =   1725
      VariousPropertyBits=   679495711
      MaxLength       =   9
      BorderStyle     =   1
      Size            =   "3043;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   4875
      TabIndex        =   8
      Top             =   7740
      Width           =   1695
      Caption         =   " Save            "
      PicturePosition =   327683
      Size            =   "2990;661"
      Picture         =   "frmEditItem_VC.frx":0054
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
      Height          =   375
      Left            =   6630
      TabIndex        =   9
      Top             =   7740
      Width           =   1695
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2990;661"
      Picture         =   "frmEditItem_VC.frx":04A2
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmEditItem_VC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEdit As String

Dim dFOB As Double, dCIFSea As Double, dCIFAir As Double, dFOBSialkot As Double, dCnFAir As Double, dCnFSea As Double, dCnISea As Double, dCnIAir As Double
Dim dFOBT As Double, dCIFSeaT As Double, dCIFAirT As Double, dFOBSialkotT As Double, dCnFAirT As Double, dCnFSeaT As Double, dCnISeaT As Double, dCnIAirT As Double

Dim iEntryID As Long
Dim strDesc As String
Dim strCustomerItemCode As String

Public Sub EditItem(TempstrEdit As String)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    strEdit = TempstrEdit
    
    TBox(3).Locked = True
    Dim c As Control
    With rs
        .Open "SELECT * FROM Vendor_Catalog WHERE Vend_AccNo+ItemID='" & strEdit & "'", con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    c.ID = .Fields(c.Tag) & ""
                Else
                    If c.Name = "NBox" Then
                        c.Text = Round(Val(.Fields(c.Tag) & ""), 4)
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        
        strCustomerItemCode = !ITemID & ""
        
        dFOB = Round(Val(![FOB] & ""), 4)
        dFOBSialkot = Round(Val(![ExWorks] & ""), 4)
        dCIFAir = Round(Val(![CIFAir] & ""), 4)
        dCIFSea = Round(Val(![CIFSea] & ""), 4)
        dCnFSea = Round(Val(![CnFSea] & ""), 4)
        dCnFAir = Round(Val(![CnFAir] & ""), 4)
        dCnISea = Round(Val(![CnISea] & ""), 4)
        dCnIAir = Round(Val(![CnIAir] & ""), 4)
        '-------------------------------------------
        dFOBT = Round(Val(![FOBTop] & ""), 4)
        dFOBSialkotT = Round(Val(![ExWorksTop] & ""), 4)
        dCIFAirT = Round(Val(![CIFAirTop] & ""), 4)
        dCIFSeaT = Round(Val(![CIFSeaTop] & ""), 4)
        dCnFSeaT = Round(Val(![CnFSeaTop] & ""), 4)
        dCnFAirT = Round(Val(![CnFAirTop] & ""), 4)
        dCnISeaT = Round(Val(![CnISeaTop] & ""), 4)
        dCnIAirT = Round(Val(![CnIAirTop] & ""), 4)
        
        strDesc = ![Description] & ""
        
        iEntryID = Val(![EntryID] & "")
        txtFilePath = !BarcodeFile & ""
        
        .Close
        
'        .Open "SELECT EntryID,DT,UserName FROM FCustCatalRevs WHERE RefID=" & iEntryID, con, adOpenForwardOnly, adLockReadOnly
'        lstRev.Clear
'        Dim strTemp As String
'        Do Until .EOF
'            If IsNull(!UserName) Then
'                strTemp = Format(![DT], "dd-MMM-yyyy   HH:MM AMPM")
'            Else
'                strTemp = "(" & !UserName & ") " & Format(![DT], "dd-MMM-yyyy   HH:MM AMPM")
'            End If
'
'            lstRev.AddItem strTemp 'Format(![DT], "dd-MMM-yyyy   HH:MM AMPM")
'            lstRev.ItemData(lstRev.ListCount - 1) = ![EntryID]
'            .MoveNEXT
'        Loop
'        lstRev.AddItem "<Current Prices>"
'        lstRev.ListIndex = lstRev.ListCount - 1
'        .Close
        
    End With
    
    Set rs = Nothing
    TBox(1).SelStart = 0
    TBox(1).SelLength = Len(TBox(1).Text)
    
    
    Dim bAllowSave As Boolean
    bAllowSave = True 'GetSingleBooleanValue("EditCustCatalog", "Users", " WHERE UserName='" & CurrentUserName & "'")
    cmdOk.Enabled = bAllowSave
    
    Me.Show 1
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillCmbs()
        
    
    '----------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select UnitName From Units Order By UnitName", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbUnit.AddItem .Fields(0) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Call AddToCombo(cmbFinQuality, "Description", "ItemsFinQuality", , , "Description")
    '----------------------------------------------------------------
End Sub

Private Sub cmdBrowseBarcode_Click(Index As Integer)

    On Error GoTo err
    If Index = 0 Then
        CDBarcode.Filter = "Program Files|*.exe"
        On Error Resume Next
        CDBarcode.ShowOpen
        If err.Number = 32755 Then Exit Sub
        If CDBarcode.FileName Then
            txtProgramPath = CDBarcode.FileName
            Call SaveSetting(App.Title, "Barcode", "ProgramPath", txtProgramPath)
        End If
    ElseIf Index = 1 Then
        CDBarcode.Filter = "All Files|*.*"
        On Error Resume Next
        CDBarcode.ShowOpen
        If err.Number = 32755 Then Exit Sub
        If CDBarcode.FileName Then
            txtFilePath = CDBarcode.FileName
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdBrowsePDF_Click()
    
    On Error Resume Next
    Dim Picfile As String
    CD2.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD2.FileName
        txtPDFPath = Picfile
    End If


End Sub

Private Sub cmdOpenBarcode_Click()

    On Error GoTo err
    
    If txtProgramPath = "" Then
        MsgBox "Please Enter Path.", vbInformation
        Exit Sub
    ElseIf txtFilePath = "" Then
        MsgBox "Please Enter File Path.", vbInformation
        Exit Sub
    End If
    Shell """" & txtProgramPath & """" & " " & """" & txtFilePath & """", vbNormalFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then
        frmVendor_Catalog.RefreshLV
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    txtProgramPath = GetSetting(App.Title, "Barcode", "ProgramPath")
    SSTab1.Tab = 0
    Me.KeyPreview = True
    'DTPicker1.value = ServerDate
    FillCmbs
    
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err

    Dim OpBalance As Double
    
    If Trim(cmbUnit) = "" Then
        MsgBox "No Unit Selected.", vbInformation
        SaveItem = False
        Exit Function
    ElseIf TBox(1) = "" Then
        MsgBox "Invalid Item ID.", vbInformation
        Exit Function
    End If
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(ItemID)", "Vendor_Catalog", " WHERE ItemID='" & TBox(1) & "' AND Vend_AccNo='" & TBox(2) & "' AND Vend_AccNo+ItemID<>'" & strEdit & "'")
    
    If lCount > 0 Then
        If MsgBox("Customer Code Exists, Do you want to Continue?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    Dim SQL As String
    
    Call StartTrans(con)

    SQL = CreateUpdateSQL("Vendor_Catalog")
    SQL = SQL & " WHERE Vend_AccNo+ItemID='" & strEdit & "'"
    
    con.Execute SQL
    con.Execute "UPDATE Vendor_Catalog SET BarcodeFile='" & Replace(txtFilePath, "'", "''") & _
     "' WHERE Vend_AccNo+ItemID='" & strEdit & "'"
    
    Dim bRegularPriceChanged As Boolean, bTopPriceChanged As Boolean
    If Val(NBox(0)) <> dFOB Or Val(NBox(8)) <> dCIFSea Or Val(NBox(9)) <> dCIFAir Or Val(NBox(10)) <> dCnFSea Or Val(NBox(11)) <> dCnFAir Or dFOBSialkot <> Val(NBox(1)) Then
        bRegularPriceChanged = True
    Else
        bRegularPriceChanged = False
    End If
    
    If Val(NBox(17)) <> dFOBT Or Val(NBox(15)) <> dCIFSeaT Or Val(NBox(16)) <> dCIFAirT Or Val(NBox(4)) <> dCnFSeaT Or Val(NBox(5)) <> dCnFAirT Or dFOBSialkotT <> Val(NBox(20)) Then
        bTopPriceChanged = True
    Else
        bTopPriceChanged = False
    End If
    
'    If (bRegularPriceChanged Or bTopPriceChanged) And (lstRev.ListIndex = (lstRev.ListCount - 1)) Then
'        con.Execute "INSERT INTO FCustCatalRevs(RefID,Description,FOB,CIFSea,CIFAir,CnFSea,CnFAir,FOBSialkot,CnISea," & _
'         "CnIAir,FOBTop,CIFSeaTop,CIFAirTop,CnFSeaTop,CnFAirTop,ExWorksTop,CnISeaTop,CnIAirTop,DT,UserName) " & _
'         "VALUES (" & iEntryID & ",'" & strDesc & "'," & dFOB & "," & dCIFSea & "," & dCIFAir & "," & _
'         dCnFSea & "," & dCnFAir & "," & dFOBSialkot & "," & dCnISea & "," & dCnIAir & "," & dFOBT & "," & _
'         dCIFSeaT & "," & dCIFAirT & "," & dCnFSeaT & "," & dCnFAirT & "," & dFOBSialkotT & "," & _
'         dCnISeaT & "," & dCnIAirT & ",'" & Now & "','" & CurrentUserName & "')"
'    End If
         
    'Update CustCode in FOrderItems
'    If strCustomerItemCode <> TBox(1) Then
'        con.Execute "UPDATE FOrderItems SET ItemCode='" & TBox(1) & "' FROM FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo=FOrderItems.OrderNo " & _
'         "WHERE CustCode='" & TBox(2) & "' AND Country='" & TBox(5) & "' AND ItemCode='" & strCustomerItemCode & "'"
'    End If
'
    con.CommitTrans
     
    Call AttachPDF(strEdit)
    
    strEdit = ""
    SaveItem = True
    
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub lstRev_Click()
    On Error GoTo err
    
    If lstRev.ListIndex = -1 Then Exit Sub
    
    If lstRev.ListIndex = lstRev.ListCount - 1 Then
    
        TBox(0) = strDesc
        '--------------------------------------------
        NBox(0) = dFOB
        NBox(8) = dCIFSea
        NBox(9) = dCIFAir
        NBox(10) = dCnFSea
        NBox(11) = dCnFAir
        NBox(1) = dFOBSialkot
        NBox(2) = dCnISea
        NBox(3) = dCnIAir
        'Top
        NBox(17) = dFOBT
        NBox(20) = dFOBSialkotT
        NBox(19) = dCnISeaT
        NBox(18) = dCnIAirT
        NBox(15) = dCIFSeaT
        NBox(16) = dCIFAirT
        NBox(4) = dCnFSeaT
        NBox(5) = dCnFAirT
    Else
        Dim rs As New ADODB.Recordset
'        With rs
'            .Open "SELECT * FROM FCustCatalRevs WHERE EntryID=" & lstRev.ItemData(lstRev.ListIndex), con, adOpenForwardOnly, adLockReadOnly
'
'            NBox(0) = Val(![FOB] & "")
'            NBox(8) = Val(![CIFSea] & "")
'            NBox(9) = Val(![CIFAir] & "")
'            NBox(10) = Val(![CnFSea] & "")
'            NBox(11) = Val(![CnFAir] & "")
'            TBox(0) = ![Description] & ""
'            NBox(1) = Val(!FOBSialkot & "")
'            NBox(2) = Val(!CnISea & "")
'            NBox(3) = Val(!CnIAir & "")
'
'            TBox(0) = !Description & ""
'            '------------------------------------
'            'Top
'            NBox(17) = Val(!FOBTop & "")
'            NBox(20) = Val(!ExWorksTop & "")
'            NBox(19) = Val(!CnISeaTop & "")
'            NBox(18) = Val(!CnIAirTop & "")
'            NBox(15) = Val(!CIFSeaTop & "")
'            NBox(16) = Val(!CIFAirTop & "")
'            NBox(4) = Val(!CnFSeaTop & "")
'            NBox(5) = Val(!CnFAirTop & "")
'            .Close
'        End With
        Set rs = Nothing
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub NBox_Change(Index As Integer)
    Select Case Index
        Case 4, 5, 7
            'NBox(6) = Val(NBox(4)) - Val(NBox(5)) - Val(NBox(7))
    End Select
    
End Sub



Private Sub Tbox_Change(Index As Integer)
    Select Case Index
    
    Case 8, 3
        'TBox(4) = TBox(8) & TBox(3)
    End Select
    
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Exit Sub
    Select Case Index
    Case 2, 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 5
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
    
End Sub

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)
    Exit Sub
    If Index = 5 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0.00")
    ElseIf Index = 2 Or Index = 3 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0")
    End If
End Sub

Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "UPDATE " & strTable & " SET "
    '---------------------------------------------------------------------------
    For Each c In Controls
    
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Then '
                If c.Name = "NBox" Then
                    If lstRev.ListIndex = (lstRev.ListCount - 1) Then
                        RetString = RetString & c.Tag & "=" & Val(c.Text) & ","
                    End If
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.ID & "',"
            End If
        End If
    Next
    
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function

Private Sub AttachPDF(p_strEdit As String)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtPDFPath
    If strPDFPath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,ArtworkPDF FROM Vendor_Catalog WHERE Vend_AccNo+ItemID='" & p_strEdit & "'", con, adOpenStatic, adLockOptimistic
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


