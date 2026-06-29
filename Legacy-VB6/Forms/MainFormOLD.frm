VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MainFormOLD 
   BackColor       =   &H8000000C&
   Caption         =   "Contractor Package"
   ClientHeight    =   9450
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   6870
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":0000
            Key             =   "test"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicTitle 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   10665
      TabIndex        =   2
      Top             =   0
      Width           =   10725
   End
   Begin VB.PictureBox PicBar 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      Height          =   8880
      Left            =   0
      ScaleHeight     =   8820
      ScaleWidth      =   2325
      TabIndex        =   0
      Top             =   570
      Width           =   2385
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   8415
         Left            =   0
         TabIndex        =   1
         Top             =   -30
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   14843
         BandCount       =   15
         ImageList       =   "ImageList4"
         BandBorders     =   0   'False
         VariantHeight   =   0   'False
         Orientation     =   1
         _CBWidth        =   2310
         _CBHeight       =   8415
         _Version        =   "6.0.8169"
         Image1          =   "1"
         Child1          =   "PicToolBar (1)"
         MinHeight1      =   2250
         Width1          =   1665
         NewRow1         =   0   'False
         BandTag1        =   "Company"
         Child2          =   "PicToolBar (2)"
         MinHeight2      =   1425
         Width2          =   1665
         NewRow2         =   0   'False
         BandTag2        =   "Policies"
         Child3          =   "PicToolBar (3)"
         MinHeight3      =   1425
         Width3          =   2415
         NewRow3         =   0   'False
         BandTag3        =   "BusPart"
         Child4          =   "PicToolBar (4)"
         MinHeight4      =   1425
         Width4          =   1470
         NewRow4         =   0   'False
         BandTag4        =   "Orders"
         Child5          =   "PicToolBar (5)"
         MinHeight5      =   1425
         Width5          =   1470
         NewRow5         =   0   'False
         BandTag5        =   "Inventory"
         Child6          =   "PicToolBar (6)"
         MinHeight6      =   1425
         Width6          =   1470
         NewRow6         =   0   'False
         BandTag6        =   "Purchases"
         Child7          =   "PicToolBar (7)"
         MinHeight7      =   1425
         Width7          =   1560
         NewRow7         =   0   'False
         BandTag7        =   "Production"
         MinHeight8      =   360
         NewRow8         =   0   'False
         BandTag8        =   "Payroll"
         MinHeight9      =   360
         NewRow9         =   0   'False
         BandTag9        =   "Accounts"
         MinHeight10     =   360
         NewRow10        =   0   'False
         BandTag10       =   "Export"
         MinHeight11     =   360
         NewRow11        =   0   'False
         BandTag11       =   "Management"
         MinHeight12     =   360
         NewRow12        =   0   'False
         BandTag12       =   "Stats"
         MinHeight13     =   360
         NewRow13        =   0   'False
         BandTag13       =   "Glossary"
         MinHeight14     =   360
         NewRow14        =   0   'False
         BandTag14       =   "Developers"
         MinHeight15     =   360
         NewRow15        =   0   'False
         BandTag15       =   "test"
         Begin VB.PictureBox PicToolBar 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   15
            Index           =   7
            Left            =   435
            ScaleHeight     =   15
            ScaleWidth      =   1425
            TabIndex        =   21
            Top             =   7305
            Width           =   1425
            Begin MSComctlLib.Toolbar ToolBar 
               Height          =   1320
               Index           =   7
               Left            =   0
               Negotiate       =   -1  'True
               TabIndex        =   22
               Top             =   0
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   2328
               ButtonWidth     =   3201
               ButtonHeight    =   582
               Style           =   1
               TextAlignment   =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   4
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Requsition"
                     Key             =   "Requsition"
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Production Planning"
                     Key             =   "PPlan"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Authorise"
                     Key             =   "Authorise"
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Processes"
                     Key             =   "Processes"
                  EndProperty
               EndProperty
               MousePointer    =   99
            End
            Begin MSForms.ScrollBar SBar 
               Height          =   1605
               Index           =   7
               Left            =   1560
               TabIndex        =   23
               Top             =   0
               Width           =   225
               Size            =   "397;2831"
               Min             =   1
               Max             =   3
               Position        =   1
            End
         End
         Begin VB.PictureBox PicToolBar 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   15
            Index           =   6
            Left            =   435
            ScaleHeight     =   15
            ScaleWidth      =   1425
            TabIndex        =   18
            Top             =   7170
            Width           =   1425
            Begin MSComctlLib.Toolbar ToolBar 
               Height          =   330
               Index           =   6
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   582
               ButtonWidth     =   2699
               ButtonHeight    =   582
               Style           =   1
               TextAlignment   =   1
               _Version        =   393216
               MousePointer    =   99
            End
            Begin MSForms.ScrollBar SBar 
               Height          =   1605
               Index           =   6
               Left            =   1560
               TabIndex        =   20
               Top             =   0
               Width           =   225
               Size            =   "397;2831"
               Min             =   1
               Max             =   3
               Position        =   1
            End
         End
         Begin VB.PictureBox PicToolBar 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   15
            Index           =   5
            Left            =   435
            ScaleHeight     =   15
            ScaleWidth      =   1425
            TabIndex        =   15
            Top             =   7035
            Width           =   1425
            Begin MSComctlLib.Toolbar ToolBar 
               Height          =   330
               Index           =   5
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   582
               ButtonWidth     =   2699
               ButtonHeight    =   582
               Style           =   1
               TextAlignment   =   1
               _Version        =   393216
               MousePointer    =   99
            End
            Begin MSForms.ScrollBar SBar 
               Height          =   1605
               Index           =   5
               Left            =   1560
               TabIndex        =   17
               Top             =   0
               Width           =   225
               Size            =   "397;2831"
               Min             =   1
               Max             =   3
               Position        =   1
            End
         End
         Begin VB.PictureBox PicToolBar 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   930
            Index           =   4
            Left            =   435
            ScaleHeight     =   930
            ScaleWidth      =   1425
            TabIndex        =   12
            Top             =   5910
            Width           =   1425
            Begin MSComctlLib.Toolbar ToolBar 
               Height          =   330
               Index           =   4
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   582
               ButtonWidth     =   2699
               ButtonHeight    =   582
               Style           =   1
               TextAlignment   =   1
               _Version        =   393216
               MousePointer    =   99
            End
            Begin MSForms.ScrollBar SBar 
               Height          =   1605
               Index           =   4
               Left            =   1560
               TabIndex        =   14
               Top             =   0
               Width           =   225
               Size            =   "397;2831"
               Min             =   1
               Max             =   3
               Position        =   1
            End
         End
         Begin VB.PictureBox PicToolBar 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   1410
            Index           =   1
            Left            =   30
            ScaleHeight     =   1410
            ScaleWidth      =   2250
            TabIndex        =   9
            Top             =   225
            Width           =   2250
            Begin MSComctlLib.Toolbar ToolBar 
               Height          =   1980
               Index           =   1
               Left            =   0
               TabIndex        =   10
               Top             =   0
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   3493
               ButtonWidth     =   2963
               ButtonHeight    =   582
               Style           =   1
               TextAlignment   =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   6
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Company Detail"
                     Key             =   "CompDetail"
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Company Catalog"
                     Key             =   "CompCatalog"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Foreign Catalog"
                     Key             =   "ForCatalog"
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Departments"
                     Key             =   "CompDept"
                  EndProperty
                  BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Resources"
                     Key             =   "CompResources"
                  EndProperty
                  BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Banks"
                     Key             =   "Banks"
                  EndProperty
               EndProperty
               MousePointer    =   99
            End
            Begin MSForms.ScrollBar SBar 
               Height          =   1605
               Index           =   1
               Left            =   1950
               TabIndex        =   11
               Top             =   0
               Width           =   225
               Size            =   "397;2831"
               Min             =   1
               Max             =   3
               Position        =   1
            End
         End
         Begin VB.PictureBox PicToolBar 
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            Height          =   1470
            Index           =   2
            Left            =   435
            ScaleHeight     =   1470
            ScaleWidth      =   1425
            TabIndex        =   6
            Top             =   1830
            Width           =   1425
            Begin MSComctlLib.Toolbar ToolBar 
               Height          =   2640
               Index           =   2
               Left            =   -15
               TabIndex        =   7
               Top             =   0
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   4657
               ButtonWidth     =   2249
               ButtonHeight    =   582
               Style           =   1
               TextAlignment   =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   8
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Holidays"
                     Key             =   "Holidays"
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Pymt. Policy"
                     Key             =   "PymtPolicy"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Job Policy"
                     Key             =   "JobPolicy"
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Adv. Policy"
                     Key             =   "AdvPolicy"
                  EndProperty
                  BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Short Term"
                     Key             =   "ShortTerm"
                  EndProperty
                  BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Long Term"
                     Key             =   "LongTerm"
                  EndProperty
                  BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Resources"
                     Key             =   "Resources"
                  EndProperty
                  BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Bonus"
                     Key             =   "Bonus"
                  EndProperty
               EndProperty
               MousePointer    =   99
            End
            Begin MSForms.ScrollBar SBar 
               Height          =   1605
               Index           =   2
               Left            =   1560
               TabIndex        =   8
               Top             =   0
               Width           =   225
               Size            =   "397;2831"
               Min             =   1
               Max             =   3
               Position        =   1
            End
         End
         Begin VB.PictureBox PicToolBar 
            BorderStyle     =   0  'None
            Height          =   2220
            Index           =   3
            Left            =   435
            ScaleHeight     =   2220
            ScaleWidth      =   1425
            TabIndex        =   3
            Top             =   3495
            Width           =   1425
            Begin MSComctlLib.Toolbar ToolBar 
               Height          =   4290
               Index           =   3
               Left            =   -15
               TabIndex        =   4
               Top             =   0
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   7567
               ButtonWidth     =   2461
               ButtonHeight    =   582
               Style           =   1
               TextAlignment   =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   13
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "F. Customers"
                     Key             =   "FCustomers"
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "L. Customers"
                     Key             =   "LCustomers"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Importers"
                     Key             =   "Importers"
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Employees"
                     Key             =   "Employees"
                  EndProperty
                  BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Contractors"
                     Key             =   "Contractors"
                  EndProperty
                  BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "R.M. Venders"
                     Key             =   "RMVenders"
                  EndProperty
                  BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Steel Venders"
                     Key             =   "SteelVend"
                  EndProperty
                  BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Suppliers"
                     Key             =   "GeneralSupp"
                  EndProperty
                  BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Providers"
                     Key             =   "ServiceProviders"
                  EndProperty
                  BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Test Labs"
                     Key             =   "TestLabs"
                  EndProperty
                  BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Banks"
                     Key             =   "Banks"
                  EndProperty
                  BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "S. Agencies"
                     Key             =   "SAgencies"
                     Object.ToolTipText     =   "Shipping Agencies"
                  EndProperty
                  BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "C. Agencies"
                     Key             =   "C.Agencies"
                     Object.ToolTipText     =   "Courier Agencies"
                  EndProperty
               EndProperty
               MousePointer    =   99
            End
            Begin MSForms.ScrollBar SBar 
               Height          =   1605
               Index           =   3
               Left            =   1560
               TabIndex        =   5
               Top             =   0
               Width           =   225
               Size            =   "397;2831"
               Min             =   1
               Max             =   3
               Position        =   1
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8340
      Top             =   1920
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   7950
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8610
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please Select Payroll Database to Connect."
      Filter          =   """Payroll Database|Payroll.mdb"""
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6870
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":02A1
            Key             =   "Cont"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":07A3
            Key             =   "PaySheet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":0842
            Key             =   "PayEntry"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":0920
            Key             =   "Issue"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":09F5
            Key             =   "Recieve"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":0A97
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":0B7A
            Key             =   "ContLedger"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5895
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   130
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":115C
            Key             =   "test"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":282E
            Key             =   "Company"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":2E15
            Key             =   "Glossary"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":33F9
            Key             =   "Management"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":3BB7
            Key             =   "Policies"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":4542
            Key             =   "Purchases"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":4BBD
            Key             =   "BusPart"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":557B
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":5F3E
            Key             =   "Developers"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":6633
            Key             =   "Inventory"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":6FD4
            Key             =   "Orders"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":7912
            Key             =   "Payroll"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":834B
            Key             =   "Production"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":8A03
            Key             =   "Stats"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":9011
            Key             =   "Accounts"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5895
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormOLD.frx":96ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Use Calculator"
         Index           =   1
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "User Manager"
         Index           =   2
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Change Password"
         Index           =   4
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Log Off"
         Index           =   5
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "General Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Set Initial Lot Process"
         Index           =   1
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Percentage Ranges"
         Index           =   3
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&Components"
         Index           =   5
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&Manufacturing Processes"
         Index           =   7
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&FDA Processes"
         Index           =   8
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Steel Vendors"
         Index           =   10
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Test Labs"
         Index           =   11
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Measuring &Units"
         Index           =   13
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Compositions"
         Index           =   14
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Cities && Ports"
         Index           =   16
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Bank List"
         Index           =   18
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Payment Terms"
         Index           =   19
      End
   End
   Begin VB.Menu mnupopemp 
      Caption         =   "PopMenuEmp"
      Visible         =   0   'False
      Begin VB.Menu mnuempRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuempdash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowInactive 
         Caption         =   "Show Inactive Employees/Departments"
      End
      Begin VB.Menu mnuempdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMake 
         Caption         =   "Make This Employee Inactive"
      End
      Begin VB.Menu dd1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuempNewDept 
         Caption         =   "Add New &Department"
      End
      Begin VB.Menu mnuEmpNewEmp 
         Caption         =   "Add New &Employee"
      End
      Begin VB.Menu mnuempdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuEmpdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssItems 
         Caption         =   "Show &Assigned Items"
      End
      Begin VB.Menu mnuEmpPrint 
         Caption         =   "&Print Employee List"
      End
   End
   Begin VB.Menu mnupopdept 
      Caption         =   "popMenuDept"
      Visible         =   0   'False
      Begin VB.Menu mnuDeptRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu popdeptdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowDept 
         Caption         =   "Show Inactive Departments"
      End
      Begin VB.Menu mnuDeptdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptInact 
         Caption         =   "Make This Department Inactive"
      End
      Begin VB.Menu mnuDeptdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptnew 
         Caption         =   "Add &New Department"
      End
      Begin VB.Menu mnuDepdash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeptdash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptPrint 
         Caption         =   "&Print Department List"
      End
   End
   Begin VB.Menu mnuVenders 
      Caption         =   "Venders"
      Visible         =   0   'False
      Begin VB.Menu mnuVendIssItems 
         Caption         =   "Issue Items By Orders"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendIssItemsByItems 
         Caption         =   "Issue Items By Items"
         Visible         =   0   'False
      End
      Begin VB.Menu mnurecdash1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendRcvItems 
         Caption         =   "Recieve Items From Venders"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendRcvItemsFromSupervisor 
         Caption         =   "Recieve Items From Supervisor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendRcvWst 
         Caption         =   "Recieve Wastage"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendRepairDash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIssRepair 
         Caption         =   "Issue Repair"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRcvRepair 
         Caption         =   "Recieve Repair"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRcvRepairWastage 
         Caption         =   "Receive Repair Wastage"
         Visible         =   0   'False
      End
      Begin VB.Menu VendPaydash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendPymt 
         Caption         =   "Vender Advance Payment"
      End
      Begin VB.Menu mnuVendPaySheetdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendPaySheet 
         Caption         =   "Vender Payment Sheet"
      End
      Begin VB.Menu mnuAdvdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendLoans 
         Caption         =   "Loans"
         Begin VB.Menu mnuVendAdvShort 
            Caption         =   "New Short Term Advance"
         End
         Begin VB.Menu mnuVendAdvLong 
            Caption         =   "New Long Term Loan"
         End
         Begin VB.Menu mnuVendLoanLedgersdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVendShortLoanLedger 
            Caption         =   "Short Term Loan Ledger"
         End
         Begin VB.Menu mnuVendLongLoanLedger 
            Caption         =   "Long Term Loan Ledger"
         End
      End
   End
   Begin VB.Menu mnuRMVend 
      Caption         =   "Raw Material"
      Visible         =   0   'False
      Begin VB.Menu mnuRMGroups 
         Caption         =   "RM Groups"
      End
      Begin VB.Menu mnuRM 
         Caption         =   "Raw Material"
      End
      Begin VB.Menu RMPOdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRMPO 
         Caption         =   "New Purchase Order"
      End
      Begin VB.Menu mnuPOrderList 
         Caption         =   "Purchase Order List"
      End
      Begin VB.Menu RRMdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecRM 
         Caption         =   "Receive Raw Material"
      End
      Begin VB.Menu mnuIssRM 
         Caption         =   "Issue Raw Material"
      End
   End
   Begin VB.Menu mnuHICORes 
      Caption         =   "Hico Resources"
      Visible         =   0   'False
      Begin VB.Menu mnuHRGroups 
         Caption         =   "Hico Resources Groups"
      End
   End
   Begin VB.Menu mnuCompanya 
      Caption         =   "Company"
      Visible         =   0   'False
      Begin VB.Menu mnuItems 
         Caption         =   "Items"
         Begin VB.Menu mnuItmGrps 
            Caption         =   "Item Groups"
         End
         Begin VB.Menu mnuItemsFinQuality 
            Caption         =   "Item's Finishing Quality"
         End
         Begin VB.Menu mnuItemTypes 
            Caption         =   "Item Types"
         End
      End
   End
   Begin VB.Menu mnuCustomers 
      Caption         =   "Customers"
      Visible         =   0   'False
      Begin VB.Menu mnuNewOrd 
         Caption         =   "New Order"
      End
      Begin VB.Menu mnuNewStockOrder 
         Caption         =   "New Stock Order"
      End
      Begin VB.Menu mnuCustOrdList 
         Caption         =   "Customer Order List"
      End
      Begin VB.Menu custdash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdStatus 
         Caption         =   "Order Status"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrdStatus1 
         Caption         =   "Order Status(Other Format)"
      End
      Begin VB.Menu mnuOrdStatusComp 
         Caption         =   "Order Status Complete"
         Visible         =   0   'False
      End
      Begin VB.Menu custdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewProforma 
         Caption         =   "New Proforma"
      End
      Begin VB.Menu mnuCustProformaList 
         Caption         =   "Customer Proforma List"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuDailyRpts 
         Caption         =   "Daily Reports"
      End
      Begin VB.Menu mnuDailyRptsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIssRpts 
         Caption         =   "Issue Reports"
      End
      Begin VB.Menu mnuRcvRpts 
         Caption         =   "Receive Reports"
      End
      Begin VB.Menu mnuRcvRptsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRMListrpt 
         Caption         =   "Raw Material List"
      End
      Begin VB.Menu mnuRMVendListrpt 
         Caption         =   "Raw Material Vender List"
      End
      Begin VB.Menu mnuRMIssRpts 
         Caption         =   "Raw Material Issue Reports"
      End
      Begin VB.Menu mnuRMRcvRpts 
         Caption         =   "Raw Material ReceiveReports"
      End
      Begin VB.Menu mnuRMRcvRptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHRrpt 
         Caption         =   "Hico Resources"
      End
      Begin VB.Menu mnuAssHRrpt 
         Caption         =   "Assigned Hico Resources"
      End
      Begin VB.Menu mnuCustListrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustListrpt 
         Caption         =   "Customer List"
      End
      Begin VB.Menu mnuReqrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReqrpt 
         Caption         =   "Requsition Report"
      End
      Begin VB.Menu mnuPORpt 
         Caption         =   "P.O. Report"
      End
      Begin VB.Menu mnuPORptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProdPlanrpt 
         Caption         =   "Production Planning Report"
      End
      Begin VB.Menu mnuPPlanrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendListrpt 
         Caption         =   "Vender List"
      End
   End
   Begin VB.Menu mnuProdPhases 
      Caption         =   "Production Phases"
      Visible         =   0   'False
      Begin VB.Menu mnuReq 
         Caption         =   "Requsition"
         Visible         =   0   'False
      End
      Begin VB.Menu prodplandash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdPlan 
         Caption         =   "Prdoduction Planning"
         Visible         =   0   'False
      End
      Begin VB.Menu TestDash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAuth 
         Caption         =   "Authorize"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAuthDash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FORGING"
         Index           =   1
         Begin VB.Menu mnuIssForg 
            Caption         =   "Issue Items"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecVForg 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSForg 
            Caption         =   "Receive Items From Supervisor"
         End
         Begin VB.Menu mnudForg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastForg 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuForgSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuForgSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "MACHINING"
         Index           =   2
         Begin VB.Menu mnuIssMach 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVMach 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSMach 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudMach 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastMach 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuMachSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMachSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "DRILLING"
         Index           =   3
         Begin VB.Menu mnuIssDril 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVDril 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSDril 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudDril 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastDril 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuDrilSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDrilSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FILLING"
         Index           =   4
         Begin VB.Menu mnuIssFill 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFill 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFill 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFill 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFill 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFillSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFillSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FUNTIONAL INSPECTION UN-FINISH"
         Index           =   5
         Begin VB.Menu mnuIssFunc 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFunc 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFunc 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFunc 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFunc 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFuncSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFuncSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "HEAT TREATMENT"
         Index           =   6
         Begin VB.Menu mnuIssHeat 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVHeat 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSHeat 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudHeat 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastHeat 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuHeatSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHeatSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "POLISHING"
         Index           =   7
         Begin VB.Menu mnuIssPoli 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPoli 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPoli 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecWastPoli 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnudPoli 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIssRepairPoli 
            Caption         =   "Issue Repair"
         End
         Begin VB.Menu mnuRcvRepairPoli 
            Caption         =   "Receive Repair"
         End
         Begin VB.Menu mnuPoliSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPoliSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "PASSIVATION"
         Index           =   8
         Begin VB.Menu mnuIsspass 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPass 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPass 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPass 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPass 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPassSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPassSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FUNTIONAL INSPECTION"
         Index           =   9
         Begin VB.Menu mnuIssFunt 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFunt 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFunt 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFunt 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFunt 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFuntdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFuntSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "POLISHING && RE-BUFFING"
         Index           =   10
         Begin VB.Menu mnuIssPoRe 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPoRe 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPoRe 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPoRe 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPoRe 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPORESSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPORESS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "ULTRASONIC CLEAING"
         Index           =   11
         Begin VB.Menu mnuIssUltr 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVUltr 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSUltr 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudUltr 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastUltr 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuUltrSSDash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUltrSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "PACKING && LABELING"
         Index           =   12
         Begin VB.Menu mnuIssPack 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPack 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPack 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPack 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPack 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPackSSDash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPackSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "HOLDING && DISTRIBUTION"
         Index           =   13
         Begin VB.Menu mnuIssHold 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVHold 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSHold 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudHold 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastHold 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuHoldSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHoldSS 
            Caption         =   "Stock Status"
         End
      End
   End
   Begin VB.Menu mnumainSteel 
      Caption         =   "Steel"
      Visible         =   0   'False
      Begin VB.Menu mnuSteel 
         Caption         =   "&Steel List"
      End
      Begin VB.Menu mnuGuage 
         Caption         =   "Guage List"
      End
   End
   Begin VB.Menu mnuPopCust 
      Caption         =   "Customers"
      Visible         =   0   'False
      Begin VB.Menu PopSubCust 
         Caption         =   "Add New Customer"
         Index           =   0
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer List"
         Index           =   1
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopProdVend 
      Caption         =   "Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Add New Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender List"
         Index           =   1
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopRMVend 
      Caption         =   "R.M. Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubRMVend 
         Caption         =   "Add New R.M. Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "R.M. Vender List"
         Index           =   1
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "R.M. Vender Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopSteelVend 
      Caption         =   "Steel Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubSteelVend 
         Caption         =   "Add New Steel Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubSteelVend 
         Caption         =   "Steel Vender List"
         Index           =   1
      End
   End
End
Attribute VB_Name = "MainFormOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dashBankLIst_Click()
    frmBankList.Show
End Sub


Private Sub BusPartToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Customers"
            Me.PopUpMenu mnuPopCust, 512, PicBar.Width
        Case "ProdVend"
            Me.PopUpMenu mnuPopProdVend, 512, PicBar.Width
        Case "RMVend"
            Me.PopUpMenu mnuPopRMVend, 512, PicBar.Width
    End Select
End Sub


Private Sub CoolBar1_Resize()
    For i = 1 To CoolBar1.Bands.Count
        If CoolBar1.Bands(i).NewRow Then CoolBar1.Bands(i).NewRow = False
    Next i
End Sub


Private Sub MDIForm_Load()

    Dim i As Long
    'mnucont.Visible = CONTRACTORS_ENABLED
    'Toolbar1.Buttons("Cont").Visible = CONTRACTORS_ENABLED
    'Toolbar1.Buttons("ContLedger").Visible = CONTRACTORS_ENABLED
    On Error Resume Next
    
    'Dim ctr As Control
    'For Each ctr In Controls
    '    If TypeOf ctr Is Menu Then
    '        ctr.Visible = False
    '    End If
    'Next ctr
    
    On Error GoTo err
    
    'Hide The Title Bar  Of Main Form.
    SetWindowLong Me.hwnd, -16, WS_DLGFRAME
    
    
    PicBar.Visible = False
    CoolBar1.Visible = False
    CoolBar1.Width = PicBar.Width
    
    
    
    With CoolBar1
        Set CoolBar1.ImageList = ImageList2
        For i = 1 To .Bands.Count
            .Bands(i).Image = ImageList2.ListImages(.Bands(i).Tag).Key
        Next i
    End With
    
    
    
    'Set Scrolls
    For i = SBar.LBound To SBar.UBound
        If ToolBar(i).Buttons.Count > 6 Then
            SBar(i).Max = ToolBar(i).Buttons.Count - 5
        Else
            SBar(i).Visible = False
        End If
    Next i
    
    PicBar.Visible = True
    CoolBar1.Visible = True
    
    
    For i = 1 To CoolBar1.Bands.Count - 1
        CoolBar1.Bands(i).Width = 420
    Next i
    
    Exit Sub
    
    With Toolbar1
        Set Toolbar1.ImageList = ImageList1
        For i = 1 To .Buttons.Count
            If Not .Buttons(i).Style = tbrSeparator Then
                '.Buttons(i).Image = ImageList1.ListImages(.Buttons(i).Key).Key
            End If
        Next i
    End With
    
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set con = Nothing
End Sub

Private Sub MDIForm_Resize()
    CoolBar1.Move 0, 20, PicBar.Width, PicBar.Height - 100
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
   ' Con.RollbackTrans
    'If Con.State = 1 Then Con.Close
    Set con = Nothing
End Sub

Private Sub mnuAssHRrpt_Click()
    Call ShowAssHRrpt
End Sub

Private Sub mnuAssItems_Click()
    Call ContractorList.showItems
End Sub

Private Sub mnuAuth_Click()
    frmAuthorizeIssRcv.Show
End Sub

Private Sub mnuCalc_Click()

    On Error GoTo err
    Call Shell("calc.exe")
    Exit Sub
err:
    MsgBox "Cannot Run Calculater Due To Error: " & vbNewLine & err.Description, vbInformation

End Sub

Private Sub mnuComp_Click()
    frmComponents.Show 1
End Sub

Private Sub mnuCompCatal_Click()
    frmItems.Show
End Sub

Private Sub mnuCompositions_Click()
    frmCompositions.Show
End Sub

Private Sub mnuCustCatal_Click()
    frmCustCatal.Show
End Sub

Private Sub mnuCustList_Click()
    frmFCustomerList.Show
End Sub

Private Sub mnuCustListrpt_Click()
    Call ShowCustListrpt
End Sub

Private Sub mnuCustOrdList_Click()
    frmFCustOrderList.Show
End Sub

Private Sub mnuCustProformaList_Click()
    frmFProformaList.Show
End Sub

Private Sub mnuDailyRpts_Click()
    frmDailyIssRcvRpt.Show
End Sub

Private Sub mnuDeptEdit_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.mnuEdit(mnuDeptEdit.Tag)
    Else
        Call ContDeptList.mnuEdit(mnuDeptEdit.Tag)
    End If
End Sub

Private Sub mnuDeptInact_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.MakeInactive
    Else
        Call ContDeptList.MakeInactive
    End If
End Sub

Private Sub mnuDeptList_Click()
    DeptList.Show
End Sub

Private Sub mnuDeptnew_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.mnuAddDept
    Else
        Call ContDeptList.mnuAddDept
    End If
End Sub

Private Sub mnuDeptPrint_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.PrintDeptList
    Else
        Call ContDeptList.PrintDeptList
    End If
End Sub

Private Sub mnuDeptRefresh_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.RefreshIt
    Else
        Call ContDeptList.RefreshIt
    End If
End Sub

Private Sub mnuDrilSS_Click()
    Call ShowStockStatus(3)
End Sub

Private Sub mnuEmpEdit_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuEdit(mnuEmpEdit.Tag)
    Else
        Call ContractorList.mnuEdit(mnuEmpEdit.Tag)
    End If
End Sub

Private Sub mnuempNewDept_Click()
    If Me.mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuAddDept
    Else
        Call ContractorList.mnuAddDept
    End If
End Sub
Private Sub mnuEmpNewEmp_Click()
    If Me.mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuAdd
    Else
        Call ContractorList.mnuAdd
    End If
End Sub


Private Sub mnuEmpPrint_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.PrintEmpList
    Else
          Call ContractorList.PrintEmpList
    End If
End Sub

Private Sub mnuempRefresh_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.RefreshIt
    Else
        Call ContractorList.RefreshIt
    End If
End Sub
Private Sub mnuExit_Click()
    If con.State = 1 Then con.Close
    Set con = Nothing
    End
End Sub

Private Sub mnuFDAProcess_Click()
    frmFDAProcesses.Show 1
End Sub

Private Sub mnuFillSS_Click()
    Call ShowStockStatus(4)
End Sub

Private Sub mnuForeignCatal_Click()
    frmForeignCatalog.Show 1
End Sub

Private Sub mnuForgSS_Click()
    Call ShowStockStatus(1)
End Sub

Private Sub mnuFuncSS_Click()
    Call ShowStockStatus(5)
End Sub

Private Sub mnuFuntSS_Click()
    Call ShowStockStatus(9)
End Sub

Private Sub mnuGuage_Click()
    frmGages.Show 1
End Sub

Private Sub mnuHeatSS_Click()
    Call ShowStockStatus(6)
End Sub

Private Sub mnuHICOResources_Click()
    frmHICOR.Show
End Sub

Private Sub mnuHoldSS_Click()
    Call ShowStockStatus(13)
End Sub

Private Sub mnuHolidays_Click()
    Load Holidays
    'Holidays.Move 3000, 500
    With Holidays
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuHRGroups_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("HRGroups", "Hico Resources Groups", "Hico Resource Name")
End Sub

Private Sub mnuHRrpt_Click()
    Call ShowHRrpt
End Sub

Private Sub mnuIssDril_Click()
    Call ShowIssForm(3)
    
End Sub

Private Sub mnuIssFill_Click()
    Call ShowIssForm(4)
End Sub

Private Sub mnuIssForg_Click()
    Call ShowIssForm(1)
End Sub

Private Sub mnuIssFunc_Click()
    Call ShowIssForm(5)
End Sub

Private Sub mnuIssFunt_Click()
    Call ShowIssForm(9)
End Sub

Private Sub mnuIssHeat_Click()
    Call ShowIssForm(6)
End Sub

Private Sub mnuIssHold_Click()
    frmVendRcvItemsHnD.Show
End Sub

Private Sub mnuIssMach_Click()
    Call ShowIssForm(2)
End Sub

Private Sub mnuIssPack_Click()
    Call ShowIssForm(12)
End Sub

Private Sub mnuIsspass_Click()
    Call ShowIssForm(8)
End Sub

Private Sub mnuIssPoli_Click()
    Call ShowIssForm(7)
End Sub

Private Sub mnuIssPoRe_Click()
    Call ShowIssForm(10)
End Sub

Private Sub mnuIssRepair_Click()
    frmVendIssRepair.Show
End Sub

Private Sub mnuIssRepairPoli_Click()
    frmVendIssRepair.Show
End Sub

Private Sub mnuIssRM_Click()
    frmRMIss.Show
End Sub

Private Sub mnuIssRpts_Click()
    Dim f As New frmRptIssRcv
    Call f.ShowForIss
End Sub

Private Sub mnuIssUltr_Click()
    Call ShowIssForm(11)
End Sub

Private Sub mnuItemsFinQuality_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("ItemsFinQuality", "Items Fin. Quality", "Finishing Quality")
End Sub

Private Sub mnuItemTypes_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("ItemTypes", "Item Types", "Item Types")
End Sub

Private Sub mnuItmGrps_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
End Sub

Private Sub mnuLabs_Click()
    frmTestLabs.Show
End Sub

Private Sub mnuLogoff_Click()

    If MsgBox("Are You Sure To Log Off ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Unload MainForm
    LoggingOff = True
    Call Main
    
End Sub

Private Sub mnuLotProc_Click()
    frmInitialLotProcess.Show
End Sub

Private Sub mnuMachSS_Click()
    Call ShowStockStatus(2)
End Sub

Private Sub mnuMake_Click()

    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.MakeInactive
    Else
        Call ContractorList.MakeInactive
    End If
    
End Sub

Private Sub mnuMatList_Click()
    frmMaterial.Show
End Sub

Private Sub mnuNewOrd_Click()
    frmFNewCustOrder.Show
End Sub

Private Sub mnuNewProforma_Click()
    frmFNewProforma.Show
End Sub

Private Sub mnuNewStockOrder_Click()
    frmFNewStockOrder.Show
End Sub

Private Sub mnuOrdStatus_Click()
    frmFCustOrdStatus.Show
End Sub

Private Sub mnuOrdStatus1_Click()
    frmFCustOrdStatus1.Show
End Sub

Private Sub mnuOrdStatusComp_Click()
    frmFCustOrdStatusComp.Show
End Sub

Private Sub mnuPackSS_Click()
    Call ShowStockStatus(12)
End Sub

Private Sub mnuPassSS_Click()
    Call ShowStockStatus(8)
End Sub

Private Sub mnuPassword_Click()
    
    With frmUserInfo
        .Show_Option = Opt_ChangeUserPassword
        .Tag = CurrentUserName
        .Show 1
    End With
    
End Sub

Private Sub mnuPer_Click()
    frmPercentRanges.Show
End Sub

Private Sub mnuPoliSS_Click()
    Call ShowStockStatus(7)
End Sub

Private Sub mnuPOrderList_Click()
    frmRMPOrdList.Show
End Sub

Private Sub mnuPORESS_Click()
    Call ShowStockStatus(10)
End Sub

Private Sub mnuPORpt_Click()
    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .rpt = New rptPOrder
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Venders>", "0"
        .cmbCust.AddVals con, "VendID1", "Venders", "VendID"
        .SelFor = "{VPOrdersrpt.RecieptID}='"
        .JobCondField = " VendID "
        .JobField = " Distinct RecieptID "
        .JobIDField = " RecieptID "
        .JobTable = " VPOrdersrpt "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPorts_Click()
    frmPorts.Show
End Sub

Private Sub mnuProcess_Click()
    frmProcesses.Show 1
End Sub


Private Sub mnuReqForm_Click()
    'frmRequisition.Show
    frmGMPs.Show
End Sub

Private Sub mnuProdPlan_Click()
    frmProdPlan.Show
End Sub

Private Sub mnuProdPlanrpt_Click()

    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .rpt = New rptProdPlan
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Customers>", "0"
        .cmbCust.AddVals con, "CustCode", "ForeignCustomers", "CustCode"
        .SelFor = "{VOrderPlanning.OrderNo}='"
        .JobCondField = " CustCode "
        .JobField = " Distinct OrderNo "
        .JobIDField = " OrderNo "
        .JobTable = " VFItemsQtyToPlan "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPymtTerms_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("PaymentTerms", "Payment Terms", "Payment Terms")
End Sub

Private Sub mnuRcvRepair_Click()
    frmVendRcvRepair.Show
End Sub

Private Sub mnuRcvRepairPoli_Click()
    frmVendRcvRepair.Show
End Sub

Private Sub mnuRcvRepairWastage_Click()
    frmVendRcvRepairWaste.Show
End Sub

Private Sub mnuRcvRpts_Click()
    Dim f As New frmRptIssRcv
    Call f.ShowForIss(True)
End Sub

Private Sub mnuRecRM_Click()
    frmRMRcvItems.Show
End Sub

Private Sub mnuRecSDril_Click()
    Call ShowRcvSForm(3)
End Sub

Private Sub mnuRecSFill_Click()
    Call ShowRcvSForm(4)
End Sub

Private Sub mnuRecSForg_Click()
    Call ShowRcvSForm(1)
End Sub

Private Sub mnuRecSFunc_Click()
    Call ShowRcvSForm(5)
End Sub

Private Sub mnuRecSFunt_Click()
    Call ShowRcvSForm(9)
End Sub

Private Sub mnuRecSHeat_Click()
    Call ShowRcvSForm(6)
End Sub

Private Sub mnuRecSHold_Click()
    Call ShowRcvSForm(13)
End Sub

Private Sub mnuRecSMach_Click()
    Call ShowRcvSForm(2)
End Sub

Private Sub mnuRecSPack_Click()
    Call ShowRcvSForm(12)
End Sub

Private Sub mnuRecSPass_Click()
    Call ShowRcvSForm(8)
End Sub

Private Sub mnuRecSPoli_Click()
    Call ShowRcvSForm(7)
End Sub

Private Sub mnuRecSPoRe_Click()
    Call ShowRcvSForm(10)
End Sub

Private Sub mnuRecSUltr_Click()
    Call ShowRcvSForm(11)
End Sub

Private Sub mnuRecVDril_Click()
    Call ShowRcvVForm(3)
End Sub

Private Sub mnuRecVFill_Click()
    Call ShowRcvVForm(4)
End Sub

Private Sub mnuRecVForg_Click()
    Call ShowRcvVForm(1)
End Sub

Private Sub mnuRecVFunc_Click()
    Call ShowRcvVForm(5)
End Sub

Private Sub mnuRecVFunt_Click()
    Call ShowRcvVForm(9)
End Sub

Private Sub mnuRecVHeat_Click()
    Call ShowRcvVForm(6)
End Sub

Private Sub mnuRecVHold_Click()
    Call ShowRcvVForm(13)
End Sub

Private Sub mnuRecVMach_Click()
    Call ShowRcvVForm(2)
End Sub

Private Sub mnuRecVPack_Click()
    Call ShowRcvVForm(12)
End Sub

Private Sub mnuRecVPass_Click()
    Call ShowRcvVForm(8)
End Sub

Private Sub mnuRecVPoli_Click()
    Call ShowRcvVForm(7)
End Sub

Private Sub mnuRecVPoRe_Click()
    Call ShowRcvVForm(10)
End Sub

Private Sub mnuRecVUltr_Click()
    Call ShowRcvVForm(11)
End Sub

Private Sub mnuRecWastDril_Click()
    Call ShowWastRcvForm(3)
End Sub

Private Sub mnuRecWastFill_Click()
    Call ShowWastRcvForm(4)
End Sub

Private Sub mnuRecWastForg_Click()
    Call ShowWastRcvForm(1)
End Sub

Private Sub mnuRecWastFunc_Click()
    Call ShowWastRcvForm(5)
End Sub

Private Sub mnuRecWastFunt_Click()
    Call ShowWastRcvForm(9)
End Sub

Private Sub mnuRecWastHeat_Click()
    Call ShowWastRcvForm(6)
End Sub

Private Sub mnuRecWastHold_Click()
    Call ShowWastRcvForm(13)
End Sub

Private Sub mnuRecWastMach_Click()
    Call ShowWastRcvForm(2)
End Sub

Private Sub mnuRecWastPack_Click()
    Call ShowWastRcvForm(12)
End Sub

Private Sub mnuRecWastPass_Click()
    Call ShowWastRcvForm(8)
End Sub

Private Sub mnuRecWastPoli_Click()
    Call ShowWastRcvForm(7)
End Sub

Private Sub mnuRecWastPoRe_Click()
    Call ShowWastRcvForm(10)
End Sub

Private Sub mnuRecWastUltr_Click()
    Call ShowWastRcvForm(11)
End Sub

Private Sub mnuReq_Click()
    'frmRequisition.Show
    frmGMPs.Show
End Sub

Private Sub mnuReqRpt_Click()

    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .rpt = New rptRequsition
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Processes>", "0"
        .cmbCust.AddVals con, "Description", "Processes", "ProcessID"
        .SelFor = "{VRequsitionrpt.RecieptID}='"
        .JobCondField = " ProcessID "
        .JobField = " Distinct RecieptID "
        .JobIDField = " RecieptID "
        .JobTable = " VRequsitionrpt "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuRM_Click()
    frmRM.Show
End Sub

Private Sub mnuRMAssItems_Click()
    frmRMVendAssItems.Show
End Sub

Private Sub mnuRMGroups_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("RMGroups", "RM Groups", "RM Name")
End Sub

Private Sub mnuRMListrpt_Click()
    Call ShowRMListrpt
End Sub

Private Sub mnuRMPO_Click()
    frmRMPurchOrd.Show
End Sub

Private Sub mnuRMVendList_Click()
    frmRMVenderList.Show
End Sub

Private Sub mnuRMVendListrpt_Click()
    Call ShowRMVendListrpt
End Sub

Private Sub mnuShowDept_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.ShowInactive(mnuShowDept)
    Else
        Call ContDeptList.ShowInactive(mnuShowDept)
    End If
End Sub

Private Sub mnuShowInactive_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.ShowInactive(mnuShowInactive)
    Else
        Call ContractorList.ShowInactive(mnuShowInactive)
    End If
End Sub

Private Sub mnuSteel_Click()
    frmSteelList.Show
End Sub

Private Sub mnuSteelVends_Click()
    frmSteelVends.Show
End Sub

Private Sub mnuSubGeneral_Click(Index As Integer)

    Select Case Index
        Case 1
            frmInitialLotProcess.Show
        Case 3
            frmPercentRanges.Show
        Case 5
            frmComponents.Show 1
        Case 7
            frmProcesses.Show 1
        Case 8
            frmFDAProcesses.Show 1
        Case 10
            frmSteelVends.Show
        Case 11
            frmTestLabs.Show
        Case 13
            frmUnits.Show
        Case 14
            frmCompositions.Show
        Case 16
            frmPorts.Show
        Case 19
            Call mnuPymtTerms_Click
    End Select
    
End Sub

Private Sub mnuSubOptions_Click(Index As Integer)

    Select Case Index
        Case 1
            Call mnuCalc_Click
        Case 2
            frmUsers.Show 1
        Case 4
            Call mnuPassword_Click
        Case 5
            Call mnuLogoff_Click
        Case 7
            Call mnuExit_Click
    End Select
    
End Sub

Private Sub mnuTaxRanges_Click()
    Load TaxRanges
    'Holidays.Move 3000, 500
    With TaxRanges
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuUltrSS_Click()
    Call ShowStockStatus(11)
End Sub

Private Sub mnuUnits_Click()
    frmUnits.Show
End Sub

Private Sub mnuUserManager_Click()
    frmUsers.Show 1
End Sub

Public Sub ApplySecuritySettings(UserID As Long)
   
    On Error GoTo err
    
    'General Settings...
    mnuSubGeneral(1).Enabled = UserHasAccess("SetInitLotProc", UserID)
    mnuSubGeneral(3).Enabled = UserHasAccess("PercentRanges")
    mnuSubGeneral(5).Enabled = UserHasAccess("Components")
    mnuSubGeneral(7).Enabled = UserHasAccess("Processes")
    mnuSubGeneral(8).Enabled = UserHasAccess("FDAProcesses")
    mnuSubGeneral(10).Enabled = UserHasAccess("SteelVenders")
    mnuSubGeneral(11).Enabled = UserHasAccess("TestLabs")
    mnuSubGeneral(13).Enabled = UserHasAccess("MeasuringUnits")
    mnuSubGeneral(14).Enabled = UserHasAccess("Compositions")
    mnuSubGeneral(16).Enabled = UserHasAccess("CitiesNPorts")
    mnuSubGeneral(18).Enabled = UserHasAccess("PymtTerms")
    
    'Venders
    'mnuVendList.Enabled = UserHasAccess("VenderList")
    'mnuVendAssItems.Enabled = UserHasAccess("VendAssItems")
    mnuVendPymt.Enabled = UserHasAccess("VendAdvPymt")
    mnuVendPaySheet.Enabled = UserHasAccess("VendPymtSheet")
    mnuVendAdvShort.Enabled = UserHasAccess("NewShortTerm")
    mnuVendAdvLong.Enabled = UserHasAccess("NewLongTerm")
    mnuVendShortLoanLedger.Enabled = UserHasAccess("ShortTermLedger")
    mnuVendLongLoanLedger.Enabled = UserHasAccess("LongTermLedger")
    'mnuVendLedger.Enabled = UserHasAccess("VendLedger")
    
    'Raw Material
    mnuRMGroups.Enabled = UserHasAccess("RMGroups")
    mnuRM.Enabled = UserHasAccess("RawMaterial")
    'mnuRMVendList.Enabled = UserHasAccess("RMVendList")
    'mnuRMAssItems.Enabled = UserHasAccess("RMVAssItems")
    mnuRMPO.Enabled = UserHasAccess("RMVNewPO")
    mnuPOrderList.Enabled = UserHasAccess("RMVPOList")
    mnuRecRM.Enabled = UserHasAccess("RMRcvRM")
    mnuIssRM.Enabled = UserHasAccess("RMIssRM")
    
    'HICO Resources
    mnuHRGroups.Enabled = UserHasAccess("HICOResGroups")
    'mnuHICOResources.Enabled = UserHasAccess("HICOResources")
    'mnuVendAssR.Enabled = UserHasAccess("AssHICORes")
        
    'Company
    'mnuCompDeatil.Enabled = UserHasAccess("CompanyDetail")
    'mnuCompCatal.Enabled = UserHasAccess("CompanyCatalog")
    mnuItmGrps.Enabled = UserHasAccess("ItemGroups")
    mnuItemsFinQuality.Enabled = UserHasAccess("ItemFinQuality")
    mnuItemTypes.Enabled = UserHasAccess("ItemTypes")
    'mnuForeignCatal.Enabled = UserHasAccess("ForeignCatalog")
    'mnuCustList.Enabled = UserHasAccess("CustomerList")
    'mnuCustCatal.Enabled = UserHasAccess("CustomerCatalog")
    mnuNewOrd.Enabled = UserHasAccess("CustNewOrder")
    mnuNewStockOrder.Enabled = UserHasAccess("StockNewOrder")
    mnuCustOrdList.Enabled = UserHasAccess("CustOrdList")
    mnuOrdStatus1.Enabled = UserHasAccess("CustOrdStatus")
    mnuNewProforma.Enabled = UserHasAccess("CustNewProforma")
    mnuCustProformaList.Enabled = UserHasAccess("CustProList")
        
    mnuReq.Enabled = UserHasAccess("Requsition")
    mnuProdPlan.Enabled = UserHasAccess("AddProdPlan")
    mnuAuth.Enabled = UserHasAccess("Authorize")
    
    'Reports
    
    'Users
    mnuSubOptions(2).Enabled = UserHasAccess("UserManagement")
    mnuSubOptions(4).Enabled = UserHasAccess("ChangePassword")
    
    
    HicoVisible = UserHasAccess("HicoVisible")
    'If UserID = 1 Then mnuAuth.Enabled = True Else mnuAuth.Enabled = False
    
    If UserID <> 1 Then Call ApplySecurityForIssRcv(UserID)
    '
    'With Toolbar1.Buttons
    '   '.Item("TakeAtt").Enabled = mnutake.Enabled
    '   '.Item("NewShort").Enabled = mnuShort.Enabled
    '   '.Item("NewLong").Enabled = mnuLong.Enabled
    '   '.Item("EmpLedger").Enabled = mnuEmpLedger.Enabled
    '   '.Item("SalSheet").Enabled = mnuSalSheet.Enabled
    'End With
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub ApplySecurityForIssRcv(UserID As Long)
    
    For i = 1 To mnuIssRcvProcs.Count
        mnuIssRcvProcs(i).Enabled = False
    Next i
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select ProcessID From UserAssProcesses Where UserID=" & UserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            mnuIssRcvProcs(![ProcessID]).Enabled = True
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub
Public Sub mnuVendAdvLong_Click()

    With frmVendNewAdv
        .LongTerm = True
        .Show 1
    End With
    
End Sub

Public Sub mnuVendAdvShort_Click()
    
    With frmVendNewAdv
        .LongTerm = False
        .Show 1
    End With
    
End Sub

Private Sub mnuVendAssItems_Click()
    frmVendAssItems.Show
End Sub

Private Sub mnuVendAssR_Click()
    frmVendAssHR1.Show
End Sub

Private Sub mnuVendIssItems_Click()
    frmVendIssItems.Show
End Sub

Private Sub mnuVendIssItemsByItems_Click()
    frmVendIssItemsByItem.Show
End Sub

Private Sub mnuVendLedger_Click()
    frmVendLedger.Show
End Sub

Private Sub mnuVendList_Click()
    frmVenderList.Show
End Sub

Private Sub mnuVendListrpt_Click()
    Call ShowVendListrpt
End Sub

Private Sub mnuVendLongLoanLedger_Click()
    frmVendLoanLedger.Show
End Sub

Private Sub mnuVendPaySheet_Click()
    frmVendSS.Show    'frmVendPaySheet.Show
End Sub

Private Sub mnuVendPymt_Click()
    frmVendAdvSalary.Show
End Sub

Private Sub mnuVendRcvItems_Click()
    frmVendRcvItems.Show
End Sub

Private Sub mnuVendRcvItemsFromSupervisor_Click()
    frmVendRcvItems1.Show
End Sub

Private Sub mnuVendRcvWst_Click()
    frmVendRcvWastage.Show
End Sub

Private Sub mnuVendShortLoanLedger_Click()
    frmVendShortAdvLedger.Show
End Sub

'-----------------------------------------------------------
'Code Of Pop Menu Ports.
Private Sub NewCity_Click()
    Call frmPorts.NewCity_Click
End Sub


Private Sub Rename_Click()
    Call frmPorts.Rename_Click
End Sub

Private Sub del_Click()
    Call frmPorts.del_Click
End Sub

Private Sub PicBusPart_Resize()

    Dim Temp As Long
    PicBusPart.Left = 0
    PicBusPart.Width = 1950
    
    'If PicBusPart.Height = 1995 Then 'Means Its Open
    Temp = Round(PicBusPart.Height / 350)
     'MsgBox "Total Buttons=" & Round(PicBusPart.Height / 350)
    'End If
    
    VSBusPart.Move PicBusPart.Width - VSBusPart.Width - 50, VSBusPart.Top, VSBusPart.Width, PicBusPart.Height - 15
    VSBusPart.Max = (BusPartToolBar.Buttons.Count - Temp) + 1
    'VSBusPart.Left = PicBusPart.Width - VSBusPart.Width - 50
    
    'PicBusPart.Height = 1665
End Sub





Private Sub PicBar_Paint()
    For i = 1 To ToolBar.Count
        ToolBar(i).BorderStyle = ccNone
    Next i
End Sub

Private Sub PicTitle_Click()
    CoolBar1.Refresh
    ToolBar(1).Refresh
    
End Sub

Private Sub PopSubCust_Click(Index As Integer)
    Select Case Index
        Case 0
            frmFCustomer.Show 1
        Case 1
            frmFCustomerList.Show
        Case 3
            frmCustCatal.Show
        Case 4
    End Select
End Sub

Private Sub PopSubProdVend_Click(Index As Integer)
    
    On Error GoTo err
    
    Select Case Index
        Case 0
            Dim f As New frmVender
            With f
                .Add = True
                .Show 1
            End With
        Case 1
            frmVenderList.Show
        Case 3
            frmVendAssItems.Show
        Case 4
            frmVendLedger.Show
            'Call showform(frmVendLedger)
    End Select
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub PopSubRMVend_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim f As New frmRMVender
            With f
                .Add = True
                .Show 1
            End With
        Case 1
            frmRMVenderList.Show
        Case 3
            frmRMVendAssItems.Show
        Case 4
    End Select
End Sub

Private Sub ProdToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 4 Then
        Me.PopUpMenu mnuProdPhases, 512, PicBar.Width
    End If
End Sub


'-----------------------------------------------------------
Private Sub Timer1_Timer()

    Timer1.Interval = 1000
    Dim SeverTime1 As Date
    SeverTime1 = con.Execute("Select GetDate()").Fields(0)
    ServerTime = Format(SeverTime1, "hh:mm:ss AMPM")
    MainForm.Caption = "Production Package." & Space$(5) & " SERVER TIME : " & Format(SeverTime1, "dddd dd-MMM-yyyy hh:mm:ss AMPM")
    
    If Format(ServerDate, "dd-MMM-yyyy") <> Format(SeverTime1, "dd-MMM-yyyy") Then ServerDate = Format(SeverTime1, "dd-MMM-yyyy")
    
End Sub

Private Sub PicToolBar_Resize(Index As Integer)

    Dim Temp As Long
    PicToolBar(Index).Left = 0
    PicToolBar(Index).Width = 2310
    ToolBar(Index).Move 0, 0, PicToolBar(Index).Width - SBar(Index).Width - 40
    'If PicBusPart.Height = 1995 Then 'Means Its Open
    Temp = Round(PicToolBar(Index).Height / 350)
     'MsgBox "Total Buttons=" & Round(PicBusPart.Height / 350)
    'End If
    
    SBar(Index).Move PicToolBar(Index).Width - SBar(Index).Width - 50, 10, SBar(Index).Width, PicToolBar(Index).Height - 15
    SBar(Index).Max = (ToolBar(Index).Buttons.Count - Temp) + 1
    
End Sub
Private Sub SBar_Change(Index As Integer)
    ToolBar(Index).Top = -((SBar(Index).value - 1) * 350)
End Sub

Private Sub ToolBar_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Index
        Case 1
            Call company(Button)
        Case 2
            'Call Policies(Button)
        Case 3
            Call BusPart(Button)
        Case 4
            'Call Orders(Button)
        Case 7
            Call Production(Button)
    End Select
End Sub
Private Sub company(ByVal Button As MSComctlLib.Button)

    On Error GoTo err
    Dim strRet As String
    Select Case Button.Key
        Case "CompDetail"
        Case "CompCatalog"
            'frmItems.Show
            'Exit Sub
            strRet = ShowPopUpMenu(PicBar.Width, Button.Top + CoolBar1.Top, "&Group List", "-", "&Item List", "Add &New Item")
            Select Case strRet
                Case "&Group List"
                    Dim f As New frmItemGroups
                    Load f
                    Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
                Case "&Item List"
                    frmItems.Show
                    
                Case "Add &New Item"
                    With frmNewItem
                        Load frmNewItem
                        .Show 1
                    End With
            End Select
        Case "ForCatalog"
            frmForeignCatalog.Show 1
        Case "CompDept"
                
        Case "CompResources"
            frmHICOR.Show
            
        Case "Banks"
            strRet = ShowPopUpMenu(PicBar.Width, Button.Top + CoolBar1.Top, "&Bank List", "&Add New Bank", "-", "Add &New Cheque Book")
            
    End Select
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub BusPart(ByVal Button As MSComctlLib.Button)

    On Error GoTo err
    Dim strRet As String
    Select Case Button.Key
        Case "FCustomers"
        strRet = ShowPopUpMenu(PicBar.Width, Button.Top + CoolBar1.Top, "&Customer List", "&Add New Customer", "-", "Customer Ledger", "-", "Customer Mailing Labels")
        Select Case strRet
            Case "&Customer List"
                frmFCustomerList.Show
                
            Case "&Add New Customer"
                frmFCustomer.Show 1
            Case "Customer Ledger"
                
            Case "Customer Mailing Labels"
                
        End Select
        
            
    End Select
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub Production(ByVal Button As MSComctlLib.Button)

    On Error GoTo err
    Dim strRet As String
    Select Case Button.Key
        Case "Requsition"
            frmGMPs.Show

        Case "PPlan"
            frmProdPlan.Show

        Case "Authorise"
            frmAuthorizeIssRcv.Show

        Case "Processes"
            'mnuReq.Visible = False
            'mnuProdPlan.Visible = False
            'mnuAuth.Visible = False
            
            Me.PopUpMenu mnuProdPhases, 512, PicBar.Width
    End Select
        
            
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
