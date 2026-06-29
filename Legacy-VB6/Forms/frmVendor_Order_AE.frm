VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendor_Order_AE 
   ClientHeight    =   9285
   ClientLeft      =   -120
   ClientTop       =   -120
   ClientWidth     =   14100
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   14100
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9225
      Left            =   45
      TabIndex        =   11
      Top             =   30
      Width           =   14055
      Begin VB.TextBox txtOrderRevision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
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
         Height          =   285
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   84
         Top             =   8835
         Width           =   1410
      End
      Begin VB.CommandButton cmdAddRevision 
         Caption         =   "Add Revision"
         Height          =   555
         Left            =   6090
         TabIndex        =   83
         Top             =   8565
         Width           =   810
      End
      Begin ComboList.Usercontrol1 cmbItemSearchOrder 
         Height          =   285
         Left            =   1365
         TabIndex        =   69
         Top             =   7875
         Width           =   6405
         _ExtentX        =   11298
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
         TextForeColor   =   16711680
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   75
         TabIndex        =   19
         Top             =   945
         Width           =   4275
         _ExtentX        =   7541
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
         TextForeColor   =   16711680
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   1455
         Width           =   3720
         _ExtentX        =   6562
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   6105
         Left            =   75
         TabIndex        =   22
         Top             =   2355
         Width           =   13905
         _ExtentX        =   24527
         _ExtentY        =   10769
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Article Detail"
         TabPicture(0)   =   "frmVendor_Order_AE.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblTotWords"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblTotal"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblGoTo(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdAddFromQuotation"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "LV"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "PicProduct"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtEditPrice"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtEditQty"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "FNew"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmdDelete"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "cmdAddItem"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkUrgent"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "cmbFinQualityE"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "FraRemarks"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "cmdUpdate"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "PicDelDate"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "Other Detail"
         TabPicture(1)   =   "frmVendor_Order_AE.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtPackaging"
         Tab(1).Control(1)=   "CD1"
         Tab(1).Control(2)=   "Label4"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Import From Excel"
         TabPicture(2)   =   "frmVendor_Order_AE.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdImportExcel"
         Tab(2).Control(1)=   "cmdAddToOrder"
         Tab(2).Control(2)=   "cmdAddToCustomerCatalog"
         Tab(2).Control(3)=   "cmdPrintNFICC"
         Tab(2).Control(4)=   "LVNICompany"
         Tab(2).Control(5)=   "LVNICustomer"
         Tab(2).Control(6)=   "LVOK"
         Tab(2).Control(7)=   "CDExcel"
         Tab(2).Control(8)=   "cmdAddToCompanyCatalog"
         Tab(2).Control(9)=   "Label1(6)"
         Tab(2).Control(10)=   "Label1(10)"
         Tab(2).Control(11)=   "Label1(11)"
         Tab(2).ControlCount=   12
         Begin VB.CommandButton cmdImportExcel 
            Caption         =   "Import from Excel"
            Height          =   465
            Left            =   -69480
            TabIndex        =   75
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdAddToOrder 
            Caption         =   "Add to Order"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -65775
            TabIndex        =   73
            Top             =   5730
            Width           =   4485
         End
         Begin VB.CommandButton cmdAddToCustomerCatalog 
            Caption         =   "Add to Customer Catalog"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -70290
            TabIndex        =   72
            Top             =   5730
            Width           =   4485
         End
         Begin VB.CommandButton cmdPrintNFICC 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74835
            TabIndex        =   71
            Top             =   5730
            Width           =   4485
         End
         Begin VB.PictureBox PicDelDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   10050
            ScaleHeight     =   585
            ScaleWidth      =   2325
            TabIndex        =   54
            Top             =   3570
            Visible         =   0   'False
            Width           =   2355
            Begin VB.CommandButton cmdDelDTEditCancel 
               Caption         =   "Cancel"
               Height          =   285
               Left            =   1665
               TabIndex        =   59
               Top             =   285
               Width           =   630
            End
            Begin VB.CommandButton cmdDelDTEditOK 
               Caption         =   "OK"
               Height          =   285
               Left            =   1020
               TabIndex        =   58
               Top             =   285
               Width           =   630
            End
            Begin VB.CheckBox chkUpdateAll 
               BackColor       =   &H80000005&
               Caption         =   "Update All"
               Height          =   210
               Left            =   0
               TabIndex        =   57
               Top             =   330
               Width           =   1050
            End
            Begin VB.CheckBox chkDelUrgentEdit 
               BackColor       =   &H80000005&
               Caption         =   "Urgent"
               Height          =   210
               Left            =   1515
               TabIndex        =   56
               Top             =   60
               Width           =   885
            End
            Begin MSComCtl2.DTPicker DTDeliveryItemEdit 
               Height          =   300
               Left            =   -15
               TabIndex        =   55
               Top             =   0
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox        =   -1  'True
               CustomFormat    =   "dd-MMM-yy"
               Format          =   170065923
               CurrentDate     =   37250
            End
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            Height          =   285
            Left            =   8370
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1440
            Width           =   900
         End
         Begin VB.Frame FraRemarks 
            BackColor       =   &H00E7EBEF&
            Caption         =   "Remarks :"
            Height          =   2025
            Left            =   9450
            TabIndex        =   61
            Top             =   3045
            Visible         =   0   'False
            Width           =   4410
            Begin VB.TextBox txtRemarks 
               Appearance      =   0  'Flat
               Height          =   1395
               Left            =   30
               MultiLine       =   -1  'True
               TabIndex        =   64
               Top             =   225
               Width           =   4305
            End
            Begin VB.CommandButton cmdOKPackingDetails 
               Caption         =   "OK"
               Height          =   300
               Left            =   2700
               TabIndex        =   63
               Top             =   1650
               Width           =   795
            End
            Begin VB.CommandButton cmdCancelPackDetails 
               Caption         =   "Cancel"
               Height          =   300
               Left            =   3540
               TabIndex        =   62
               Top             =   1650
               Width           =   795
            End
         End
         Begin VB.ComboBox cmbFinQualityE 
            Height          =   315
            Left            =   7710
            TabIndex        =   53
            Top             =   2055
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chkUrgent 
            Caption         =   "Urgent Delivery"
            Height          =   210
            Left            =   6915
            TabIndex        =   52
            Top             =   1485
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddItem 
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
            Height          =   555
            Left            =   3660
            Picture         =   "frmVendor_Order_AE.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   2490
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txtPackaging 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   2520
            Left            =   -74775
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   825
            Width           =   7260
         End
         Begin VB.CommandButton cmdDelete 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   13335
            Picture         =   "frmVendor_Order_AE.frx":0586
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   1440
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Frame FNew 
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   75
            TabIndex        =   26
            Top             =   375
            Width           =   13785
            Begin VB.TextBox txtWeight 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   11415
               MaxLength       =   150
               TabIndex        =   35
               Top             =   765
               Width           =   975
            End
            Begin VB.TextBox txtRemarksToAdd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   4065
               TabIndex        =   66
               Top             =   765
               Width           =   3795
            End
            Begin VB.TextBox txtStampsItem 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   1920
               TabIndex        =   31
               Top             =   765
               Width           =   2160
            End
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   10500
               MaxLength       =   150
               TabIndex        =   29
               Top             =   765
               Width           =   930
            End
            Begin VB.TextBox txtPrice 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   9585
               MaxLength       =   150
               TabIndex        =   34
               Top             =   765
               Width           =   930
            End
            Begin VB.TextBox txtUnit 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "RptTime"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   9585
               Locked          =   -1  'True
               MaxLength       =   150
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   765
               Visible         =   0   'False
               Width           =   870
            End
            Begin MSComCtl2.DTPicker DTDelivery 
               Height          =   300
               Left            =   7860
               TabIndex        =   32
               Top             =   765
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox        =   -1  'True
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   170065923
               CurrentDate     =   37250
            End
            Begin MSForms.ComboBox cmbItemFinishingType 
               Height          =   285
               Left            =   11460
               TabIndex        =   93
               Top             =   240
               Width           =   2340
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "4128;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbFinQuality 
               Height          =   285
               Left            =   0
               TabIndex        =   30
               Tag             =   "FinQuality"
               Top             =   765
               Width           =   1935
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "3413;503"
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
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   $"frmVendor_Order_AE.frx":0AB8
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
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   51
               Top             =   525
               Width           =   13785
            End
            Begin MSForms.ComboBox cmbDesc 
               Height          =   285
               Left            =   6300
               TabIndex        =   38
               Top             =   240
               Width           =   5160
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "9102;503"
               ListWidth       =   10583
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbItems 
               Height          =   285
               Left            =   3150
               TabIndex        =   37
               Top             =   240
               Width           =   3150
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5556;503"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblAmt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   12375
               TabIndex        =   36
               Top             =   765
               Width           =   1410
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   $"frmVendor_Order_AE.frx":0BA2
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
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Width           =   13785
            End
            Begin MSForms.ComboBox cmbCustItemID 
               Height          =   300
               Left            =   0
               TabIndex        =   28
               Top             =   225
               Width           =   3150
               VariousPropertyBits=   746604575
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5556;529"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.TextBox txtEditQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5175
            TabIndex        =   25
            Top             =   1740
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtEditPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3945
            TabIndex        =   24
            Top             =   1740
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.PictureBox PicProduct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2865
            Left            =   75
            ScaleHeight     =   2835
            ScaleWidth      =   2760
            TabIndex        =   23
            Top             =   1785
            Visible         =   0   'False
            Width           =   2790
            Begin VB.Image ImgProduct 
               Height          =   2835
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2760
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   3720
            Left            =   75
            TabIndex        =   40
            Top             =   1785
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   6562
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
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cust ID"
               Object.Width           =   1931
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   4128
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "P. Unit"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Price"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Amount"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Sort"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Delivery"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Finishing"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Stamps"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Remarks"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   13
               Text            =   "Weight"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Item Finish Type"
               Object.Width           =   2469
            EndProperty
         End
         Begin VB.CommandButton cmdAddFromQuotation 
            Caption         =   "Add Items from Quotation"
            Height          =   315
            Left            =   75
            TabIndex        =   49
            Top             =   1455
            Visible         =   0   'False
            Width           =   2220
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   -71805
            Top             =   4410
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "PDF Files Only|*.pdf"
         End
         Begin MSComctlLib.ListView LVNICompany 
            Height          =   4575
            Left            =   -74835
            TabIndex        =   74
            Top             =   1155
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   8070
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S#"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVNICustomer 
            Height          =   4575
            Left            =   -70290
            TabIndex        =   76
            Top             =   1155
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   8070
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S#"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item Name"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView LVOK 
            Height          =   4575
            Left            =   -65775
            TabIndex        =   77
            Top             =   1155
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   8070
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S#"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Weight"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComDlg.CommonDialog CDExcel 
            Left            =   -72315
            Top             =   225
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "Excel Files Only|*.xls"
         End
         Begin VB.CommandButton cmdAddToCompanyCatalog 
            Caption         =   "Add to Company Catalog"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74835
            TabIndex        =   81
            Top             =   5730
            Visible         =   0   'False
            Width           =   4485
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NOT FOUND IN COMPANY CATALOG"
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
            Height          =   255
            Index           =   6
            Left            =   -74835
            TabIndex        =   80
            Top             =   900
            Width           =   4485
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NOT FOUND IN CUSTOMER CATALOG"
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
            Height          =   255
            Index           =   10
            Left            =   -70290
            TabIndex        =   79
            Top             =   900
            Width           =   4485
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OK TO PROCEED"
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
            Height          =   255
            Index           =   11
            Left            =   -65775
            TabIndex        =   78
            Top             =   900
            Width           =   4485
         End
         Begin VB.Label lblGoTo 
            AutoSize        =   -1  'True
            Caption         =   "Search Item : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   68
            Top             =   5505
            Width           =   1185
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Special Instructions :"
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
            Left            =   -74760
            TabIndex        =   46
            Top             =   585
            Width           =   1770
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   12555
            TabIndex        =   44
            Top             =   5760
            Width           =   1305
         End
         Begin VB.Label lblTotWords 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1890
            TabIndex        =   43
            Top             =   5760
            Width           =   9810
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Amount (In Words) :"
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
            Height          =   270
            Index           =   2
            Left            =   75
            TabIndex        =   42
            Top             =   5760
            Width           =   1830
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Amount :"
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
            Height          =   270
            Index           =   1
            Left            =   11685
            TabIndex        =   41
            Top             =   5760
            Width           =   885
         End
      End
      Begin VB.TextBox txtOrderRcvdvia 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9915
         MaxLength       =   150
         TabIndex        =   21
         Top             =   945
         Width           =   1740
      End
      Begin VB.CheckBox chkShowPicture 
         Caption         =   "Show Picture"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   270
         TabIndex        =   20
         Top             =   8775
         Width           =   1560
      End
      Begin VB.TextBox txtStamps 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6300
         TabIndex        =   6
         Top             =   945
         Width           =   2055
      End
      Begin VB.TextBox txtInternalRefNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8340
         MaxLength       =   150
         TabIndex        =   8
         Top             =   945
         Width           =   1590
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3795
         MaxLength       =   150
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1455
         Width           =   1005
      End
      Begin VB.TextBox txtOrderNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4785
         MaxLength       =   150
         TabIndex        =   2
         Top             =   1455
         Width           =   2340
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
         Left            =   -30
         TabIndex        =   14
         Top             =   0
         Width           =   14040
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Vendor Order"
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
            TabIndex        =   15
            Top             =   165
            Width           =   13950
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "New Vendor Order"
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
            TabIndex        =   16
            Top             =   180
            Width           =   13935
         End
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   7110
         TabIndex        =   3
         Top             =   1455
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   170065923
         CurrentDate     =   37250
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   6975
         TabIndex        =   12
         Top             =   8460
         Width           =   6990
         Begin MSForms.CommandButton cmdSaveNStay 
            Height          =   360
            Left            =   60
            TabIndex        =   60
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Stay"
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   116
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
            Left            =   5220
            TabIndex        =   10
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Left            =   1800
            TabIndex        =   7
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   3510
            TabIndex        =   9
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
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
      Begin MSComCtl2.DTPicker dtRevision 
         Height          =   300
         Left            =   4515
         TabIndex        =   82
         Top             =   8820
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   170065923
         CurrentDate     =   37250
      End
      Begin VB.TextBox txtQuality 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4350
         TabIndex        =   5
         Top             =   945
         Visible         =   0   'False
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker DeliveryDT 
         Height          =   285
         Left            =   11745
         TabIndex        =   87
         Top             =   1965
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   170065923
         CurrentDate     =   37250
      End
      Begin MSForms.ComboBox cmbPartial 
         Height          =   285
         Left            =   7485
         TabIndex        =   92
         Top             =   1980
         Width           =   4245
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7488;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbTrans 
         Height          =   285
         Left            =   3780
         TabIndex        =   91
         Top             =   1980
         Width           =   3705
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6535;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbPayment 
         Height          =   285
         Left            =   75
         TabIndex        =   90
         Top             =   1980
         Width           =   3705
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6535;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPymtTerms 
         Alignment       =   2  'Center
         Caption         =   "This Is Not Visible"
         Height          =   165
         Left            =   1815
         MouseIcon       =   "frmVendor_Order_AE.frx":0C96
         MousePointer    =   99  'Custom
         TabIndex        =   88
         Top             =   1770
         Width           =   1335
      End
      Begin MSForms.ComboBox cmbSMethod 
         Height          =   285
         Left            =   12180
         TabIndex        =   86
         Top             =   1455
         Width           =   1800
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3175;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Order Revision"
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
         Height          =   270
         Index           =   12
         Left            =   3090
         TabIndex        =   85
         Top             =   8580
         Width           =   2940
      End
      Begin MSForms.ComboBox cmbQuality 
         Height          =   285
         Left            =   4350
         TabIndex        =   70
         Top             =   945
         Width           =   1965
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3466;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbOrderType 
         Height          =   285
         Left            =   11640
         TabIndex        =   67
         Top             =   945
         Width           =   2340
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4128;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmVendor_Order_AE.frx":0DE8
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
         Height          =   255
         Index           =   4
         Left            =   75
         TabIndex        =   48
         Top             =   705
         Width           =   13905
      End
      Begin MSForms.ComboBox cmbCatID 
         Height          =   285
         Left            =   0
         TabIndex        =   47
         Top             =   345
         Visible         =   0   'False
         Width           =   1545
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2725;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   18
         Top             =   8505
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkCust 
         Height          =   360
         Left            =   270
         TabIndex        =   17
         Top             =   8760
         Visible         =   0   'False
         Width           =   2685
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "4736;635"
         Value           =   "0"
         Caption         =   "Show Customer Item Codes"
         Accelerator     =   73
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbTradeTerms 
         Height          =   285
         Left            =   8655
         TabIndex        =   4
         Top             =   1455
         Width           =   3525
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6218;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmVendor_Order_AE.frx":0EC2
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
         Height          =   255
         Index           =   7
         Left            =   75
         TabIndex        =   13
         Top             =   1215
         Width           =   13905
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmVendor_Order_AE.frx":0FB2
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
         Height          =   255
         Index           =   3
         Left            =   75
         TabIndex        =   89
         Top             =   1725
         Width           =   13680
      End
   End
End
Attribute VB_Name = "frmVendor_Order_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim strDelList As String
Dim iColNo As Integer
Dim bCopyOrder As Boolean

Public Sub EditOrder(TempOrdNo As String, Optional p_bCopyOrder As Boolean)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    strDelList = ""
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    bCopyOrder = p_bCopyOrder
    
    With rs
        .Open "SELECT * FROM VVendor_Orders_Items WHERE OrderNo='" & strOrderNo & "' ORDER BY SortNo", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCompany.ID = !CompanyRefID
        cmbCust.Text = ![AccTitle] & ""
        cmbCountry = ![Country] & ""
        txtOrderNo = ![OrderNo] & ""
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        
        cmbPayment = ![PaymentTerms] & ""
        cmbTrans = ![TransShipment] & ""
        cmbPartial = ![PartialShipment] & ""
        txtPackaging = ![Packaging] & ""
        DeliveryDT = ![DeliveryDT] & ""
        cmbQuality = !Quality & ""
        txtInternalRefNo = !InternalRefNo & ""
        txtStamps = !StampDT & ""
        cmbOrderType = !OrderType & ""
        txtOrderRcvdvia = !OrderRcvdvia & ""
        
        cmbSMethod = ![ShippingMode] & ""
        
        txtOrderRevision = Val(!OrderRevisionNo & "")
        dtRevision = !OrderRevisionDT

'        If IsNull(!StampDT) = False Then
'            DTStamp = !StampDT
'        End If
        Dim strDelivery As String
        LV.ListItems.Clear
        Do Until .EOF
        
            strDelivery = Format(!DeliveryDTItem, "dd-MMM-yyyy")
            
            Set ITM = LV.ListItems.add(, ![CompItemCode] & strDelivery & !QualityItem & !StampsItem & !Remarks & !Item_Finishing_Type_Text & "", .AbsolutePosition)
            
            ITM.Tag = ![CompITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add , , !ITemID & "" '![CatID] & ""
            ITM.ListSubItems.add , , ![CompITemID] & ""
            ITM.ListSubItems.add(, , ![ItemName] & "").Tag = ![Description] & ""
            
            ITM.ListSubItems.add , , !Unit & ""
            
            ITM.ListSubItems.add(, , Format(![Price] & "", "#0.0000")).Tag = ![CustomPrice] & ""
            ITM.ListSubItems.add(, , Val(![Qty] & "")).Tag = Val(![Qty] & "") - Val(![InvQty] & "")
            
            ITM.ListSubItems.add(, , Val(![Qty] & "") * Val(![Price])).Tag = ![ID] & ""
            
            ITM.ListSubItems.add , , Val(![SortNo] & "")
            '----------------------------------------------------------------
            ITM.ListSubItems.add(, , IIf(strDelivery = "01-Jan-1900", "", strDelivery)).Tag = Val(!DeliveryStatus & "")
            ITM.ListSubItems.add , , !QualityItem & ""
            ITM.ListSubItems.add(, , !StampsItem & "").Tag = !Remarks & "" 'GetSingleStringValue("Remarks", "Vendor_Orders_Items_MoreDetails", " WHERE FOI_RefID=" & Val(!ID & ""))
            ITM.ListSubItems.add , , !Remarks & ""
            ITM.ListSubItems.add , , Val(!Weight & "")
             
            ITM.ListSubItems.add(, , !Item_Finishing_Type_Text & "").Tag = Val(!Item_Finishing_Type & "")
            
            If Val(!VOI_RefID & "") > 0 Then
                ITM.ListSubItems(12).Tag = "Added"
            End If
            ITM.ListSubItems(10).Tag = Val(!Qty & "")
            ITM.ListSubItems(11).Tag = Val(!Price & "")
            '----------------------------------------------------------------
            If bCopyOrder Then
                ITM.ListSubItems(7).Tag = "0"
            End If
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT Count(*) FROM FProformaOrders WHERE OrderEntryID IN(SELECT ID FROM Vendor_Orders_Items WHERE OrderNo='" & strOrderNo & "')", con, adOpenForwardOnly, adLockReadOnly
        If .Fields(0).Value = 0 Then
            cmbCompany.Enabled = True
        Else
            cmbCompany.Enabled = False
        End If
        .Close
    End With
    
    Call GetTotal
     
    cmbItemSearchOrder.ClearVals
    cmbItemSearchOrder.AddVals con_ServerSide, "CompItemCode + ' {' + CAST(ID AS VARCHAR(50)) + '}'", "VVendor_Orders_Items", "ID", "WHERE OrderNo='" & strOrderNo & "' " & strOrderBy
    
    txtQty = 0
'    cmbItems.SetFocus
    If bCopyOrder Then
        txtInternalRefNo = GetInternalRefNo
        txtOrderNo = txtInternalRefNo
        strOrderNo = ""
    End If
    
    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub chkCust_Click()

    If cmbCatID.MatchFound Then
        Call FillItems
        For i = 1 To LV.ListItems.count
            If chkCust Then
                LV.ListItems(i).SubItems(2) = LV.ListItems(i).key
                LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).key
            Else
                LV.ListItems(i).SubItems(2) = LV.ListItems(i).Tag
                LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).Tag
            End If
        Next i
    End If
    
End Sub

Private Sub chkShowPicture_Click()
    If chkShowPicture.Value = vbChecked Then
        Call cmbItems_Change
    Else
        PicProduct.Visible = False
    End If
End Sub

'Private Sub cmbCatID_Change()
'    Call cmbCatID_Click
'End Sub
'
'Private Sub cmbCatID_Click()
'    If cmbCatID.MatchFound = False Then Exit Sub
'    Call FillItems
'End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
     
    Dim rs As New ADODB.Recordset
    'txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    With rs
        '.Open "Select Distinct CatID,CatDesc From VVendor_Catalog Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCatID.Clear
        cmbCatID.AddItem "<All Catagories>"
        Call FillItems
        'Do Until .EOF
       '     cmbCatID.AddItem .Fields(1) & ""
      '      cmbCatID.List(cmbCatID.ListCount - 1, 1) = .Fields(0) & ""
     '       .MoveNext
    '    Loop
    '    .Close
'        .Open "SELECT TradeTerms,PaymentTerms,FinishingQuality,Stamps,RcvdVia FROM ForeignCustomers WHERE CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
''        cmbTradeTerms.Text = !TradeTerms & ""
'        cmbPayment.Text = !PaymentTerms & ""
'        cmbFinQuality = !FinishingQuality & ""
'        txtStampsItem = !Stamps & ""
'        txtOrderRcvdvia = !RcvdVia & ""
        '.Close
    End With
    
    Set rs = Nothing
    If cmbCatID.ListCount > 0 Then cmbCatID.ListIndex = 0
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbCustItemID_Change()
    Call cmbCustItemID_Click
End Sub

Private Sub cmbCustItemID_Click()
    cmbItems.ListIndex = cmbCustItemID.ListIndex
End Sub

Private Sub cmbCustItemID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtPrice.SetFocus
    End If
End Sub

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()
    If cmbDesc.MatchFound Then
        cmbItems.ListIndex = cmbDesc.ListIndex
    End If
End Sub

Private Sub cmbDesc_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbFinQualityE_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If cmbFinQualityE.ListIndex = -1 Then
            MsgBox "Please Select from the List.", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.SubItems(iColNo - 1) = cmbFinQualityE.Text
        cmbFinQualityE.Visible = False
    End If
    
End Sub

Private Sub cmbFinQualityE_LostFocus()
    cmbFinQualityE.Visible = False
End Sub

Private Sub cmbItemFinishingType_Change()
    Call cmbItemFinishingType_Click
End Sub

Private Sub cmbItemFinishingType_Click()

    If cmbItemFinishingType.ListIndex > 0 Then
        Dim strFieldName As String
        If cmbItemFinishingType.ListIndex = 1 Then
            strFieldName = "Temper_Rate"
        ElseIf cmbItemFinishingType.ListIndex = 2 Then
            strFieldName = "First_Inspection_Rate"
        End If
        If strFieldName <> "" Then
            Dim dRate As Double
            dRate = GetSingleDoubleValue(strFieldName, "FCustomerCatalog", " WHERE CustCode='" & cmbCust.ID & "' AND CompItemID='" & cmbItems.Text & "'")
            txtPrice = Round(dRate, 4)
        End If
    End If
    
End Sub

Private Sub cmbItems_Change()

    Call cmbItems_Click
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT FinQuality,StampInstructions,SpecialInstructions FROM VVendor_Catalog WHERE ItemID='" & cmbItems.Text & "' AND Vend_AccNo='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            cmbFinQuality.Text = !FinQuality & ""
            txtStampsItem = !StampInstructions & ""
            txtRemarksToAdd = !SpecialInstructions & ""
        End If
        .Close
    End With
    ''''''''Show Picture
    If chkShowPicture.Value = vbChecked Then
        
        With rs
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![ItemPic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "ItemPic"
            Else
                ImgProduct.Picture = LoadPicture("")
            End If
            .Close
        End With
        PicProduct.Visible = True
    End If
    Set rs = Nothing
End Sub

Private Sub cmbItems_Click()

    If cmbItems.MatchFound Then
        
        cmbCustItemID.ListIndex = cmbItems.ListIndex
        cmbDesc.ListIndex = cmbItems.ListIndex
        txtUnit = cmbItems.List(cmbItems.ListIndex, 1)
'        If cmbQuality.ListIndex <= 0 Then
'        If cmbTradeTerms.ListIndex > 9 Then
'            txtPrice = "0" 'Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
'        Else
'            txtPrice = Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
'        End If
'        Else    'Top Quality
'            txtPrice = Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "") 'Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
'        End If
        Dim dPrice As Double, dWeight As Double
        'dPrice = GetSingleDoubleValue("Price", "Vendor_Orders_Items", " WHERE ID IN(SELECT MAX(ID) FROM Vendor_Orders_Items INNER JOIN Vendor_Orders ON Vendor_Orders.OrderNo=Vendor_Orders_Items.OrderNo WHERE CompItemCode='" & cmbItems.Text & "' AND CustCode='" & cmbCust.ID & "')")
        dWeight = GetSingleDoubleValue("FinishedWeight", "Items", " WHERE ItemID='" & cmbItems.Text & "'")
        'If dPrice = 0 Then
            dPrice = Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2))
        'End If
        txtPrice = dPrice
        txtWeight = dWeight
    Else
        txtUnit = ""
        txtPrice = ""
        txtPrice.Tag = ""
    End If
    
End Sub

Private Sub cmbItems_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtPrice.SetFocus
    End If
End Sub

Private Sub cmbQuality_Change()
    Call cmbQuality_Click
End Sub

Private Sub cmbQuality_Click()
    If cmbItems.MatchFound = False Then Exit Sub
    If cmbQuality.ListIndex <= 0 Then
        txtPrice = Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
    Else    'Top Quality
        txtPrice = Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "") 'Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
    End If
End Sub

Private Sub cmbItemSearchOrder_matched()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If Val(LV.ListItems(i).ListSubItems(7).Tag) = Val(cmbItemSearchOrder.ID) Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
        End If
    Next
End Sub

Private Sub cmbTradeTerms_Change()
    Call cmbTradeTerms_Click
End Sub

Private Sub cmbTradeTerms_Click()

    If cmbTradeTerms.MatchFound Then
        cmbCatID.Locked = False
        cmbItems.Locked = False
        cmbCustItemID.Locked = False
        cmbDesc.Locked = False
        If cmbItems.MatchFound Then
            txtPrice = cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2)
            txtPrice.Tag = Val(txtPrice)
        End If
    Else
        cmbCatID.Locked = True
        cmbItems.Locked = True
        cmbCustItemID.Locked = True
        cmbDesc.Locked = True
    End If
    
End Sub

Private Sub cmdAddFromQuotation_Click()

    Dim f As New frmGetSingleTextValue
    Dim strQuotationNo As String
    Call AddToCombo(f.cmbComments, "QuotationNo", "CustomerQuotations", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'")
    If f.getTextValue(strQuotationNo, True, "QuotationNo", "Quotations") = False Then Exit Sub
        
        
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SNo As Integer
    Dim strKey As String
    With rs
        .Open "SELECT * FROM VCustomerQuotations WHERE QuotationNo=" & Val(strQuotationNo), con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If LV.ListItems.count > 0 Then
                SNo = Val(LV.ListItems(LV.ListItems.count).Text) + 1
            Else
                SNo = 1
            End If
            If IsNull(!DeliveryDT) Then
                strKey = !ITemID & "" & "01-Jan-1900" 'Format(DeliveryDT.Value, "dd-MMM-yy")
            Else
                strKey = !ITemID & "" & Format(!DeliveryDT, "dd-MMM-yy")
            End If
            Set ITM = LV.ListItems.add(, strKey, SNo)
            ITM.Tag = !ItemCode & ""
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add(, , !ItemName & "").Tag = !ItemName & ""
            ITM.ListSubItems.add , , !Unit & ""
            ITM.ListSubItems.add(, , Format(Val(!Price & ""), "#0.0000")).Tag = Val(!Price & "")
            ITM.ListSubItems.add(, , Val(!Qty & "")).Tag = "New"
            ITM.ListSubItems.add , , Val(!Qty & "") * Val(!Price & "")
            ITM.ListSubItems.add , , SNo
            ITM.ListSubItems.add(, , IIf(IsNull(!DeliveryDT), "", Format(!DeliveryDT, "dd-MMM-yy"))).Tag = "1"
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Private Sub cmdAddRevision_Click()

    If MsgBox("Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
        txtOrderRevision = Val(txtOrderRevision) + 1
        dtRevision = Date
    End If

End Sub

Private Sub cmdAddToCustomerCatalog_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    con.Execute "INSERT INTO FCustomerCatalog(CustCode,Country,ItemID,Unit,CompItemID,Description,CustomDescription) SELECT '" & cmbCust.ID & _
     "','" & cmbCountry.Text & "',Items.ItemID,'Pcs',Items.ItemID,TImport.ItemName,TImport.ItemName FROM Items INNER JOIN (SELECT ItemID,ItemName FROM Import_Excel_CustomerOrder WHERE ItemID NOT IN(SELECT CompItemID FROM FCustomerCatalog WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "')) TImport ON Items.ItemID=TImport.ItemID"
     
    Call ShowImportedData
    
End Sub

Private Sub cmdAddToOrder_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim i As Integer, ITM As ListItem
    Dim iMaxSortNo As Integer
    Dim SNo As Integer
    iMaxSortNo = 0
    For i = 1 To LV.ListItems.count
        If Val(LV.ListItems(i).SubItems(8)) > iMaxSortNo Then
            iMaxSortNo = Val(LV.ListItems(i).SubItems(8))
        End If
    Next
     
    If LV.ListItems.count > 0 Then
        SNo = Val(LV.ListItems(LV.ListItems.count).Text) + 1
    Else
        SNo = 1
    End If
    
    Dim dPrice As Double
    Dim iDeliveryStatus As Integer, strDelivery As String
    Dim strItemID As String, lQty As Long, strItemName As String, strCustItemID As String
    Dim strFinQuality As String, strStamps As String, strRemarks As String
    Dim dWeight As Double
    For i = 1 To LVOK.ListItems.count
        strItemID = LVOK.ListItems(i).SubItems(1)
        lQty = Val(LVOK.ListItems(i).SubItems(2))
        strItemName = LVOK.ListItems(i).ListSubItems(1).Tag
        strCustItemID = LVOK.ListItems(i).ListSubItems(2).Tag
        strKey = strCustItemID & "01-Jan-1900" & strFinQuality & strStamps & strRemarks
        dPrice = Val(LVOK.ListItems(i).SubItems(3))
        dWeight = Val(LVOK.ListItems(i).SubItems(4))
        For j = 1 To LV.ListItems.count
            If LV.ListItems(j).key = strKey Then
                'MsgBox "Item Already Added.", vbInformation
                'Exit Sub
                'GoTo SKIPITEM
                strKey = strCustItemID & "01-Jan-1900" & strFinQuality & strStamps & j
            End If
        Next
        Set ITM = LV.ListItems.add(, strKey, SNo)
        ITM.Tag = strItemID
    
        ITM.ListSubItems.add , , strCustItemID
        ITM.ListSubItems.add , , strItemID
        ITM.ListSubItems.add(, , strItemName).Tag = cmbCustItemID
    
        ITM.ListSubItems.add , , "Pcs"
        
        ITM.ListSubItems.add(, , Format(dPrice, "#0.0000")).Tag = dPrice
        ITM.ListSubItems.add(, , lQty).Tag = "New"
        ITM.ListSubItems.add , , dPrice * lQty
        ITM.ListSubItems.add , , iMaxSortNo + 1
    
'        If chkUrgent.Value = vbChecked Then
'            iDeliveryStatus = 2         'Urgent
'        ElseIf IsNull(DTDelivery) Then
            iDeliveryStatus = 0         'No Date
            strDelivery = ""
'        Else
'            iDeliveryStatus = 1         'Date Given
'        End If
    
        ITM.ListSubItems.add(, , strDelivery).Tag = iDeliveryStatus
        
        ITM.ListSubItems.add , , ""
        ITM.ListSubItems.add , , ""
    
        ITM.ListSubItems.add , , ""
        ITM.ListSubItems.add(, , dWeight).Tag = "0"
        
        
        ITM.ForeColor = LVOK.ListItems(i).ForeColor
        For j = 1 To ITM.ListSubItems.count
            ITM.ListSubItems(j).ForeColor = LVOK.ListItems(i).ForeColor
        Next
        
        SNo = SNo + 1
        iMaxSortNo = iMaxSortNo + 1
        
SKIPITEM:
    Next

    Call GetTotal
End Sub

Private Sub cmdCancelPackDetails_Click()
    FraRemarks.Visible = False
End Sub

Private Sub cmdDelDTEditCancel_Click()

    PicDelDate.Visible = False
    
End Sub

Private Sub cmdDelDTEditOK_Click()

    Dim iDeliveryStatus As Integer
    
    If chkDelUrgentEdit.Value = vbChecked Then
        iDeliveryStatus = 2         'Urgent
    ElseIf IsNull(DTDeliveryItemEdit) Then
        iDeliveryStatus = 0         'No Date
    Else
        iDeliveryStatus = 1         'Date Given
    End If
    Dim myDT As String
    myDT = LV.SelectedItem.SubItems(iColNo - 1)
    LV.SelectedItem.SubItems(iColNo - 1) = Format(DTDeliveryItemEdit, "dd-MMM-yy")
    LV.SelectedItem.ListSubItems(iColNo - 1).Tag = iDeliveryStatus
    
    If chkUpdateAll.Value = vbChecked Then
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                If myDT = .SubItems(iColNo - 1) Then
                    .SubItems(iColNo - 1) = Format(DTDeliveryItemEdit, "dd-MMM-yy")
                    .ListSubItems(iColNo - 1).Tag = iDeliveryStatus
                End If
            End With
        Next
    End If
    PicDelDate.Visible = False
    
End Sub

Private Sub cmdImportExcel_Click()
    
    If cmbCust.MatchFound = False Then
        MsgBox "Please Select Customer.", vbInformation
        Exit Sub
    End If
    If cmbCountry.MatchEntry = False Then
        MsgBox "Please Select Country.", vbInformation
        Exit Sub
    End If
     
    On Error Resume Next
    Dim strFileName As String
    CDExcel.ShowSave
    
    If Not (err.Number = 32755) Then
        strFileName = CDExcel.FileName
        cmdImportExcel.Tag = strFileName
    Else
        Exit Sub
    End If
    On Error GoTo err
    
    Dim Con_Excel As ADODB.Connection
    Set Con_Excel = New ADODB.Connection
    Con_Excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strFileName & ";" & "Extended Properties=""Excel 8.0;HDR=YES;"""
    
    sTableName = "[sheet1$]"
      
    Call StartTrans(con)
    con.Execute "DELETE FROM Import_Excel_CustomerOrder"
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM " & sTableName, Con_Excel, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            'Fields(0)=ItemID;1=Description;2=Qty;3=Unit;4=Rate;6=Weight;7=Type;8=UDMS
            If Val(.Fields(2) & "") > 0 Then
                con.Execute "INSERT INTO Import_Excel_CustomerOrder(ItemID,ItemName,OrderQty,Unit,Rate,Weight,Steel_Type,UDMS) VALUES('" & (.Fields(0) & "") & _
                 "','" & (.Fields(1) & "") & "'," & Val(.Fields(2) & "") & ",'" & (.Fields(3) & "") & "'," & Val(.Fields(4) & "") & "," & Val(.Fields(6) & "") & _
                 ",'" & (.Fields(7) & "") & "','" & (.Fields(8) & "") & "')"
            End If
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Set Con_Excel = Nothing
    con.CommitTrans
    
    Call ShowImportedData
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub ShowImportedData()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim strItemCodes As String, strCustomerItemCodes As String
    With rs
        'con.Execute "UPDATE Import_Excel_CustomerOrder SET CodeInCompanyCatalog=Items.ItemID FROM Import_Excel_CustomerOrder INNER JOIN Items ON Items.ItemID=Import_Excel_CustomerOrder.ItemID"
        con.Execute "UPDATE Import_Excel_CustomerOrder SET CodeInCompanyCatalog=FCustomerCatalog.CompItemID,CodeInCustomerCatalog=FCustomerCatalog.ItemID FROM Import_Excel_CustomerOrder INNER JOIN FCustomerCatalog ON FCustomerCatalog.ItemID=Import_Excel_CustomerOrder.ItemID WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
        con.Execute "UPDATE Import_Excel_CustomerOrder SET CodeInCompanyCatalog=FCustomerCatalog.CompItemID,CodeInCustomerCatalog=FCustomerCatalog.ItemID,FromOtherCustomer=1 FROM Import_Excel_CustomerOrder INNER JOIN FCustomerCatalog ON FCustomerCatalog.ItemID=Import_Excel_CustomerOrder.ItemID WHERE CodeInCompanyCatalog IS NULL"
        con.Execute "UPDATE Import_Excel_CustomerOrder SET CodeInCompanyCatalog=Items.ItemID FROM Import_Excel_CustomerOrder INNER JOIN Items ON Items.ItemID=Import_Excel_CustomerOrder.ItemID WHERE Import_Excel_CustomerOrder.CodeInCompanyCatalog IS NULL"
        
        'Now Assign to Customer Catalog WHERE FromOtherCustomer=1
        con.Execute "INSERT INTO FCustomerCatalog(CustCode,Country,ItemID,CompItemID,Unit,Description,ExWorks) SELECT '" & cmbCust.ID & "','" & _
         cmbCountry.Text & "',CodeInCustomerCatalog,CodeInCompanyCatalog,'Pcs','',0 FROM Import_Excel_CustomerOrder WHERE FromOtherCustomer=1"
         
        '.Open "SELECT * FROM Import_Excel_CustomerOrder WHERE ItemID NOT IN(SELECT ItemID FROM Items) ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM Import_Excel_CustomerOrder WHERE CodeInCompanyCatalog IS NULL ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVNICompany.ListItems.Clear
        Do Until .EOF
            Set ITM = LVNICompany.ListItems.add(, , LVNICompany.ListItems.count + 1)
            ITM.ListSubItems.add , , !ITemID & ""
            strItemCodes = strItemCodes & "'" & !ITemID & "',"
            ITM.ListSubItems.add , , Val(!OrderQty & "")
            .MoveNEXT
        Loop
        .Close
        
        'Items Found in Company Catalog But not in Customer Catalog...
        '.Open "SELECT * FROM Import_Excel_CustomerOrder WHERE ItemID IN(SELECT ItemID FROM Items) AND ItemID NOT IN(SELECT CompItemID FROM FCustomerCatalog WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "') ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM Import_Excel_CustomerOrder WHERE CodeInCompanyCatalog IS NOT NULL AND CodeInCustomerCatalog IS NULL ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVNICustomer.ListItems.Clear
        Do Until .EOF
            Set ITM = LVNICustomer.ListItems.add(, , LVNICustomer.ListItems.count + 1)
            ITM.ListSubItems.add , , !ITemID & ""
            strItemCodes = strItemCodes & "'" & !ITemID & "',"
            ITM.ListSubItems.add , , Val(!OrderQty & "")
            ITM.ListSubItems.add , , !ItemName & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT Import_Excel_CustomerOrder.*,FCustomerCatalog.ItemID AS CustItemID,FCustomerCatalog.ExWorks FROM Import_Excel_CustomerOrder INNER JOIN FCustomerCatalog ON Import_Excel_CustomerOrder.CodeInCompanyCatalog=FCustomerCatalog.CompItemID WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' ORDER BY Import_Excel_CustomerOrder.EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVOK.ListItems.Clear
        Do Until .EOF
            Set ITM = LVOK.ListItems.add(, , LVOK.ListItems.count + 1)
            ITM.ListSubItems.add(, , !CodeInCompanyCatalog & "").Tag = !ItemName & ""
            ITM.ListSubItems.add(, , Val(!OrderQty & "")).Tag = !CustItemID & ""
            ITM.ListSubItems.add , , Val(!Rate & "") 'Val(!ExWorks & "")
            ITM.ListSubItems.add , , Val(!Weight & "")
            .MoveNEXT
        Loop
        .Close
                
        .Open "SELECT Import_Excel_CustomerOrder.ItemID,COUNT(*) FROM Import_Excel_CustomerOrder INNER JOIN FCustomerCatalog ON Import_Excel_CustomerOrder.ItemID=FCustomerCatalog.CompItemID WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' GROUP BY Import_Excel_CustomerOrder.ItemID HAVING COUNT(*)>1", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            For i = 1 To LVOK.ListItems.count
                If !ITemID & "" = LVOK.ListItems(i).SubItems(1) Then
                    Set ITM = LVOK.ListItems(i)
                    lForeColor = vbRed
                    ITM.ForeColor = lForeColor
                    For j = 1 To ITM.ListSubItems.count
                        ITM.ListSubItems(j).ForeColor = lForeColor
                    Next
                End If
               
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

Private Sub cmdOKPackingDetails_Click()

    
    LV.SelectedItem.SubItems(12) = txtRemarks.Text
    
    FraRemarks.Visible = False
    
End Sub

Private Sub cmdPrintNFICC_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Import_Excel_CustomerOrder_NICompany.rpt")
    Dim strSQLQuery As String
    strSQLQuery = "SELECT * FROM Import_Excel_CustomerOrder WHERE ItemID NOT IN(SELECT ItemID FROM Items)"
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strSQLQuery
    
End Sub

Private Sub cmdSavenStay_Click()
    If Saved Then
        Call EditOrder(txtOrderNo)
    End If
End Sub

Private Sub cmdUpdate_Click()
    
    If Val(txtPrice) <> Val(txtPrice.Tag) Then  'Price Changed
    
        Dim strField As String
        If cmbTradeTerms.ListIndex = 0 Then     'FOB
            strField = "FOB"
        ElseIf cmbTradeTerms.ListIndex = 1 Then     'FOB-Sialkot
            strField = "ExWorks"
        ElseIf cmbTradeTerms.ListIndex = 2 Then     '
            strField = "CnFAir"
        ElseIf cmbTradeTerms.ListIndex = 3 Then     '
            strField = "CnFSea"
        ElseIf cmbTradeTerms.ListIndex = 4 Then     '
            strField = "CIFAir"
        ElseIf cmbTradeTerms.ListIndex = 5 Then     '
            strField = "CIFSea"
        ElseIf cmbTradeTerms.ListIndex = 6 Then     '
            strField = "CnIAir"
        ElseIf cmbTradeTerms.ListIndex = 7 Then     '
            strField = "CnISea"
        End If
        
        con.Execute "UPDATE FCustomerCatalog SET " & strField & "=" & Val(txtPrice) & " WHERE CompItemID='" & cmbItems.Text & "' AND CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
        Dim lEntryID As Long, strDescription As String
        lEntryID = GetSingleLongValue("EntryID", "FCustomerCatalog", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND CompItemID='" & cmbItems.Text & "'")
        strDescription = GetSingleStringValue("Description", "FCustomerCatalog", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND CompItemID='" & cmbItems.Text & "'")
        
        con.Execute "INSERT INTO FCustCatalRevs(RefID,Description,FOB,CIFSea,CIFAir,CnFSea,CnFAir,FOBSialkot,CnISea," & _
         "CnIAir,DT,UserName) SELECT " & lEntryID & ",Description,FOB,CIFSea,CIFAir,CnFSea,CnFAir,ExWorks,CnISea," & _
         "CnIAir,getDate(),'" & CurrentUserName & "' FROM FCustomerCatalog WHERE EntryID=" & lEntryID
                 
        Dim lEntryID1 As Long
        
        lEntryID1 = GetSingleLongValue("MAX(EntryID)", "FCustCatalRevs")
        con.Execute "UPDATE FCustCatalRevs SET " & IIf(strField = "ExWorks", "FOBSialkot", strField) & "=" & Val(txtPrice.Tag) & " WHERE EntryID=" & lEntryID1
        
    End If
    
End Sub


Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    If Not cmbItems.MatchFound Or cmbItems = "" Then
        MsgBox "Invalid Item Selected", vbInformation
        Exit Sub
    'ElseIf Val(txtPrice) = 0 Then
    '    MsgBox "Price Cannot Be Zero", vbInformation
    '    Exit Sub
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity To Order Cannot Be Zero", vbInformation
        Exit Sub
    End If

    Dim ITM As ListItem
    Dim iMaxSortNo As Integer
    iMaxSortNo = 0
    For i = 1 To LV.ListItems.count
        If Val(LV.ListItems(i).SubItems(8)) > iMaxSortNo Then
            iMaxSortNo = Val(LV.ListItems(i).SubItems(8))
        End If
    Next
    
    Dim SNo As Integer
    If LV.ListItems.count > 0 Then
        SNo = Val(LV.ListItems(LV.ListItems.count).Text) + 1
    Else
        SNo = 1
    End If
    
    Dim iDeliveryStatus As Integer
    Dim strKey As String
    If IsNull(DTDelivery.Value) Then
        strKey = cmbCustItemID.Text & "01-Jan-1900" & cmbFinQuality.Text & txtStampsItem & txtRemarksToAdd & cmbItemFinishingType.Text
    Else
        strKey = cmbCustItemID.Text & Format(DTDelivery.Value, "dd-MMM-yy") & cmbFinQuality.Text & txtStampsItem & txtRemarksToAdd & cmbItemFinishingType.Text
    End If
    
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).key = strKey Then
            MsgBox "Item Already Added.", vbInformation
            Exit Sub
        End If
    Next
    Set ITM = LV.ListItems.add(, strKey, SNo)
    ITM.Tag = cmbItems.Text 'cmbItems.List(cmbItems.ListIndex, 2)
    
    ITM.ListSubItems.add , , cmbCustItemID.Text
    ITM.ListSubItems.add , , cmbItems
    ITM.ListSubItems.add(, , cmbDesc).Tag = cmbCustItemID.Text
    
    ITM.ListSubItems.add , , txtUnit
    
    ITM.ListSubItems.add(, , Format(txtPrice, "#0.0000")).Tag = Val(txtPrice.Tag)
    ITM.ListSubItems.add(, , txtQty).Tag = "New"
    ITM.ListSubItems.add , , lblAmt
    ITM.ListSubItems.add , , iMaxSortNo + 1
    
    If chkUrgent.Value = vbChecked Then
        iDeliveryStatus = 2         'Urgent
    ElseIf IsNull(DTDelivery) Then
        iDeliveryStatus = 0         'No Date
    Else
        iDeliveryStatus = 1         'Date Given
    End If
    
    ITM.ListSubItems.add(, , Format(DTDelivery, "dd-MMM-yy")).Tag = iDeliveryStatus
    
    ITM.ListSubItems.add , , cmbFinQuality.Text
    ITM.ListSubItems.add , , txtStampsItem
    
    ITM.ListSubItems.add , , txtRemarksToAdd
    ITM.ListSubItems.add , , Val(txtWeight)
    ITM.ListSubItems.add(, , cmbItemFinishingType.Text).Tag = cmbItemFinishingType.ListIndex
    'cmbItemID.RemoveItem cmbItemID.ListIndex
    'cmbItem.RemoveItem cmbItem.ListIndex
    
    Call GetTotal

    txtQty = 0
    txtWeight = ""
    cmbCustItemID = ""
    'cmbCustItemID.SetFocus
    cmbItems.SetFocus
    txtRemarksToAdd = ""
    
    cmbItemFinishingType.ListIndex = 0
    
    ITM.EnsureVisible
    ITM.Selected = True
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LV.SelectedItem.ListSubItems(6).Tag <> "New" Then
        'If Val(LV.SelectedItem.ListSubItems(6).Tag) <> Val(LV.SelectedItem.ListSubItems(6)) Then
'        If Val(LV.SelectedItem.ListSubItems(6).Tag) <> 0 Then
'            MsgBox "Can't Remove This Item.This Item Is In Use In Some Proforma.", vbInformation
'            Exit Sub
'        End If
    End If
    
    If Val(LV.SelectedItem.ListSubItems(7).Tag) <> 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(7).Tag) & ","
    End If
    
    LV.ListItems.Remove LV.SelectedItem.Index
    
End Sub

Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
       
    On Error GoTo err
    strOrderNo = ""
    Call cmbCust.AddVals(con, "AccTitle + ' {'+AccNo+'}'", "VVenders", "AccNo", " WHERE Active=1")
    OrderDT = ServerDate
    DeliveryDT = DateAdd("m", 2, ServerDate)
    DTDelivery.Value = DateAdd("m", 2, ServerDate)
     
    DTStamp = ServerDate
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    Call FillCmbs
    
    With lblPymtTerms
        .Caption = ""
        .BackStyle = 0
    End With
    
    cmbItemSearchOrder.ListHeight = 2000
    
    SSTab1.Tab = 0
        
    dtRevision = Date
    bCopyOrder = False
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Public Sub AddNew(Optional bSample As Boolean = False)

'    If bSample Then
'        cmbOrderType.ListIndex = 1
'        cmbOrderType.Enabled = False
'    End If
    
    txtInternalRefNo = "" 'GetInternalRefNo
    txtOrderNo = "" 'txtInternalRefNo
    
    Me.Show
    
End Sub

Private Function GetInternalRefNo() As String
 
    On Error GoTo err
    
    Dim lRefNo As Long, strInternalRefNo As String
    If strCompany = "IAA" Or strCompany = "Kami" Then
        lRefNo = GetSingleLongValue("MAX(CAST(SUBSTRING(InternalRefNo,5,LEN(InternalRefNo)-4) AS INT))", "Vendor_Orders", " WHERE LEFT(InternalRefNo,4)='MED/'")
        lRefNo = lRefNo + 1
    Else
        lRefNo = GetSingleLongValue("MAX(CAST(InternalRefNo AS INT))", "Vendor_Orders")
        If lRefNo = 0 Then
            lRefNo = Val(InputBox("Please Enter the Starting Number for Internal Ref. #", "Internal Reference No.", "596"))
        Else
            lRefNo = lRefNo + 1
        End If
    End If
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        strInternalRefNo = "MED/" & Format(lRefNo, "0000")
    Else
        strInternalRefNo = lRefNo
    End If
    GetInternalRefNo = strInternalRefNo
    
    Exit Function
err:
    MsgBox err.Description

    
End Function

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub lblPymtTerms_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("PaymentTerms", "Payment Terms", "Payment Terms")
End Sub

Private Sub LV_Click()
    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_DblClick()

    If iColNo = 6 Or iColNo = 14 Then
'        Dim strInvoiceNo As String
'        strInvoiceNo = UsedInPInvoice()
'        If strInvoiceNo <> "" Then
'            MsgBox "Article has been used in " & strInvoiceNo & ".Can't Edit Price Now.'", vbInformation
'            Exit Sub
'        End If
        'Editing of Price
        With txtEditPrice
            .Alignment = 2
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNo = 9 Then
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(9).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(9).Width
            .Text = LV.SelectedItem.SubItems(8)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNo = 10 Then
        With PicDelDate
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top
            '.Value = LV.SelectedItem.SubItems(iColNo - 1)
            Dim iDelStatus As Integer
            iDelStatus = Val(LV.SelectedItem.ListSubItems(iColNo - 1).Tag)
            If iDelStatus = 0 Then      'No Delivery...
                DTDeliveryItemEdit.Value = Null
                chkDelUrgentEdit.Value = vbUnchecked
            ElseIf iDelStatus = 1 Then      'Delivery Date
                DTDeliveryItemEdit.Value = LV.SelectedItem.SubItems(iColNo - 1)
                chkDelUrgentEdit.Value = vbUnchecked
            ElseIf iDelStatus = 2 Then      'Urgent
                DTDeliveryItemEdit.Value = LV.SelectedItem.SubItems(iColNo - 1)
                chkDelUrgentEdit.Value = vbChecked
            End If
            
            .Visible = True
            DTDeliveryItemEdit.SetFocus
            
        End With
    ElseIf iColNo = 11 Then
    
        With cmbFinQualityE
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top
            If LV.SelectedItem.SubItems(iColNo - 1) <> "" Then
                For i = 0 To .ListCount - 1
                    If .List(i) = LV.SelectedItem.SubItems(iColNo - 1) Then
                        .ListIndex = i
                        Exit For
                    End If
                Next
            End If
            .Visible = True
            .SetFocus
        End With
        
    ElseIf iColNo = 12 Then
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    Else
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(7).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(7).Width
            .Text = LV.SelectedItem.SubItems(6)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    End If
    
End Sub

Private Function UsedInPInvoice() As String

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim lRefID As Long
    lRefID = Val(LV.SelectedItem.ListSubItems(7).Tag)
    With rs
        .Open "Select Top 1 CustomInvoice From CustomInvoiceItems Where RefID In(Select EntryID From FProformaOrders Where OrderEntryID=" & lRefID & ")", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            UsedInPInvoice = ""
        Else
            UsedInPInvoice = .Fields(0)
        End If
        .Close
    End With
    
    Exit Function
    
err:
    MsgBox err.Description
    
End Function

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    With cmdDelete
'        .Move .Left, LV.Top + Item.Top
'        .Visible = True
'    End With
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then 'Delete
        Call cmdDelete_Click
    End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > LV.ColumnHeaders(i).Left And X < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
     
    Dim strValue As String
    strValue = LV.SelectedItem.ListSubItems(12).Tag
    
    If strValue <> "" And strValue <> "Remove" Then
        strRet = ShowPopUpMenu1("Add Remarks", "View PDF", "Remove PDF")
    Else
        strRet = ShowPopUpMenu1("Add Remarks", "Attach PDF")
    End If
    
    If strRet = "Add Remarks" Then
    
        With txtRemarks
            .Text = LV.SelectedItem.SubItems(12)
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
        With FraRemarks
            .Caption = "Remarks of " & LV.SelectedItem.SubItems(2) & ":"
            If LV.SelectedItem.Top + LV.Top > 3045 Then
                .Move FraRemarks.Left, 3045
            Else
                .Move FraRemarks.Left, LV.SelectedItem.Top + LV.Top
            End If
            .Visible = True
        End With
        txtRemarks.SetFocus
    ElseIf strRet = "Attach PDF" Then
        Call BrowsePDF
    ElseIf strRet = "Remove PDF" Then
        LV.SelectedItem.ListSubItems(12).Tag = "Remove"
    ElseIf strRet = "View PDF" Then
        If LV.SelectedItem.ListSubItems(12).Tag = "Added" Then
            Call ShowTechDrawing("Vendor_Orders_Items_PDF", "PDF", " WHERE FOI_RefID=" & Val(LV.SelectedItem.ListSubItems(7).Tag), "Sample Attached", Me.CD1)
        Else
            Call OpenPDFFile(LV.SelectedItem.ListSubItems(12).Tag, Me.CD1)
        End If
        
    End If
    
End Sub

Private Sub BrowsePDF()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        LV.SelectedItem.ListSubItems(12).Tag = Picfile
    End If
       
End Sub

Private Sub OrderDT_Change()
    If strCompany = "IAA" Or strCompany = "Kami" Then
        txtInternalRefNo = GetInternalRefNo
    End If
    DeliveryDT.Value = DateAdd("m", 2, OrderDT.Value)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        cmbItemSearchOrder.Visible = True
    Else
        cmbItemSearchOrder.Visible = False
    End If
End Sub


Private Sub txtEditPrice_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditPrice.Visible = False
    ElseIf KeyAscii = 13 Then
        If iColNo = 14 Then
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEditPrice)
            txtEditPrice.Visible = False
        Else
            LV.SelectedItem.SubItems(5) = Val(txtEditPrice)
            LV.SelectedItem.SubItems(7) = Val(txtEditPrice) * Val(LV.SelectedItem.SubItems(6))
            If LV.SelectedItem.Index < LV.ListItems.count Then
                LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                Call LV_DblClick
            Else
                txtEditPrice.Visible = False
            End If
        End If
    End If

End Sub

Private Sub txtEditQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditQty.Visible = False
    ElseIf KeyAscii = 13 Then
        If iColNo = 9 Then
            If Val(txtEditQty) <= 0 Then
                MsgBox "Invalid Sort No.", vbInformation
                txtEditQty.SetFocus
                Exit Sub
            End If
            LV.SelectedItem.SubItems(8) = Val(txtEditQty)
            If LV.SelectedItem.Index < LV.ListItems.count Then
                Dim iIndex As Integer
                iIndex = LV.SelectedItem.Index + 1
                LV.ListItems(iIndex).Selected = True
                With txtEditQty
                    .Top = LV.SelectedItem.Top + LV.Top
                    .Text = LV.SelectedItem.SubItems(8)
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
            Else
                txtEditQty.Visible = False
            End If
        ElseIf iColNo = 12 Then
            LV.SelectedItem.SubItems(iColNo - 1) = txtEditQty
            txtEditQty.Visible = False
            'txtEditQty.Visible = False
        Else
'            If Val(txtEditQty) <= 0 Or Val(txtEditQty) < Val(LV.SelectedItem.ListSubItems(6).Tag) Then
'                MsgBox "Invalid Qty.", vbInformation
'                txtEditQty.SetFocus
'                Exit Sub
'            End If
            LV.SelectedItem.SubItems(6) = Val(txtEditQty)
            LV.SelectedItem.SubItems(7) = Val(txtEditQty) * Val(LV.SelectedItem.SubItems(5))
            If LV.SelectedItem.Index < LV.ListItems.count Then
                LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                Call LV_DblClick
            Else
                txtEditQty.Visible = False
            End If
            
        End If
    End If
    
End Sub

Private Sub txtPrice_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtPrice_GotFocus()
'    txtQty.TabIndex = 41
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtQty.SetFocus
    End If
    KeyAscii = OnlyNumber(KeyAscii, True)
   
End Sub

Private Sub txtPrice_LostFocus()
'    txtQty.TabIndex = 36
End Sub

Private Sub txtQty_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtQty_GotFocus()
    cmdAddItem.Default = True
    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtQty_LostFocus()
    cmdAddItem.Default = False
'    txtQty.TabIndex = 41
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)

   'KeyAscii = OnlyNumber(KeyAscii)
    If KeyAscii = 13 Then
        Call cmdAddItem_Click
    End If
    
End Sub


Private Sub GetTotal()

    Dim i As Integer
    Dim total As Double
    total = 0

    For i = 1 To LV.ListItems.count
        total = total + Val(LV.ListItems(i).ListSubItems(7).Text)
    Next i

    lblTotal = Format(total, "#0.00")
    lblTotWords = ConvertInWords(lblTotal, txtCurrency.Text)
    
End Sub

Private Function Saved() As Boolean
            
    On Error GoTo err
    
'    If strCompany = "Dr-Frgz" And cmbOrderType.Text = "Sample Order" Then   'Create Auto Proforma...
'        bAutoProforma = True
'    Else
'        bAutoProforma = False
'    End If
    bAutoProforma = True
    Saved = False
    If cmbCompany.MatchFound = False Then
        MsgBox "Please Select Company Name.", vbInformation
        Exit Function
    ElseIf txtOrderNo = "" Then
        MsgBox "Please Enter Order No.", vbInformation
        txtOrderNo.SetFocus
        Exit Function
    End If
    
    If LV.ListItems.count = 0 Then
        MsgBox "No Items Added to This Order", vbInformation
        Exit Function
    End If
         
    Dim myOrderDeliveryDT As Date
    If IsNull(DeliveryDT.Value) Then
        myOrderDeliveryDT = #1/1/1900#
    Else
        myOrderDeliveryDT = DeliveryDT.Value
    End If
    
    Dim lEntryIDForMoreDetails As String
    Call StartTrans(con)
    '----------------------------------------------------------------
    'If strOrderNo <> "" Then
    '    con.Execute "Delete from Vendor_Orders Where OrderNo='" & _
         strOrderNo & "'"
    '     'Now Deleting is giving a problem...
    'End If
    
    'Save Order
    
    If Not cmbCust.Text = "HICO" Then
        If strOrderNo <> "" Then
            con.Execute "UPDATE Vendor_Orders SET OrderNo='" & txtOrderNo & "',DT='" & _
             OrderDT & "',TradeTerms='" & cmbTradeTerms & "',PartialShipment='" & _
             cmbPartial & "',PaymentTerms='" & cmbPayment & "',TransShipment='" & _
             cmbTrans & "',Packaging='" & txtPackaging & "',DeliveryDT='" & _
             myOrderDeliveryDT & "',CompanyRefID=" & cmbCompany.ID & ",StampDT='" & txtStamps & _
             "',Quality='" & cmbQuality & "',InternalRefNo='" & txtInternalRefNo & _
             "',OrderRcvdVia='" & txtOrderRcvdvia & "',OrderType='" & cmbOrderType.Text & "',OrderRevisionNo=" & Val(txtOrderRevision) & ",OrderRevisionDT='" & dtRevision & "',ShippingMode='" & cmbSMethod.Text & "' WHERE OrderNo='" & strOrderNo & "'"
             
             If strDelList <> "" Then
                strDelList = Left(strDelList, Len(strDelList) - 1)
                con.Execute "DELETE FROM Vendor_Orders_Items WHERE ID IN(" & strDelList & ")"
            End If
            
        Else
        
            con.Execute "INSERT INTO Vendor_Orders(Vend_AccNo," & _
             "OrderNo,DT,TradeTerms,PartialShipment,PaymentTerms," & _
             "TransShipment,Packaging,DeliveryDT,CompanyRefID,StampDT,Quality,InternalRefNo,OrderRcvdVia,OrderType,OrderRevisionNo,OrderRevisionDT,ShippingMode) VALUES('" & _
             cmbCust.ID & "','" & txtOrderNo & "','" & OrderDT & "','" & cmbTradeTerms & "','" & cmbPartial & _
             "','" & cmbPayment & "','" & cmbTrans & "','" & txtPackaging & "','" & myOrderDeliveryDT & "'," & cmbCompany.ID & ",'" & txtStamps & "','" & _
             cmbQuality & "','" & txtInternalRefNo & "','" & txtOrderRcvdvia & "','" & cmbOrderType.Text & "'," & Val(txtOrderRevision) & ",'" & dtRevision & "','" & cmbSMethod.Text & "')"
           
        End If
    End If
    
    'Now Save Items
    Dim strItemCode As String, strDelivery As String
    Dim lFOI_RefID As Long
    Dim iRecAffect_CC As Integer
    Dim iItemEdited As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            'strItemCode = Left(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 1)
            strItemCode = .SubItems(1)
'SaveNewItem:
            If .ListSubItems(9).Tag = 0 Then
                strDelivery = #1/1/1900#
            Else
                strDelivery = .ListSubItems(9)
            End If
            
            If Val(.ListSubItems(7).Tag) = 0 Then
                
                con.Execute "INSERT INTO Vendor_Orders_Items(OrderNo,ItemID_Vendor,ItemID," & _
                 "Price,Qty,InvQty,CustomPrice,SortNo,DeliveryDT,Quality,Stamps,DeliveryStatus,Remarks,Weight,Item_Finishing_Type) VALUES('" & txtOrderNo & "','" & _
                 strItemCode & "','" & .Tag & "'," & Val(.SubItems(5)) & "," & Val(.SubItems(6)) & "," & _
                 Val(.SubItems(6)) & "," & Val(.ListSubItems(5).Tag) & "," & Val(.SubItems(8)) & ",'" & _
                 strDelivery & "','" & .SubItems(10) & "','" & .SubItems(11) & "'," & Val(.ListSubItems(9).Tag) & ",'" & .SubItems(12) & "'," & Val(.SubItems(13)) & "," & Val(.ListSubItems(14).Tag) & ")"
                 
                'Now Check If Item Present in Customer Catalog...
'                If cmbItemFinishingType.ListIndex = 0 Then
'                    con.Execute "UPDATE FCustomerCatalog SET ExWorks=" & Val(.SubItems(5)) & " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND CompItemID='" & .Tag & "'", iRecAffect_CC
'                    If iRecAffect_CC = 0 Then
'                        con.Execute "INSERT INTO FCustomerCatalog(CustCode,Country,CompItemID,ItemID,Description,ExWorks) VALUES('" & cmbCust.ID & "','" & cmbCountry.Text & _
'                         "','" & .Tag & "','" & .Tag & "','" & .SubItems(3) & "'," & Val(.SubItems(5)) & ")"
'                    End If
'                End If
                If .ListSubItems(12).Tag <> "" Then
                    lFOI_RefID = GetSingleLongValue("MAX(ID)", "Vendor_Orders_Items")
                    If .ListSubItems(12).Tag = "Remove" Then
                        con.Execute "DELETE FROM Vendor_Orders_Items_PDF WHERE FOI_RefID=" & lFOI_RefID
                    ElseIf .ListSubItems(12).Tag = "Added" Then
                        'Do Nothing....
                    Else
                        Call AttachPDF(lFOI_RefID, .ListSubItems(12).Tag)
                    End If
                End If
                       
'                If bAutoProforma Then
'
'                    lProformaRefID = GetSingleLongValue("MAX(ID)", "Vendor_Orders_Items")
'                    con.Execute "INSERT INTO FProformaOrders(PInvoice,OrderEntryID,Qty,InvQty,Price,CustomPrice,ItemDescription) VALUES('" & _
'                     txtOrderNo & "'," & lProformaRefID & "," & Val(LV.ListItems(i).SubItems(6)) & "," & _
'                     Val(LV.ListItems(i).SubItems(6)) & "," & Val(LV.ListItems(i).SubItems(5)) & "," & _
'                     Val(LV.ListItems(i).SubItems(5)) & ",'" & LV.ListItems(i).SubItems(3) & "')"
'
'                    con.Execute "UPDATE Vendor_Orders_Items SET InvQty=0 WHERE ID=" & lProformaRefID
'                End If
                
            Else
                
                con.Execute "UPDATE Vendor_Orders_Items SET Qty=" & Val(.SubItems(6)) & ",InvQty=" & Val(.SubItems(6)) - Val(.ListSubItems(6).Tag) & _
                 ",Price=" & Val(.SubItems(5)) & ",CustomPrice=" & Val(.SubItems(5)) & ",SortNo=" & Val(.SubItems(8)) & ",DeliveryDT='" & strDelivery & _
                 "',Quality='" & .SubItems(10) & "',Stamps='" & .SubItems(11) & "',DeliveryStatus=" & Val(.ListSubItems(9).Tag) & _
                 ",Remarks='" & .SubItems(12) & "',Weight=" & Val(.SubItems(13)) & ",Item_Finishing_Type=" & _
                 Val(.ListSubItems(14).Tag) & " WHERE ID=" & Val(.ListSubItems(7).Tag)
                
                If Val(.SubItems(6)) <> Val(.ListSubItems(10).Tag) Or Val(.SubItems(5)) <> Val(.ListSubItems(11).Tag) Then
                    con.Execute "UPDATE Vendor_Orders_Items SET Item_Edited=1 WHERE ID=" & Val(.ListSubItems(7).Tag)
                End If
                
                If .ListSubItems(12).Tag <> "" Then
                    lFOI_RefID = Val(.ListSubItems(7).Tag)
                    If .ListSubItems(12).Tag = "Remove" Then
                        con.Execute "DELETE FROM Vendor_Orders_Items_PDF WHERE FOI_RefID=" & lFOI_RefID
                    ElseIf .ListSubItems(12).Tag = "Added" Then
                        'Do Nothing....
                    Else
                        Call AttachPDF(lFOI_RefID, .ListSubItems(12).Tag)
                    End If
                End If
                
'                If bAutoProforma Then
'                    con.Execute "UPDATE FProformaOrders SET Qty=" & Val(LV.ListItems(i).SubItems(6)) & _
'                    ",InvQty=InvQty-(Qty-" & Val(LV.ListItems(i).SubItems(6)) & ") WHERE OrderEntryID=" & _
'                    Val(LV.ListItems(i).ListSubItems(7).Tag)
'                End If
                
            End If
            
        End With
    Next
      
    con.CommitTrans
    
    Saved = True
    If chkPrint.Value = True Then
        Call PrintOrder(txtOrderNo)
    End If
    
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Function IsOrderQtyChanges(p_strOrderNo As String) As Boolean

    Dim rs As New ADODB.Recordset
    Dim strItemCode As String
    Dim myCollection As New Collection
    Dim myObj As clsOrderDetails
    Dim bOrderUpdated As Boolean
    bOrderUpdated = False
    With rs
        .Open "SELECT * FROM Vendor_Orders_Items WHERE OrderNo='" & p_strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set myObj = New clsOrderDetails
            myObj.strItemCode = !CompItemCode & ""
            myObj.lQty = Val(!Qty & "")
            myObj.dPrice = Val(!Price & "")
            
            myCollection.add myObj, !CompItemCode & "" & !DeliveryDT & !Quality & !Stamps
            Set myObj = Nothing
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    'Now Compare this Collection with Now Present Items & Their Qty.
    Dim i As Integer, j As Integer, bFound As Boolean
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            'Find it in the Collection & Compare Qty...
            bFound = False
            For j = 1 To myCollection.count
                If .SubItems(2) = myCollection.Item(j).strItemCode Then
                    bFound = True
                    Set myObj = myCollection.Item(.SubItems(2))
                    Exit For
                End If
            Next
            If bFound = False Then  'Not Present in the Collection, New Item...
                'New Item(s) Added....
                bOrderUpdated = True
                Exit For
            End If
            If myObj.lQty <> Val(.SubItems(6)) Then
                'Qty Changed....
                bOrderUpdated = True
                Exit For
            End If
        End With
    Next
    
    If bOrderUpdated Then
        IsOrderQtyChanges = True
    Else
        IsOrderQtyChanges = False
    End If
    
End Function

Private Sub PrintOrder(strOrderNo As String)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFCustOrders.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{CustOrders.OrderNo}='" & strOrderNo & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillCmbs()
    
    Call AddToCombo(cmbTradeTerms, "Description", "TradeTerms", " Order By ID")
    cmbTradeTerms.Text = "FOB-Sialkot"
    
    With cmbQuality
        .AddItem "Satin"
        .AddItem "Brilliant"
        .AddItem "Satin/Brilliant"
        .AddItem "Sand Blast"
    End With
'    With cmbTradeTerms
'        .Clear
'        .AddItem "FOB"
'        .AddItem "CIF (Air)"
'        .AddItem "CIF (Sea)"
'        .AddItem "C&F (Air)"
'        .AddItem "C&F (Sea)"
'        .ListIndex = 0
'    End With
    
    With cmbPartial
        .AddItem "Allowed"
        .AddItem "Not Allowed"
        .ListIndex = 0
    End With
    
    With cmbTrans
        .AddItem "Allowed"
        .AddItem "Not Allowed"
        .ListIndex = 0
    End With
    
    Call AddToCombo(cmbPayment, "Description", "PaymentTerms")
    
    cmbCompany.ListHeight = 6000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " ORDER BY CompanyName"
    
    Dim lCompanyID As Long
    lCompanyID = GetSingleLongValue("MIN(EntryID)", "Companies")
    cmbCompany.ID = lCompanyID
    
'    With cmbQuality
'        .AddItem "Regular"
'        .AddItem "Top"
'        .ListIndex = 0
'    End With
    
    Call AddToCombo(cmbFinQuality, "Description", "ItemsFinQuality", , , "Description")
    Call AddToCombo(cmbFinQualityE, "Description", "ItemsFinQuality")
    
    With cmbOrderType
        .AddItem "Stock Order"
        .AddItem "Sample Order"
        .AddItem "Customer Order"
        .AddItem "Replacement Order"
        .ListIndex = 2
    End With
    Call AddToCombo(cmbSMethod, "Description", "ShippingMethods", " ORDER BY EntryID")
        
    With cmbItemFinishingType
        .AddItem "Finished"
        .AddItem "Temper Inspection"
        .AddItem "Fist Inspection"
        .ListIndex = 0
    End With
    
End Sub

Private Sub FillItems()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "SELECT * FROM VVendor_Catalog WHERE Vend_AccNo='" & cmbCust.ID & "'", con_ServerSide, adOpenForwardOnly, adLockReadOnly
      
        cmbItems.Clear
        cmbDesc.Clear
        cmbCustItemID.Clear
        Do Until .EOF
            cmbCustItemID.AddItem !ITemID_Vendor & ""
            If chkCust Then
                cmbItems.AddItem ![ITemID_Vendor] & ""
                cmbDesc.AddItem ![Description] & ""
            Else
                cmbItems.AddItem ![ITemID] & ""
                'cmbDesc.AddItem ![CompDesc] & ""
                cmbDesc.AddItem ![Description] & ""
            End If
            
            'cmbItems.List(cmbItems.ListCount - 1, 1) = ![Unit] & ""
            cmbItems.List(cmbItems.ListCount - 1, 1) = Val(![PackingMode] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 2) = Val(![FOB] & "")
            cmbItems.List(cmbItems.ListCount - 1, 3) = Val(![ExWorks] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 4) = Val(![CIFAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 5) = Val(![CIFSea] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 6) = Val(![CnFAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 7) = Val(![CnFSea] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 8) = Val(![CnIAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 9) = Val(![CnISea] & "")
            
            'cmbDesc.List(cmbDesc.ListCount - 1, 1) = ![Description] & ""
            'cmbDesc.List(cmbDesc.ListCount - 1, 2) = ![CompDesc] & ""
            
            cmbDesc.List(cmbDesc.ListCount - 1, 2) = 1 'Val(![FOBTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 3) = 1 'Val(![ExWorksTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 4) = 1 'Val(![CIFAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 5) = 1 'Val(![CIFSeaTop] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 6) = 1 'Val(![CnFAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 7) = 1 'Val(![CnFSeaTop] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 8) = 1 'Val(![CnIAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 9) = 1 'Val(![CnISeaTop] & "")
            .MoveNEXT
        Loop
        .Close
    End With
    'Debug.Print Now
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub AttachPDF(p_lEntryID As Long, strPDFPath As String)
    
    On Error Resume Next
    con.Execute "DELETE FROM Vendor_Orders_Items_PDF WHERE FOI_RefID=" & p_lEntryID
    con.Execute "INSERT INTO Vendor_Orders_Items_PDF(FOI_RefID) VALUES(" & p_lEntryID & ")"
    
    If strPDFPath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,PDF FROM Vendor_Orders_Items_PDF WHERE FOI_RefID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
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

Private Sub txtWeight_KeyPress(KeyAscii As Integer)
    If keyascci = 13 Then
        Call txtQty_KeyPress(13)
    End If
End Sub


