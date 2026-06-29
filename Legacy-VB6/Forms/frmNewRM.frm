VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Material"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   9390
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
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3870
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3030
      Top             =   5925
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewRM.frx":0000
            Key             =   "0'"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewRM.frx":0892
            Key             =   "7'"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewRM.frx":0CE4
            Key             =   "5'"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewRM.frx":1136
            Key             =   "6'"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewRM.frx":3470
            Key             =   "4'"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewRM.frx":414A
            Key             =   "2'"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewRM.frx":49DC
            Key             =   "3'"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewRM.frx":718E
            Key             =   "1'"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Device Info."
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
      Height          =   5940
      Left            =   90
      TabIndex        =   18
      Top             =   15
      Width           =   9180
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
         Height          =   615
         Left            =   105
         TabIndex        =   26
         Top             =   270
         Width           =   7725
         Begin MSForms.ComboBox cmbCatID 
            Height          =   315
            Left            =   1485
            TabIndex        =   0
            Tag             =   "GroupID"
            Top             =   210
            Width           =   2640
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4657;556"
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
            Caption         =   "Category :"
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
            Index           =   16
            Left            =   585
            TabIndex        =   28
            Top             =   270
            Width           =   870
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Device Code :"
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
            Left            =   4200
            TabIndex        =   27
            Top             =   270
            Width           =   1125
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   5370
            TabIndex        =   1
            Tag             =   "RMID1"
            Top             =   225
            Width           =   1650
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2910;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4830
         Left            =   120
         TabIndex        =   19
         Top             =   1005
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   8520
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   582
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
         TabCaption(0)   =   "   General Info."
         TabPicture(0)   =   "frmNewRM.frx":7A20
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkMaleGrinding"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkReturnable"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "    Account Info."
         TabPicture(1)   =   "frmNewRM.frx":7A3C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmbSaleAcc"
         Tab(1).Control(1)=   "cmbSaleRAcc"
         Tab(1).Control(2)=   "cmbAssetAcc"
         Tab(1).Control(3)=   "label1(13)"
         Tab(1).Control(4)=   "label1(12)"
         Tab(1).Control(5)=   "label1(14)"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Picture / Tech. Drawing"
         TabPicture(2)   =   "frmNewRM.frx":7A58
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdViewTechDrawing"
         Tab(2).Control(1)=   "cmdViewPicture"
         Tab(2).Control(2)=   "cmdChangeTechDrawing"
         Tab(2).Control(3)=   "lblTechDrawing"
         Tab(2).Control(4)=   "lblPic"
         Tab(2).Control(5)=   "label1(32)"
         Tab(2).Control(6)=   "Image1"
         Tab(2).Control(7)=   "label1(60)"
         Tab(2).Control(8)=   "ImgTechDrawing"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Storewise Stocks"
         TabPicture(3)   =   "frmNewRM.frx":7A74
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame3"
         Tab(3).Control(1)=   "txtEdit"
         Tab(3).Control(2)=   "LV"
         Tab(3).Control(3)=   "txtTotalStockInHand"
         Tab(3).Control(4)=   "label1(18)"
         Tab(3).ControlCount=   5
         TabCaption(4)   =   "Used In."
         TabPicture(4)   =   "frmNewRM.frx":7A90
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label2"
         Tab(4).Control(1)=   "LVItems"
         Tab(4).Control(2)=   "cmdAdd"
         Tab(4).Control(3)=   "cmbItem"
         Tab(4).ControlCount=   4
         Begin ComboList.Usercontrol1 cmbItem 
            Height          =   285
            Left            =   -74895
            TabIndex        =   62
            Top             =   705
            Width           =   8010
            _ExtentX        =   14129
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
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   315
            Left            =   -66855
            TabIndex        =   63
            Top             =   690
            Width           =   660
         End
         Begin MSComctlLib.ListView LVItems 
            Height          =   3735
            Left            =   -74895
            TabIndex        =   60
            Top             =   1005
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   6588
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S #"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item Code"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Name"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Frame Frame3 
            Height          =   930
            Left            =   -73980
            TabIndex        =   50
            Top             =   3630
            Visible         =   0   'False
            Width           =   6915
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Machining Stock"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   10
               Left            =   1455
               TabIndex        =   54
               Top             =   210
               Width           =   1140
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   6
               Left            =   1425
               TabIndex        =   53
               Top             =   435
               Width           =   1290
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2275;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Annealing Stock"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   9
               Left            =   135
               TabIndex        =   52
               Top             =   210
               Width           =   1140
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   5
               Left            =   105
               TabIndex        =   51
               Top             =   435
               Width           =   1290
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "2275;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
            End
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   -72270
            TabIndex        =   48
            Top             =   2715
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.CommandButton cmdViewTechDrawing 
            Caption         =   "View Actual Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   -74850
            TabIndex        =   43
            Top             =   4185
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmdViewPicture 
            Caption         =   "View Actual Size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   -67425
            TabIndex        =   42
            Top             =   4215
            Width           =   1200
         End
         Begin VB.CommandButton cmdChangeTechDrawing 
            Caption         =   "Update Drawing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   -71730
            TabIndex        =   41
            Top             =   4200
            Visible         =   0   'False
            Width           =   1200
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
            Height          =   4050
            Left            =   1200
            TabIndex        =   20
            Top             =   360
            Width           =   6765
            Begin VB.TextBox txtUrduFields 
               Alignment       =   1  'Right Justify
               DataSource      =   "datPrimaryRS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   0
               Left            =   90
               TabIndex        =   13
               Top             =   3555
               Width           =   6570
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   6
               Left            =   105
               TabIndex        =   4
               Tag             =   "RM_Maker_Description"
               Top             =   1095
               Width           =   6570
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "11589;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "RM Maker Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   20
               Left            =   135
               TabIndex        =   70
               Top             =   885
               Width           =   1545
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   2
               Left            =   5025
               TabIndex        =   68
               Tag             =   "Weight_Length"
               Top             =   2940
               Width           =   1650
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2910;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   4
               Left            =   4485
               TabIndex        =   69
               Tag             =   "OpeningStock"
               Top             =   2280
               Width           =   1110
               VariousPropertyBits=   679495711
               BorderStyle     =   1
               Size            =   "1958;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   1
               Left            =   105
               TabIndex        =   67
               Tag             =   "Description"
               Top             =   2940
               Width           =   4875
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "8599;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Weight / Length"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   19
               Left            =   5025
               TabIndex        =   66
               Top             =   2730
               Width           =   1155
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Urdu Name :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   19
               Left            =   135
               TabIndex        =   57
               Top             =   3330
               Width           =   900
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Maker Rate"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   17
               Left            =   5610
               TabIndex        =   56
               Top             =   2070
               Width           =   825
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   7
               Left            =   5625
               TabIndex        =   12
               Tag             =   "MakerRate"
               Top             =   2280
               Width           =   1065
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1879;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Raw Material Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   11
               Left            =   135
               TabIndex        =   55
               Top             =   2730
               Width           =   1770
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Opening Stock"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   8
               Left            =   4485
               TabIndex        =   40
               Top             =   2055
               Width           =   1035
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Max. Level"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   7
               Left            =   2310
               TabIndex        =   39
               Top             =   2055
               Width           =   780
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   3
               Left            =   2295
               TabIndex        =   10
               Tag             =   "MaxLevel"
               Top             =   2280
               Width           =   1065
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1879;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Min. Level"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   5
               Left            =   1215
               TabIndex        =   38
               Top             =   2055
               Width           =   720
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   2
               Left            =   1200
               TabIndex        =   9
               Tag             =   "MinLevel"
               Top             =   2280
               Width           =   1065
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1879;556"
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
               Left            =   105
               TabIndex        =   8
               Tag             =   "Rate"
               Top             =   2280
               Width           =   1065
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1879;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Rate"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   37
               Top             =   2055
               Width           =   345
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Reorder Point"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   3
               Left            =   3405
               TabIndex        =   36
               Top             =   2055
               Width           =   990
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   0
               Left            =   3390
               TabIndex        =   11
               Tag             =   "ReorderPoint"
               Top             =   2280
               Width           =   1065
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1879;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Raw Material Name"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   25
               Top             =   315
               Width           =   1380
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   0
               Left            =   105
               TabIndex        =   2
               Tag             =   "RMName"
               Top             =   525
               Width           =   4875
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "8599;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox NBox 
               Height          =   315
               Index           =   8
               Left            =   5025
               TabIndex        =   6
               Tag             =   "RMSize"
               Top             =   1680
               Width           =   690
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "1217;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Size"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   15
               Left            =   5055
               TabIndex        =   24
               Top             =   1470
               Width           =   285
            End
            Begin MSForms.ComboBox cmbUnit 
               Height          =   315
               Left            =   5025
               TabIndex        =   3
               Tag             =   "Unit"
               Top             =   525
               Width           =   1650
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "2910;556"
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
               Height          =   315
               Index           =   5
               Left            =   105
               TabIndex        =   5
               Tag             =   "RMUsage"
               Top             =   1680
               Width           =   4875
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "8599;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Item Usage"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   23
               Top             =   1470
               Width           =   825
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   6
               Left            =   5055
               TabIndex        =   22
               Top             =   315
               Width           =   285
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               Caption         =   "Unit"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   24
               Left            =   5790
               TabIndex        =   21
               Top             =   1470
               Width           =   285
            End
            Begin MSForms.ComboBox cmbSizeUnit 
               Height          =   315
               Left            =   5760
               TabIndex        =   7
               Tag             =   "SizeUnit"
               Top             =   1680
               Width           =   930
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "1640;556"
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
         End
         Begin ComboList.Usercontrol1 cmbSaleAcc 
            Height          =   285
            Left            =   -74085
            TabIndex        =   30
            Tag             =   "SaleAccNo"
            Top             =   975
            Width           =   3315
            _ExtentX        =   5847
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
         Begin ComboList.Usercontrol1 cmbSaleRAcc 
            Height          =   285
            Left            =   -74070
            TabIndex        =   31
            Tag             =   "SRTAccNo"
            Top             =   1530
            Width           =   3315
            _ExtentX        =   5847
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
         Begin ComboList.Usercontrol1 cmbAssetAcc 
            Height          =   285
            Left            =   -74070
            TabIndex        =   32
            Tag             =   "AssetAccNo"
            Top             =   2100
            Width           =   3315
            _ExtentX        =   5847
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
            Height          =   2640
            Left            =   -73980
            TabIndex        =   49
            Top             =   645
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   4657
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Store Name"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Opening"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Current Stock"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSForms.CheckBox chkReturnable 
            Height          =   345
            Left            =   1200
            TabIndex        =   65
            Tag             =   "Returnable"
            Top             =   4395
            Width           =   2415
            BackColor       =   -2147483633
            ForeColor       =   12582912
            DisplayStyle    =   4
            Size            =   "4260;609"
            Value           =   "0"
            Caption         =   "Returnable Material"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CheckBox chkMaleGrinding 
            Height          =   345
            Left            =   5550
            TabIndex        =   64
            Tag             =   "RequiresMaleGrinding"
            Top             =   4395
            Width           =   2415
            BackColor       =   -2147483633
            ForeColor       =   12582912
            DisplayStyle    =   4
            Size            =   "4260;609"
            Value           =   "0"
            Caption         =   "Requires Male Grinding"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
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
            Left            =   -74895
            TabIndex        =   61
            Top             =   510
            Width           =   420
         End
         Begin MSForms.TextBox txtTotalStockInHand 
            Height          =   315
            Left            =   -70665
            TabIndex        =   59
            Top             =   3315
            Width           =   1290
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "2275;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Stock :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   18
            Left            =   -71610
            TabIndex        =   58
            Top             =   3345
            Width           =   900
         End
         Begin VB.Label lblTechDrawing 
            Alignment       =   2  'Center
            Caption         =   "Click to Add PDF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -74700
            MouseIcon       =   "frmNewRM.frx":7AAC
            MousePointer    =   99  'Custom
            TabIndex        =   44
            Top             =   2355
            Width           =   4005
         End
         Begin VB.Label lblPic 
            Alignment       =   2  'Center
            Caption         =   "Double click to Add"
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -68955
            TabIndex        =   47
            Top             =   2190
            Width           =   1125
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Picture"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   32
            Left            =   -70515
            TabIndex        =   46
            Top             =   435
            Width           =   495
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   3525
            Left            =   -70515
            Stretch         =   -1  'True
            Top             =   615
            Width           =   4290
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Tech. Drawing"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   60
            Left            =   -74805
            TabIndex        =   45
            Top             =   435
            Width           =   1035
         End
         Begin VB.Image ImgTechDrawing 
            BorderStyle     =   1  'Fixed Single
            Height          =   3525
            Left            =   -74835
            Stretch         =   -1  'True
            Top             =   630
            Width           =   4290
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sale Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   -74040
            TabIndex        =   35
            Top             =   780
            Width           =   930
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Sale Return Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   -74040
            TabIndex        =   34
            Top             =   1320
            Width           =   1470
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Asset Account"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   14
            Left            =   -74055
            TabIndex        =   33
            Top             =   1890
            Width           =   1035
         End
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   29
         Top             =   645
         Visible         =   0   'False
         Width           =   1725
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "3043;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   7725
      TabIndex        =   16
      Top             =   6195
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
      Left            =   90
      TabIndex        =   17
      Tag             =   "InActive"
      Top             =   6180
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
      Left            =   6105
      TabIndex        =   15
      Top             =   6195
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
      Left            =   4515
      TabIndex        =   14
      Top             =   6195
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
Attribute VB_Name = "frmNewRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemCode As Long

Dim SelItm As ListItem

Dim DoPrint As Boolean

Dim PicFileName As String, TechDrawingFileName As String
Dim bSampling As Boolean
Dim strRMItemsDelList As String
Dim bCopy As Boolean, lRMID_For_Copy As Long

Public Sub EditItem(strItemCode As Long)
    
    cmbItem.ClearVals
    cmbItem.AddVals con, "ItemID", "Items", "ItemID", " WHERE ItemID NOT IN(SELECT ItemID FROM ItemsRMComp WHERE RMID=" & strItemCode & ")"
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    ItemCode = strItemCode
    'cmbCatID.Locked = True
    'TBox(3).Locked = True
    With rs
        .Open "SELECT * FROM VRM WHERE RMID=" & ItemCode, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    c.ID = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is VB.Label Then
                    c.Caption = .Fields(c.Tag) & ""
                Else
                    If c.Name = "cmbCatID" Then
                        c.Text = .Fields("GroupName") & ""
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        
        'TBox(8) = ![CatID] & "-"
        Tbox(3) = ![RMID1] & ""
        txtUrduFields(0) = !UrduName & ""
        If Not IsNull(![Pic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "Pic"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
        
        If bCopy = False Then
            If Not IsNull(![TechnicalDrawing]) Then
                lblTechDrawing.Caption = "Click here to Open Tech. Drawing."
                lblTechDrawing.Tag = ""
                cmdChangeTechDrawing.Visible = True
            Else
                lblTechDrawing.Caption = "Click here to Add."
                cmdChangeTechDrawing.Visible = False
            End If
        End If
        
        Dim lCount As Long
        lCount = GetSingleLongValue("Group_ID", "RMGroupIDsForForging", "WHERE Group_ID=" & !GroupID)
        If lCount > 0 Then
            Frame3.Visible = True
            NBox(5) = Val(!AnnealingStock & "")
            NBox(6) = Val(!MachiningStock & "")
        Else
            Frame3.Visible = False
        End If
        
        .Close
        
        If bCopy = False Then
            Dim dTotalStockInHand As Double
            .Open "SELECT * FROM VRMStatusStorewise WHERE RMID=" & ItemCode, con, adOpenForwardOnly, adLockReadOnly
    '        LV.ListItems.Clear
            dTotalStockInHand = 0
            Do Until .EOF
                On Error Resume Next
                Set ITM = LV.ListItems(!StoreName & "")
                If err.Number = 35601 Then GoTo SkipStore
                ITM.SubItems(1) = Val(!OpeningBalance & "")
                ITM.ListSubItems(1).Tag = Val(!OpeningBalance & "")
                dCurrentStock = Val(!NetQty & "") 'GetSingleDoubleValue("NetQty", "VRMStatusStorewise", " WHERE RMID=" & ItemCode )
                ITM.SubItems(2) = dCurrentStock
                ITM.ListSubItems(2).Tag = dCurrentStock
                ITM.Tag = Val(!EntryID)
                dTotalStockInHand = dTotalStockInHand + dCurrentStock
    '            dTotalStock = dTotalStock + dCurrentStock
    '            dTotalOpening = dTotalOpening + Val(!OpeningBalance & "")
SkipStore:
                
                .MoveNEXT
            Loop
            .Close
            On Error GoTo 0
        End If
        
        .Open "SELECT * FROM VItemsRMComp WHERE RMID=" & ItemCode, con, adOpenForwardOnly, adLockReadOnly
        LVItems.ListItems.Clear
        Do Until .EOF
            Set ITM = LVItems.ListItems.add(, , LVItems.ListItems.count + 1)
            If bCopy = False Then
                ITM.Tag = Val(!EntryID & "")
            End If
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , !ItemName & ""
            .MoveNEXT
        Loop
        .Close
        
    End With
    txtTotalStockInHand = dTotalStockInHand
    Call GetTotals
        
    Set rs = Nothing
    
    If bCopy Then
        ItemCode = 0
    Else
        Tbox(3).Locked = True
        Me.Show 1
    End If
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillCmbs()
        
    'Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PchAcc & "%' AND Parent=0 Order By ACCTitle")
    Call cmbSaleAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & SALEACC & "%' AND Parent=0 Order By ACCTitle")
    cmbSaleAcc.ID = "31-001-001"
    'Call AddToCombo(cmbPRTAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PRTAcc & "%' AND Parent=0 Order By ACCTitle")
    Call cmbSaleRAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & SRTACC & "%' AND Parent=0 Order By ACCTitle")
    cmbSaleRAcc.ID = "31-003-003"
    
    Call cmbAssetAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & STOCKASSETACC & "%' AND Parent=0 Order By ACCTitle")
    cmbAssetAcc.ID = "15-005-001"
    'Call AddToCombo(cmbAssetAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & StockAssetAcc & "%' AND Parent=0 Order By ACCTitle")
    
    'Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    'Call AddToCombo(cmbWeightUnit, "Distinct UnitWeight", "Items", "Order By UnitWeight")
    
    Call AddToCombo(cmbCatID, "Description", "RMGroups", " Order By ID ", , "ID")
    
    '----------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select UnitName From Units Order By UnitName", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbUnit.AddItem .Fields(0) & ""
            cmbSizeUnit.AddItem .Fields(0) & ""
            'cmbWeightUnit.AddItem .Fields(0) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    cmbItem.ListHeight = 2400
    cmbItem.DropDownOffLine = False
    'cmbItem.AddVals con, "'{' + ItemID + '}' + ItemName", "Items", "ItemID"
    'cmbItem.AddVals con, "ItemID", "Items", "ItemID"
    '----------------------------------------------------------------
End Sub

Private Sub cmbCatID_Change()
    Call cmbCatID_Click
End Sub

Private Sub cmbCatID_Click()
    If Not cmbCatID.MatchFound Then Exit Sub

End Sub


Private Sub cmbItem_GotFocus()
    Me.KeyPreview = False
End Sub

Private Sub cmbItem_LostFocus()
    Me.KeyPreview = True
End Sub

Private Sub cmdAdd_Click()

    If cmbItem.MatchFound = False Then
        MsgBox "Please Select Article.", vbInformation
        Exit Sub
'        Dim lCount As Long
'        lCount = GetSingleLongValue("COUNT(*)", "Items", " WHERE ItemID='" & cmbItem.Text & "'")
'        If lCount = 0 Then
'            con.Execute "INSERT INTO Items(ItemID,ItemName,InActive,GroupID,CatID) VALUES('" & cmbItem.Text & "','" & cmbItem.Text & "',0,10,3)"
'        End If
    End If
    
    Dim i As Integer
    For i = 1 To LVItems.ListItems.count
        If cmbItem.Text = LVItems.ListItems(i).SubItems(1) Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        End If
    Next
    
    Dim ITM As ListItem
    Set ITM = LVItems.ListItems.add(, , LVItems.ListItems.count + 1)
    ITM.Tag = "New"
    ITM.ListSubItems.add , , cmbItem.Text
    ITM.ListSubItems.add , , "" 'GetSingleStringValue("ItemName", "Items", " WHERE ItemID='" & cmbItem.ID & "'")
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangeTechDrawing_Click()
    AddTechDrawing
End Sub

Private Sub cmdViewPicture_Click()

    If Image1.Picture = 0 Then Exit Sub
    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromPic Image1.Picture

End Sub


Private Sub cmdNext_Click()
    If SaveItem Then
        Unload Me
        
            
        Load frmNewRM
        Call frmNewRM.AddNew
    
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub



Private Sub cmdViewTechDrawing_Click()
    If ImgTechDrawing.Picture = 0 Then Exit Sub
    Dim f As New frmViewPictureFSExt
    Load f
    f.ShowFromPic ImgTechDrawing.Picture
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Public Sub AddNew(Optional p_bSampling As Boolean = False, Optional p_bCopy As Boolean = False, Optional p_lRMID_For_Copy As Long = 0)

    bSampling = p_bSampling
    bCopy = p_bCopy
    
    lRMID_For_Copy = p_lRMID_For_Copy
        
    If bCopy Then
        Call EditItem(lRMID_For_Copy)
        Tbox(3) = ""
    End If
    Call getRMID
    
    cmbItem.ClearVals
    cmbItem.AddVals con, "ItemID", "Items", "ItemID"
    Me.Show 1
    
End Sub

Private Sub getRMID()
    
    On Error GoTo err
    If ShowingData Then Exit Sub
    Dim strVendID As String
    
    
    'strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where UPPER(Left(VendID1,3))=" & UCase(Left(cmbProcess.Text, 3))).Fields(0).value & "")
    
    'If VendID <> 0 Then
        
    '    TBox(3) = VendID1
    '    Exit Sub
        
    'End If
    
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        Tbox(3).Locked = False
    Else
        If strCompany <> "Tecno" Then
            strVendID = Val(con.Execute("SELECT MAX(Cast(SubString(RMID1,4,4) as int)) From RM").Fields(0).Value & "")
            If strVendID = 0 Then
                strVendID = 0
            End If
            strVendID = strVendID + 1
            Tbox(3) = "RM" & Format(strVendID, "-0000")
        Else
            'do nothing
        End If
        
        
        
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
'    For i = 0 To SSTab1.Tabs - 1
'        SSTab1.Tab = i
'        SSTab1.Picture = ImageList1.ListImages(i & "'").ExtractIcon
'    Next
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        SSTab1.TabVisible(3) = False
    End If
    
    SSTab1.Tab = 0
        
    Me.KeyPreview = True
    '    DTPicker1.value = ServerDate
    cmbSaleAcc.ListHeight = 1500
    cmbSaleRAcc.ListHeight = 1500
    cmbAssetAcc.ListHeight = 1500
    
    cmbCatID.Locked = False
    Tbox(3).Locked = False
    
    Call FillCmbs
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM Stores", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !StoreName & "", !StoreName & "")
            ITM.ListSubItems.add(, , "").Tag = ""
            ITM.ListSubItems.add , , ""
            .MoveNEXT
        Loop
        .Close
    End With
    strRMItemsDelList = ""
'    ItemCode = 0
    ItemCode = 0
    bCopy = False
    'FRev.Visible = False
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err

    Dim OpBalance As Double
    
    If Trim(Tbox(0)) = "" Then
        MsgBox "Invalid Item Name", vbInformation
        SaveItem = False
        Exit Function
    ElseIf Trim(cmbUnit) = "" Then
        MsgBox "No Unit Selected.", vbInformation
        SaveItem = False
        Exit Function
'    ElseIf cmbAssetAcc.Enabled And Not cmbAssetAcc.MatchFound Then
'        MsgBox "No Asset Account Selected", vbInformation
'        SaveItem = False
'        Exit Function
    End If
    
    Dim PCHACC As String, ItmPRTAcc As String, AssetAcc As String

    Dim SQL As String
    
    Call StartTrans(con)
    
    Dim iRMID As Integer
    If ItemCode = 0 Then
            
        SQL = CreateInsertSQL("RM")
        'con.Execute "Update Accounts Set OpenBal=OpenBal+" & TotalCost & " Where AccNo='" & AssetAcc & "'"
        'Call adjustLedger(AssetAcc, "1/1/1900", 0, TotalCost)
    Else
        SQL = CreateUpdateSQL("RM")
        SQL = SQL & " WHERE RMID=" & ItemCode
    End If
    
    con.Execute SQL
    If ItemCode = 0 Then
        iRMID = GetSingleLongValue("MAX(RMID)", "RM")
        con.Execute "UPDATE RM SET Sampling=" & Abs(bSampling) & " WHERE RMID=" & iRMID
    Else
        iRMID = ItemCode
    End If
    
    If bUrduSupported Then
        con.Execute "UPDATE RM SET UrduName=N'" & txtUrduFields(0).Text & "' WHERE RMID=" & iRMID, i
    End If
    
    'If ItemCode <> "" Then Call SaveRevision
    If Image1.Picture <> 0 And Trim(PicFileName) <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        'rec.Close
        Rec.Open "SELECT RMID,Pic FROM RM WHERE RMID=" & iRMID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile PicFileName
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
    End If
    If Trim(TechDrawingFileName) <> "" Then
        'rec.Close
        Rec.Open "SELECT RMID,TechnicalDrawing FROM RM WHERE RMID=" & iRMID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile TechDrawingFileName
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
    End If
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.Tag) = 0 Then
                con.Execute "INSERT INTO RMStorewiseOpeningStocks(StoreName,RMID,OpeningBalance) VALUES('" & .Text & "'," & iRMID & "," & Val(.SubItems(1)) & ")"
            Else
                con.Execute "UPDATE RMStorewiseOpeningStocks SET OpeningBalance=" & Val(.SubItems(1)) & " WHERE EntryID=" & Val(.Tag)
            End If
        End With
    Next
    
    'Save Items...
    Call SaveItems(iRMID)
    con.CommitTrans
    '---------------------------------------------------------
    'If DoPrint Then
    '    If MsgBox("Item Has Changed." & vbNewLine & "Do You Want To Print This Item?", vbQuestion + vbYesNo) = vbYes Then
    '        Load frmPrevRpt
    '        frmPrevRpt.ShowReport "{VrptDMR.ItemID}='" & TBox(4) & "'", rptDMR
    '    End If
    'End If
    '---------------------------------------------------------
    ItemCode = 0
    bCopy = False
    SaveItem = True
        
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub SaveItems(p_iRMID As Integer)

    Dim i As Integer
    Dim iProcessID As Integer
    If strRMItemsDelList <> "" Then
        strRMItemsDelList = Left(strRMItemsDelList, Len(strRMItemsDelList) - 1)
        con.Execute "DELETE FROM ItemsRMComp WHERE EntryID IN(" & strRMItemsDelList & ")"
    End If
    For i = 1 To LVItems.ListItems.count
        With LVItems.ListItems(i)
            If Val(.Tag) = 0 Then
                iProcessID = GetSingleLongValue("TOP 1 ProcessID", "ItemProcesses", " WHERE ItemID='" & .SubItems(1) & "' AND ProcessID NOT IN(1,2) ORDER BY SNo")
'                If iProcessID = 0 Then
'                    MsgBox "Can't Assign " & .SubItems(1) & ". No Process Assigned.", vbInformation
'                Else
                    con.Execute "INSERT INTO ItemsRMComp(ItemID,RMID,Qty,ProcessID) VALUES('" & _
                     .SubItems(1) & "'," & p_iRMID & ",1," & iProcessID & ")"
'                End If
            End If
        End With
    Next
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    Set RevEntry = Nothing
End Sub


Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub Image1_DblClick()
    Call AddPic
End Sub

Private Sub lblPic_DblClick()
    Call AddPic
End Sub

Private Sub AddPic()
    On Error Resume Next
    
    PicFileName = ""
    CD1.Filter = "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        PicFileName = CD1.FileName
        Image1.Picture = LoadPicture(PicFileName)
        lblPic.Visible = False
    Else
        PicFileName = ""
        Image1.Picture = LoadPicture(PicFileName)
        lblPic.Visible = True
    End If
    
End Sub


Private Sub lblTechDrawing_Click()
    If lblTechDrawing.Caption = "Click here to Open Tech. Drawing." Then
        Call ShowTechDrawing("RM", "TechnicalDrawing", " WHERE RMID=" & ItemCode, CStr(ItemCode), CD1)
    Else
        Call AddTechDrawing
    End If
End Sub

Private Sub AddTechDrawing()

    On Error Resume Next
    
    TechDrawingFileName = ""
    CD1.Filter = "PDF Files Only|*.pdf"
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        TechDrawingFileName = CD1.FileName
        'ImgTechDrawing.Picture = LoadPicture(TechDrawingFileName)
        lblTechDrawing.Caption = CD1.FileTitle
        lblTechDrawing.Tag = CD1.FileName
    Else
        TechDrawingFileName = ""
        imgttechdrawing.Picture = LoadPicture(TechDrawingFileName)
        lblTechDrawing.Visible = True
    End If
    
End Sub


Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(2).Left, LV.Top + LV.SelectedItem.Top
        .Text = Val(LV.SelectedItem.SubItems(1))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End Sub

Private Sub LVItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVItems.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If Val(LVItems.SelectedItem.Tag) > 0 Then
            strRMItemsDelList = strRMItemsDelList & Val(LVItems.SelectedItem.Tag) & ","
        End If
        LVItems.ListItems.Remove LVItems.SelectedItem.Index
    End If
End Sub

Private Sub Tbox_Change(Index As Integer)
    Exit Sub
    Select Case Index
    
    Case 8, 3
        Tbox(4) = Tbox(8) & Tbox(3)
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
        Tbox(Index) = Format(Val(Tbox(Index)), "#0.00")
    ElseIf Index = 2 Or Index = 3 Then
        Tbox(Index) = Format(Val(Tbox(Index)), "#0")
    End If
End Sub


Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    
    'Create The First Part Of Insert Query...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Then
            If c.Tag <> "" Then
                RetString = RetString & c.Tag & ","
            End If
        End If
    Next
    
    'Get For ListView
    'For i = 1 To LV2.ListItems.Count
    '    RetString = RetString & LV2.ListItems(i).Key & ","
    'Next i
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Then
            
                If c.Name = "cmbCatID" Then
                    RetString = RetString & Val(c.List(c.ListIndex, 1)) & ","
                ElseIf c.Name = "NBox" Then
                    RetString = RetString & Val(c.Text & "") & ","
                ElseIf c.Name = "cmbSteel" Or c.Name = "cmbGuage" Or c.Name = "cmbItemGroup" Then
                    RetString = RetString & c.List(c.ListIndex, 1) & ","
                Else
                    RetString = RetString & "'" & Replace(c.Text, "'", "''") & "',"
                End If
                
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                'If C.Name = cmbSteel Or C.Name = cmbGuage Then
                '    RetString = RetString & C.ID & ","
                RetString = RetString & "'" & c.ID & "',"
            End If
        End If
    Next
    
    
    'Get For ListView
    'For i = 1 To LV2.ListItems.Count
    '    RetString = RetString & Abs(LV2.ListItems(i).Checked) & ","
    'Next i
    
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
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Then '
                If c.Name = "cmbCatID" Then
                    RetString = RetString & c.Tag & "=" & Val(c.List(c.ListIndex, 1)) & ","
                ElseIf c.Name = "NBox" Then
                    RetString = RetString & c.Tag & "=" & Val(c.Text) & ","
                ElseIf c.Name = "cmbSteel" Or c.Name = "cmbGuage" Or c.Name = "cmbItemGroup" Then
                    RetString = RetString & c.Tag & "=" & c.List(c.ListIndex, 1) & ","
                Else
                    RetString = RetString & c.Tag & "='" & Replace(c.Text, "'", "''") & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.ID & "',"
            End If
        End If
    Next
    
    If NotEdited = False Then
        
        DoPrint = True
    Else
        DoPrint = False
    End If
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
'        If Val(txtEdit) <= 0 Then
'            MsgBox "Invlaid Qty.", vbInformation
'            Exit Sub
'        End If
        LV.SelectedItem.SubItems(1) = Val(txtEdit)
        If Val(LV.SelectedItem.SubItems(1)) <> Val(LV.SelectedItem.ListSubItems(1).Tag) Then
            LV.SelectedItem.SubItems(2) = Val(LV.SelectedItem.ListSubItems(2).Tag) - Val(LV.SelectedItem.ListSubItems(1).Tag) + Val(LV.SelectedItem.SubItems(1))
        Else
            LV.SelectedItem.SubItems(2) = Val(LV.SelectedItem.SubItems(1))
        End If
        Call GetTotals
        txtEdit.Visible = False
        
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub GetTotals()

    Dim i As Integer
    Dim dTotalStock As Double, dCurrentStock As Double
    For i = 1 To LV.ListItems.count
        dTotalStock = dTotalStock + Val(LV.ListItems(i).SubItems(2))
        dTotalOpening = dTotalOpening + Val(LV.ListItems(i).SubItems(1))
    Next
    NBox(4) = dTotalOpening
    'NBox(5) = dTotalStock
    
End Sub

Private Sub txtUrduFields_GotFocus(Index As Integer)
    Call SetUrduLanguage
End Sub

Private Sub txtUrduFields_LostFocus(Index As Integer)
    Call SetEnglishLanguage
End Sub

