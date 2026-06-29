VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNewMaterial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Detail"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Maerial Info."
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
      Height          =   3570
      Left            =   105
      TabIndex        =   3
      Top             =   15
      Width           =   6645
      Begin TabDlg.SSTab SSTab1 
         Height          =   2775
         Left            =   75
         TabIndex        =   8
         Top             =   675
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   4895
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
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
         TabCaption(0)   =   "Material Detail"
         TabPicture(0)   =   "frmNewMaterial.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TBox(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "TBox(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(6)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(5)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label1(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "TBox(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblItemID"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label1(4)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "TBox(0)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label1(1)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkInBalls"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkInGloves"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkInGarments"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "cmbUnit"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Accounts Info."
         TabPicture(1)   =   "frmNewMaterial.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmbPchAcc"
         Tab(1).Control(1)=   "Label1(7)"
         Tab(1).Control(2)=   "Label1(8)"
         Tab(1).Control(3)=   "TBox(4)"
         Tab(1).Control(4)=   "cmbPRTAcc"
         Tab(1).Control(5)=   "Label1(12)"
         Tab(1).Control(6)=   "cmbAssetAcc"
         Tab(1).Control(7)=   "Label1(14)"
         Tab(1).Control(8)=   "TBox(5)"
         Tab(1).Control(9)=   "Label1(9)"
         Tab(1).Control(10)=   "TBox(6)"
         Tab(1).Control(11)=   "Label1(10)"
         Tab(1).Control(12)=   "TBox(7)"
         Tab(1).Control(13)=   "Label1(13)"
         Tab(1).ControlCount=   14
         Begin VB.ComboBox cmbUnit 
            Height          =   315
            Left            =   5505
            TabIndex        =   9
            Top             =   1275
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rev. Price 3"
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
            Index           =   13
            Left            =   -71805
            TabIndex        =   38
            Top             =   1785
            Width           =   990
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   7
            Left            =   -71805
            TabIndex        =   37
            Top             =   2010
            Width           =   1455
            VariousPropertyBits=   679495711
            Size            =   "2566;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rev. Price 2"
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
            Index           =   10
            Left            =   -73305
            TabIndex        =   36
            Top             =   1785
            Width           =   990
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   -73335
            TabIndex        =   35
            Top             =   2010
            Width           =   1500
            VariousPropertyBits=   679495711
            Size            =   "2646;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rev. Price 1"
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
            Index           =   9
            Left            =   -74790
            TabIndex        =   34
            Top             =   1785
            Width           =   990
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   5
            Left            =   -74850
            TabIndex        =   33
            Top             =   2010
            Width           =   1485
            VariousPropertyBits=   679495711
            Size            =   "2619;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.CheckBox chkInGarments 
            Height          =   345
            Left            =   3360
            TabIndex        =   31
            Top             =   2235
            Width           =   1875
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   4
            Size            =   "3307;609"
            Value           =   "0"
            Caption         =   "Used In Garments"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CheckBox chkInGloves 
            Height          =   345
            Left            =   1650
            TabIndex        =   30
            Top             =   2235
            Width           =   1605
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   4
            Size            =   "2831;609"
            Value           =   "0"
            Caption         =   "Used In Gloves"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CheckBox chkInBalls 
            Height          =   345
            Left            =   150
            TabIndex        =   29
            Top             =   2235
            Width           =   1410
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   4
            Size            =   "2487;609"
            Value           =   "0"
            Caption         =   "Used In Balls"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Asset Account"
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
            Left            =   -71730
            TabIndex        =   28
            Top             =   465
            Width           =   1215
         End
         Begin MSForms.ComboBox cmbAssetAcc 
            Height          =   315
            Left            =   -71790
            TabIndex        =   27
            Top             =   690
            Width           =   3075
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "5424;556"
            ListWidth       =   8819
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Purchase Return Account"
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
            Index           =   12
            Left            =   -74835
            TabIndex        =   26
            Top             =   1095
            Width           =   2145
         End
         Begin MSForms.ComboBox cmbPRTAcc 
            Height          =   315
            Left            =   -74850
            TabIndex        =   25
            Top             =   1335
            Width           =   3015
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "5318;556"
            ListWidth       =   8819
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   -71790
            TabIndex        =   24
            Top             =   1335
            Width           =   1440
            VariousPropertyBits=   679495707
            Size            =   "2540;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Purchase Price"
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
            Left            =   -71730
            TabIndex        =   23
            Top             =   1110
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Purchase Account"
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
            Left            =   -74820
            TabIndex        =   22
            Top             =   450
            Width           =   1515
         End
         Begin MSForms.ComboBox cmbPchAcc 
            Height          =   315
            Left            =   -74835
            TabIndex        =   21
            Top             =   690
            Width           =   3015
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "5318;556"
            ListWidth       =   8819
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Material Name"
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
            Index           =   1
            Left            =   1920
            TabIndex        =   20
            Top             =   480
            Width           =   1230
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   1905
            TabIndex        =   19
            Top             =   690
            Width           =   4395
            VariousPropertyBits=   679495707
            Size            =   "7752;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Material ID"
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
            Left            =   135
            TabIndex        =   18
            Top             =   480
            Width           =   945
         End
         Begin VB.Label lblItemID 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   690
            Width           =   1725
         End
         Begin MSForms.TextBox TBox 
            Height          =   900
            Index           =   1
            Left            =   135
            TabIndex        =   16
            Top             =   1275
            Width           =   4155
            VariousPropertyBits=   -1467987941
            MaxLength       =   255
            ScrollBars      =   2
            Size            =   "7329;1587"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
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
            Index           =   2
            Left            =   135
            TabIndex        =   15
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Qty In Hand"
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
            Index           =   3
            Left            =   4365
            TabIndex        =   14
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reord Point"
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
            Index           =   5
            Left            =   4395
            TabIndex        =   13
            Top             =   1650
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Unit"
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
            Index           =   6
            Left            =   5520
            TabIndex        =   12
            Top             =   1080
            Width           =   345
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   2
            Left            =   4350
            TabIndex        =   11
            Top             =   1275
            Width           =   1125
            VariousPropertyBits=   679495707
            Size            =   "1984;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   4350
            TabIndex        =   10
            Top             =   1845
            Width           =   1965
            VariousPropertyBits=   679495707
            Size            =   "3466;556"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Catagory :"
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
         Left            =   2520
         TabIndex        =   7
         Top             =   375
         Width           =   870
      End
      Begin MSForms.ComboBox cmbCat 
         Height          =   300
         Left            =   3435
         TabIndex        =   6
         Top             =   330
         Width           =   3045
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5371;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbCatID 
         Height          =   300
         Left            =   855
         TabIndex        =   5
         Top             =   330
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cat. # :"
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
         Left            =   135
         TabIndex        =   4
         Top             =   390
         Width           =   600
      End
   End
   Begin MSForms.CheckBox chkInactive 
      Height          =   345
      Left            =   165
      TabIndex        =   32
      Top             =   3645
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
      Height          =   405
      Left            =   3360
      TabIndex        =   1
      Top             =   3690
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;714"
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
      Height          =   405
      Left            =   1635
      TabIndex        =   0
      Top             =   3690
      Width           =   1665
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2937;714"
      Picture         =   "frmNewMaterial.frx":0038
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   405
      Left            =   5085
      TabIndex        =   2
      Top             =   3690
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;714"
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
Attribute VB_Name = "frmNewMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Add As Boolean, Edit As Boolean
Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Dim dPrice As Double

Private Sub FillCmbs()

    Call AddToCombo(cmbCatID, "CatID", "ItemCatagories", "Order By CATID")
    
    Call AddToCombo(cmbCat, "Catagory", "ItemCatagories", " Order By CATID")
    
    
    
    Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PchAcc & "%' AND Parent=0 Order By ACCTitle")
    
    
    
    Call AddToCombo(cmbPRTAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PRTACC & "%' AND Parent=0 Order By ACCTitle")
    
    Call AddToCombo(cmbAssetAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & STOCKASSETACC & "%' AND Parent=0 Order By ACCTitle")
    
    Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    
End Sub



Private Sub chkHasSno_Click()
   Me.Width = IIf(chkHasSno = vbChecked, 10000, 8000)
   cmbSnos.Visible = chkHasSno = vbChecked
   Me.Visible = False
   Me.Show 1
End Sub

Private Sub cmbAssetAcc_Validate(Cancel As Boolean)
   If Not cmbAssetAcc.MatchFound Then cmbAssetAcc = ""
End Sub

Private Sub cmbCatID_Change()
    If cmbCatID.MatchFound Then Call cmbCatID_Click Else If Not cmbCatID = "" Then cmbCatID.DropDown
End Sub
Private Sub cmbCatID_Click()
    If cmbCat.ListIndex <> cmbCatID.ListIndex Then cmbCat.ListIndex = cmbCatID.ListIndex
    If Add Then lblItemID = GetNextItemID(cmbCatID)
End Sub
Private Sub cmbCatID_Validate(Cancel As Boolean)
    If Not cmbCatID.MatchFound Then cmbCatID.ListIndex = cmbCat.ListIndex
End Sub
Private Sub cmbCat_Change()
    If cmbCat.MatchFound Then Call cmbCat_Click Else If cmbCat <> "" Then cmbCat.DropDown
End Sub
Private Sub cmbCat_Click()
    If cmbCatID.ListIndex <> cmbCat.ListIndex Then cmbCatID.ListIndex = cmbCat.ListIndex
End Sub
Private Sub cmbCat_Validate(Cancel As Boolean)
   If Not cmbCat.MatchFound Then cmbCat.ListIndex = cmbCatID.ListIndex
End Sub

Private Sub cmbITemType_Click()

cmbPchAcc.Enabled = (cmbITemType.ListIndex <> 1)
cmbPRTAcc.Enabled = (cmbITemType.ListIndex = 0)
cmbSRTAcc.Enabled = (cmbITemType.ListIndex = 0)
cmbAssetAcc.Enabled = (cmbITemType.ListIndex <> 1)
TBox(3).Enabled = (cmbITemType.ListIndex <> 1)
TBox(4).Enabled = (cmbITemType.ListIndex <> 1)
TBox(5).Enabled = (cmbITemType.ListIndex <> 1)

chkHasSno.Visible = (cmbITemType.ListIndex = 0)

If cmbITemType.ListIndex <> 0 And chkHasSno = vbChecked Then chkHasSno = vbUnchecked

End Sub

Private Sub cmbsaleAcc_Validate(Cancel As Boolean)
    If Not cmbSaleAcc.MatchFound Then cmbSaleAcc = ""
End Sub

Private Sub cmbPchAcc_Validate(Cancel As Boolean)
    If Not cmbPchAcc.MatchFound Then cmbPchAcc = ""
End Sub



Private Sub cmbSnos_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cmbSnos_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If SaveItem Then
    Unload Me
    Me.Add = True
    Me.Edit = False
    Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Form_Activate()
   If Edit And Not Loaded Then Call showdata
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   Me.KeyPreview = True
'    DTPicker1.value = ServerDate
    Call FillCmbs
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err

    Dim OpBalance As Double
    Dim rec As New ADODB.Recordset

    If Not cmbCatID.MatchFound Then
        MsgBox "Invalid Item Catagory", vbInformation
        SaveItem = False
        Exit Function
    ElseIf Trim(TBox(0)) = "" Then
        MsgBox "Invalid Item Name", vbInformation
        SaveItem = False
        Exit Function
    ElseIf Trim(cmbUnit) = "" And cmbPchAcc.Enabled Then
        MsgBox "No Unit Selected.", vbInformation
        SaveItem = False
        Exit Function
    ElseIf cmbPchAcc.Enabled And Not cmbPchAcc.MatchFound Then
        MsgBox "No Purchase Account Selected", vbInformation
        SaveItem = False
        Exit Function
    ElseIf cmbAssetAcc.Enabled And Not cmbAssetAcc.MatchFound Then
        MsgBox "No Asset Account Selected", vbInformation
        SaveItem = False
        Exit Function
    'ElseIf cmbSaleAcc.Enabled And Not cmbSaleAcc.MatchFound Then
    '    MsgBox "No Sales Account Selected", vbInformation
    '    SaveItem = False
    '    Exit Function
    'ElseIf cmbSRTAcc.Enabled And Not cmbSRTAcc.MatchFound Then
    '    MsgBox "No Sale Return Account Selected", vbInformation
    '    SaveItem = False
    '    Exit Function
    ElseIf cmbPRTAcc.Enabled And Not cmbPRTAcc.MatchFound Then
        MsgBox "No Purchase Return Account Selected", vbInformation
        SaveItem = False
        Exit Function
    ElseIf TBox(4).Enabled And Val(TBox(4)) = 0 Then
        MsgBox "Invalid Purchase Price", vbInformation
        SaveItem = False
        Exit Function
    'ElseIf TBox(5).Enabled And Val(TBox(5)) = 0 Then
    '    MsgBox "Invalid Sale Price", vbInformation
    '    SaveItem = False
    '    Exit Function
    'ElseIf Val(TBox(4)) >= Val(TBox(5)) And Val(TBox(4)) > 0 Then
    '    If MsgBox("Sale Price Should Be More Than Purchase Price." & vbNewLine & "Proceed Anyway ?", vbInformation + vbYesNo) = vbNo Then
    '        SaveItem = False
    '        Exit Function
    '    End If
    
    'ElseIf chkHasSno = vbChecked Then
    '    If cmbSnos.ListCount <> Val(TBox(2)) Then
    '        MsgBox "List Of Serial No.'s is Not Equal To The Quantity.", vbInformation
    '        saveitems = False
    '        Exit Function
    '    End If
    End If


    Dim PchAcc As String, ItmPRTAcc As String, AssetAcc As String

    'con.BeginTrans
    Call StartTrans(con)

    If Add Then

        'If cmbITemType.ListIndex <> 1 Then

            Dim TotalCost As Double

            AssetAcc = getAccountNo(cmbAssetAcc)      'Item's Asset Account
   
            PchAcc = getAccountNo(cmbPchAcc)          'Item's Purchase Account
   
            'If cmbITemType.ListIndex = 0 Then
            '    ItmPRTAcc = getAccountNo(cmbPRTAcc)       'Item's Purchase Return Account
                'itmSRTAcc = getAccountNo(cmbSRTAcc)       'Item's Sale Return Account
            'End If
   
            TotalCost = Val(TBox(4)) * Val(TBox(2))
   
            con.Execute "Update Accounts Set OpenBal=OpenBal+" & TotalCost & " Where AccNo='" & AssetAcc & "'"
            Call adjustLedger(AssetAcc, "1/1/1900", 0, TotalCost)

        'End If

        'SaleAcc = getAccountNo(cmbSaleAcc)        'Item's Sale Account

   
        ''''  Save Items
   
        con.Execute "Insert into Material(MaterialID,MaterialName,CatID,Unit,MaterialDesc,AssetAccNo,PchAccno,PRTAccNo,Price,InHand,Reorder,InActive,UsedInBalls,UsedInGloves,UsedInGarments) " & _
         "Values('" & lblItemID & "','" & TBox(0).Text & "','" & cmbCatID & "','" & cmbUnit & "','" & TBox(1) & "','" & AssetAcc & "','" & _
         PchAcc & "','" & ItmPRTAcc & "'," & Val(TBox(4)) & "," & Val(TBox(2)) & "," & Val(TBox(3)) & "," & Abs(chkInactive.value) & "," & Abs(chkInBalls.value) & "," & Abs(chkInGloves.value) & "," & Abs(chkInGarments.value) & ")"
        
    Else
        Dim UpdateSQL As String
        If dPrice <> Val(TBox(4)) Then
            UpdateSQL = "Update Material Set MaterialName='" & TBox(0) & "',Unit='" & cmbUnit & "',MaterialDesc='" & TBox(1) & "',Price=" & Val(TBox(4)) & ",R1Price=Price,R2Price=R1Price,R3PRice=R2Price" & ",Reorder =" & Val(TBox(3)) & ",InActive=" & Abs(chkInactive.value) & ",UsedInBalls=" & Abs(chkInBalls.value) & ",UsedInGloves=" & Abs(chkInGloves.value) & ",UsedInGarments=" & Abs(chkInGarments.value) & " Where MaterialID='" & Me.Tag & "'"
        Else
            UpdateSQL = "Update Material Set MaterialName='" & TBox(0) & "',Unit='" & cmbUnit & "',MaterialDesc='" & TBox(1) & "',Price=" & Val(TBox(4)) & ",Reorder =" & Val(TBox(3)) & ",InActive=" & Abs(chkInactive.value) & ",UsedInBalls=" & Abs(chkInBalls.value) & ",UsedInGloves=" & Abs(chkInGloves.value) & ",UsedInGarments=" & Abs(chkInGarments.value) & " Where MaterialID='" & Me.Tag & "'"
        End If
        con.Execute UpdateSQL
            
    End If

    
    con.CommitTrans

    SaveItem = True
    Exit Function

err:
    MsgBox err.Description
    
End Function




Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 2, 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 4, 5
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
End Sub



Private Function GetNextItemID(CatID As String) As String
    Dim NextVal As String
    NextVal = GetMax("Cast(Right(MaterialID,2) as int)", "Material", "Where CatID ='" & CatID & "'")
    GetNextItemID = CatID & "-" & Format(Val(NextVal) + 1, "000")
End Function

Private Sub Tbox_Validate(Index As Integer, Cancel As Boolean)
    If Index = 4 Or Index = 5 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0.00")
    ElseIf Index = 2 Or Index = 3 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0")
    End If
End Sub


Private Sub showdata()

Dim rs As New ADODB.Recordset


With rs
    .Open "Select * From VMaterial Where MaterialID='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
    lblItemID = ![MaterialID] & ""
    cmbCatID = ![CatID] & ""
    TBox(0) = ![MaterialName] & ""
    TBox(1) = ![MaterialDesc] & ""
    TBox(2) = Val(![InHand] & "")
    TBox(3) = Val(![ReOrder] & "")
    TBox(4) = Format(Val(![Price] & ""), "#0.00")
    TBox(5) = Format(Val(![R1Price] & ""), "#0.00")
    TBox(6) = Format(Val(![R2Price] & ""), "#0.00")
    TBox(7) = Format(Val(![R3Price] & ""), "#0.00")
    
    dPrice = Val(![Price] & "")
    
    If (![PurchaseAccount] & "") <> "" Then Call FindInCombo(cmbPchAcc, ![PurchaseAccount] & "") Else cmbPchAcc = ""
    
    'If ![NonInventory] Then cmbITemType.ListIndex = 1 Else cmbITemType.ListIndex = 0
    'cmbITemType.ListIndex = Val(![ItemType] & "")
    
    'If (![SRTAccount] & "") <> "" Then Call FindInCombo(cmbSRTAcc, ![SRTAccount] & "") Else cmbSrtAccno = ""
    
    If (![PRTAccount] & "") <> "" Then Call FindInCombo(cmbPRTAcc, ![PRTAccount] & "") Else cmbPRTAcc = ""
    
    
    If (![AssetAccount] & "") <> "" Then Call FindInCombo(cmbAssetAcc, ![AssetAccount] & "") Else cmbAssetAcc = ""
    
    Call FindInCombo(cmbUnit, ![Unit] & "")

    chkInactive.value = ![Inactive]
    chkInBalls.value = ![UsedInBalls]
    chkInGloves.value = ![UsedInGloves]
    chkInGarments.value = ![UsedInGarments]
        
        
    .Close
End With

Set rs = Nothing

cmbPchAcc.Enabled = False


cmbPRTAcc.Enabled = False

cmbAssetAcc.Enabled = False

TBox(2).Enabled = False
Loaded = True

End Sub
