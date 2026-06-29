VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProductionUserSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Access Rights"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductionUserSettings.frx":0000
   LinkTopic       =   "Form24"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5040
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtUserName 
      DataField       =   "FullUserName"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2610
      TabIndex        =   9
      Top             =   225
      Width           =   2370
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   30
      TabIndex        =   5
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Common Options"
      TabPicture(0)   =   "frmProductionUserSettings.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Venders"
      TabPicture(1)   =   "frmProductionUserSettings.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Raw Material"
      TabPicture(2)   =   "frmProductionUserSettings.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Process Assignment"
      TabPicture(3)   =   "frmProductionUserSettings.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Hico Resources"
      TabPicture(4)   =   "frmProductionUserSettings.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image2"
      Tab(4).Control(1)=   "Label3"
      Tab(4).Control(2)=   "Label1(9)"
      Tab(4).ControlCount=   3
      Begin VB.Frame Frame2 
         Height          =   5220
         Left            =   -74655
         TabIndex        =   38
         Top             =   510
         Width           =   6660
         Begin MSComctlLib.ListView LV 
            Height          =   5010
            Left            =   60
            TabIndex        =   39
            Top             =   150
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   8837
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   -2147483633
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Iss. Auth."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Rcv. Auth."
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   2160
         Left            =   -74835
         TabIndex        =   29
         Top             =   540
         Width           =   5475
         Begin VB.CheckBox Check1 
            Caption         =   "Issue Raw Material"
            DataField       =   "RMIssRM"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   2835
            TabIndex        =   37
            Top             =   1241
            Width           =   1920
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Recieve Raw Material"
            DataField       =   "RMRcvRM"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   2835
            TabIndex        =   36
            Top             =   778
            Width           =   1920
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Raw Material Vender P.O. List"
            DataField       =   "RMVPOList"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   2835
            TabIndex        =   35
            Top             =   315
            Width           =   2445
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Raw Material Vender Ass. Items"
            DataField       =   "RMVAssItems"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   135
            TabIndex        =   34
            Top             =   1704
            Width           =   2640
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Raw Material Vender New P.O."
            DataField       =   "RMVNewPO"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   2835
            TabIndex        =   33
            Top             =   1704
            Width           =   2580
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Raw Material Vender List"
            DataField       =   "RMVendList"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   32
            Top             =   1241
            Width           =   2085
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Raw Material"
            DataField       =   "RawMaterial"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   31
            Top             =   778
            Width           =   1650
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Raw Material Groups"
            DataField       =   "RMGroups"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   30
            Top             =   315
            Width           =   2010
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   2385
         Left            =   -74835
         TabIndex        =   17
         Top             =   600
         Width           =   4785
         Begin VB.CheckBox Check1 
            Caption         =   "Vender Payment Sheet"
            DataField       =   "VendPymtSheet"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   135
            TabIndex        =   28
            Top             =   1503
            Width           =   2010
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Vender Ledger"
            DataField       =   "VendLedger"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   135
            TabIndex        =   27
            Top             =   1905
            Width           =   1440
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Vender Advance Payments"
            DataField       =   "VendAdvPymt"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   26
            Top             =   1102
            Width           =   2340
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Vender List"
            DataField       =   "VenderList"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   135
            TabIndex        =   25
            Top             =   300
            Width           =   1890
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Vender Assigned Items"
            DataField       =   "VendAssItems"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   135
            TabIndex        =   24
            Top             =   701
            Width           =   2010
         End
         Begin VB.Frame Frame4 
            Height          =   1830
            Left            =   2415
            TabIndex        =   19
            Top             =   285
            Width           =   2205
            Begin VB.CheckBox Check1 
               Caption         =   "Short Term Ledger"
               DataField       =   "ShortTermLedger"
               DataSource      =   "Adodc1"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   18
               Left            =   165
               TabIndex        =   23
               Top             =   1065
               Width           =   1650
            End
            Begin VB.CheckBox Check1 
               Caption         =   "New Long Term Loan"
               DataField       =   "Compositions"
               DataSource      =   "Adodc1"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   19
               Left            =   165
               TabIndex        =   22
               Top             =   690
               Width           =   1860
            End
            Begin VB.CheckBox Check1 
               Caption         =   "New Short Term Loan"
               DataField       =   "NewShortTerm"
               DataSource      =   "Adodc1"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   20
               Left            =   165
               TabIndex        =   21
               Top             =   300
               Width           =   1905
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Long Term Ledger"
               DataField       =   "LongTermLedger"
               DataSource      =   "Adodc1"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   17
               Left            =   165
               TabIndex        =   20
               Top             =   1455
               Width           =   1650
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Loans"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   2505
            TabIndex        =   18
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame7 
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
         Height          =   1635
         Left            =   585
         TabIndex        =   6
         Top             =   1005
         Width           =   4785
         Begin VB.CheckBox Check1 
            Caption         =   "Hico Visible"
            DataField       =   "HicoVisible"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   2640
            TabIndex        =   16
            Top             =   1245
            Width           =   1440
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Backup Data"
            DataField       =   "BackupData"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   225
            TabIndex        =   15
            Top             =   300
            Width           =   1440
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Restore Data"
            DataField       =   "RestoreData"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   14
            Top             =   615
            Width           =   1440
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Can Authorize"
            DataField       =   "Authorize"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   225
            TabIndex        =   13
            Top             =   930
            Width           =   1440
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Add Production Planning"
            DataField       =   "AddProdPlan"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   12
            Top             =   1245
            Width           =   2040
         End
         Begin VB.CheckBox Check1 
            Caption         =   "User Management"
            DataField       =   "UserManagement"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   2640
            TabIndex        =   11
            Top             =   300
            Width           =   1650
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Requsition"
            DataField       =   "Requsition"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   2640
            TabIndex        =   10
            Top             =   930
            Width           =   1440
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Change Password"
            DataField       =   "ChangePassword"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   7
            Top             =   615
            Width           =   1650
         End
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1260
         Left            =   -72630
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Double Click To Add Thumb Exp. /Signature"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   -72465
         TabIndex        =   41
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Thumb Exp/Signature"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   -72630
         TabIndex        =   40
         Top             =   1005
         Width           =   1545
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User Complete Name"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   1770
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   6765
      TabIndex        =   1
      Top             =   7590
      Width           =   1425
      Caption         =   " Cancel       "
      PicturePosition =   327683
      Size            =   "2514;661"
      Picture         =   "frmProductionUserSettings.frx":0396
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   375
      Left            =   5265
      TabIndex        =   0
      Top             =   7590
      Width           =   1425
      Caption         =   "OK           "
      PicturePosition =   327683
      Size            =   "2514;661"
      Picture         =   "frmProductionUserSettings.frx":04A8
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label USer 
      AutoSize        =   -1  'True
      Caption         =   "User Login Name"
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
      Left            =   675
      TabIndex        =   4
      Top             =   0
      Width           =   1410
   End
   Begin VB.Label lblUserName 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "UserName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   645
      TabIndex        =   3
      Top             =   225
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "This User Has Following Access Rights, Selected Option Shows That Access Is Granted While Unselected Access is Denied  . . . "
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
      Height          =   630
      Index           =   0
      Left            =   615
      TabIndex        =   2
      Top             =   570
      Width           =   7395
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   90
      Picture         =   "frmProductionUserSettings.frx":08FA
      Top             =   -30
      Width           =   540
   End
End
Attribute VB_Name = "frmProductionUserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As New ADODB.Recordset
Dim Loaded As Boolean
Dim ColIndex As Integer

Dim PicFileName As String

Private Sub AddSignature()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        Image2.Picture = LoadPicture(Picfile)
        PicFileName = Picfile
        Label3.Visible = False
    End If
    
End Sub
Private Sub Check1_Click(Index As Integer)
    
    
    Dim bEnable As Boolean
    If Index = 21 Then 'Loan Check Box
        bEnable = IIf(Check1(21).value = vbChecked, True, False)
        For i = 17 To 20
            Check1(i).Enabled = bEnable
        Next i
    ElseIf Index = 40 Then 'Items
        bEnable = IIf(Check1(40).value = vbChecked, True, False)
        For i = 43 To 45
            Check1(i).Enabled = bEnable
        Next i
    ElseIf Index = 58 Then 'Export --> Company
        bEnable = IIf(Check1(58).value = vbChecked, True, False)
        For i = 55 To 65
            Check1(i).value = Check1(58).value
        Next i
    ElseIf Index = 68 Then 'Export --> Customers
        bEnable = IIf(Check1(68).value = vbChecked, True, False)
        For i = 66 To 67
            Check1(i).value = Check1(68).value
        Next i
    ElseIf Index = 74 Then 'Export --> Agents
        bEnable = IIf(Check1(74).value = vbChecked, True, False)
        For i = 75 To 76
            Check1(i).value = Check1(74).value
        Next i
    ElseIf Index = 69 Then 'Export --> Orders
        bEnable = IIf(Check1(69).value = vbChecked, True, False)
        For i = 70 To 73
            Check1(i).value = Check1(69).value
        Next i
    ElseIf Index = 79 Then 'Export --> PInvoice
        bEnable = IIf(Check1(79).value = vbChecked, True, False)
        For i = 77 To 78
            Check1(i).value = Check1(79).value
        Next i
    ElseIf Index = 101 Then 'Export --> Bank Documents
        bEnable = IIf(Check1(101).value = vbChecked, True, False)
        For i = 102 To 103
            Check1(i).value = Check1(101).value
        Next i
    ElseIf Index = 100 Then 'Export --> cGMP Documents
        bEnable = IIf(Check1(100).value = vbChecked, True, False)
        For i = 96 To 99
            Check1(i).value = Check1(100).value
        Next i
    ElseIf Index = 106 Then 'Export --> Payment Status
        bEnable = IIf(Check1(106).value = vbChecked, True, False)
        For i = 104 To 108
            Check1(i).value = Check1(106).value
        Next i
    ElseIf Index = 111 Then 'Export --> Rebate Documents
        bEnable = IIf(Check1(111).value = vbChecked, True, False)
        For i = 109 To 110
            Check1(i).value = Check1(111).value
        Next i
    ElseIf Index = 84 Then 'Export --> Custom Documents
        bEnable = IIf(Check1(84).value = vbChecked, True, False)
        For i = 80 To 87
            Check1(i).value = Check1(84).value
        Next i
        Check1(112).value = Check1(84).value   'Hide Package
    ElseIf Index = 91 Then 'Export --> Commercial Documents
        bEnable = IIf(Check1(91).value = vbChecked, True, False)
        For i = 88 To 95
            Check1(i).value = Check1(91).value
        Next i
    End If
    
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo err
    rsUser.CancelUpdate
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
    
    Call StartTrans(con)
    If rsUser![UserID] = 1 Then GoTo DoItForAdmin
    
    rsUser.Update
    
    
    con.Execute "Delete From UserAssProcesses Where UserID=" & rsUser![UserID]
    
    For i = 1 To LV.ListItems.Count
        If LV.ListItems(i).Checked Then
            con.Execute "Insert Into UserAssProcesses(UserID,ProcessID,ReqIssAuth,ReqRcvAuth) Values(" & _
             rsUser![UserID] & "," & Val(LV.ListItems(i).Key) & "," & IIf(LV.ListItems(i).SubItems(1) = "Yes", 1, 0) & "," & IIf(LV.ListItems(i).SubItems(1) = "Yes", 1, 0) & ")"
        End If
    Next i
    
DoItForAdmin:
    If (Image2.Picture = 0 Or Trim(PicFileName) = "") Then GoTo TheEnd

    Dim rec As New ADODB.Recordset
    
    Dim rsThumb As New ADODB.Stream
            
    rec.Open "Select UserID,Signature from Users Where UserID=" & rsUser![UserID], con, adOpenStatic, adLockOptimistic
    
    'Save Thumb Expression
    If Image2.Picture <> 0 And Trim(PicFileName) <> "" Then
        rsThumb.Open
        rsThumb.Type = adTypeBinary
        rsThumb.LoadFromFile PicFileName
        rec.Fields("Signature") = rsThumb.Read
        rsThumb.Close
    End If
    
    rec.Update
    rec.Close
    Set rec = Nothing
    Set rsThumb = Nothing
    
TheEnd:

    con.CommitTrans
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Public Sub ShowUserSettings(iUserID As Integer)
    
    
    Me.Hide
    
    
'    For i = 0 To SSTab1
'        SSTab1.TabVisible(i) = Not iUserID = 1
'    Next i
    
    If Export_Package Then
        For i = 0 To 10
            SSTab1.TabVisible(i) = False
        Next i
        
        
        SSTab1.TabVisible(10) = True
        
    End If
    
    With rsUser
        .Open "Select * From Users Where UserID=" & iUserID, con, adOpenDynamic, adLockOptimistic
        On Error Resume Next
        For i = 0 To Check1.Count - 1
            Set Check1(i).DataSource = rsUser
        Next
        
        Set txtUserName.DataSource = rsUser
        
        If Not IsNull(![Signature]) Then
            Set Image2.DataSource = rsUser
            Image2.DataField = "Signature"
            Label3.Visible = False
        Else
            Label3.Visible = True
        End If
        
        lblUserName.Caption = " " & ![UserName] & "" 'UserName & ""
        txtUserName = ![FullUserName] & ""
            
            
    End With
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From UserAssProcesses Where UserID=" & iUserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            LV.ListItems(![ProcessID] & "'").Checked = True
            LV.ListItems(![ProcessID] & "'").SubItems(1) = IIf(![ReqIssAuth] & "" = True, "Yes", "No")
            LV.ListItems(![ProcessID] & "'").SubItems(2) = IIf(![ReqRcvAuth] & "" = True, "Yes", "No")
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Me.Show 1
    
End Sub

Private Sub Form_Load()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Loaded = False
    
    With rs
        .Open "Select ProcessID,Description From Processes Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![ProcessID] & "'", ![Description] & "")
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If rsUser.State = 1 Then rsUser.Close
    Set rsUser = Nothing
    OurOwnForm = True
End Sub


Private Sub Image2_DblClick()
    Call AddSignature
End Sub

Private Sub Label3_DblClick()
    Call AddSignature
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.Count = 0 Then Exit Sub
    
    If Not (ColIndex = 1 Or ColIndex = 2) Then Exit Sub
    If Val(LV.SelectedItem.Key) > 7 Then Exit Sub
    
    If LV.SelectedItem.SubItems(ColIndex) = "Yes" Then
        LV.SelectedItem.SubItems(ColIndex) = "No"
    Else
        LV.SelectedItem.SubItems(ColIndex) = "Yes"
    End If
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked Then
        If Val(Item.Key) > 7 Then
            Item.SubItems(1) = "No"
            Item.SubItems(2) = "No"
        Else
            Item.SubItems(1) = "Yes"
            Item.SubItems(2) = "Yes"
        End If
    Else

        Item.SubItems(1) = ""
        Item.SubItems(2) = ""
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Set Flag that mouse is in Issue column or In Rcvd.
    If CBool(x >= LV.ColumnHeaders(2).Left And x <= (LV.ColumnHeaders(2).Left + LV.ColumnHeaders(2).Width)) Then
        ColIndex = 1
    ElseIf CBool(x >= LV.ColumnHeaders(3).Left And x <= (LV.ColumnHeaders(3).Left + LV.ColumnHeaders(3).Width)) Then
        ColIndex = 2
    Else
        ColIndex = 0
    End If
    
End Sub
