VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmProcesses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Processes"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleMode       =   0  'User
   ScaleWidth      =   11003.84
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbBillingProcess 
      Height          =   285
      Left            =   6465
      TabIndex        =   31
      Top             =   1530
      Width           =   3195
      _ExtentX        =   5636
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
   Begin VB.TextBox txtUrduFields 
      DataSource      =   "datPrimaryRS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   4770
      TabIndex        =   4
      Top             =   915
      Width           =   7455
   End
   Begin ComboList.Usercontrol1 cmbInspectionProcess 
      Height          =   285
      Left            =   5145
      TabIndex        =   8
      Top             =   2175
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
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
      Enabled         =   0   'False
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbMaker 
      Height          =   285
      Left            =   2130
      TabIndex        =   6
      Top             =   1530
      Width           =   4320
      _ExtentX        =   7620
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
      Enabled         =   0   'False
      Appearance      =   0
   End
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
      Left            =   4770
      TabIndex        =   2
      Top             =   285
      Width           =   7455
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   60
      TabIndex        =   0
      Top             =   285
      Width           =   1185
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1770
      Top             =   7065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   60
      TabIndex        =   3
      Top             =   930
      Width           =   4680
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   330
      Top             =   6900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcesses.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcesses.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   6975
      Left            =   60
      TabIndex        =   12
      Top             =   2460
      Width           =   12165
      Begin ComboList.Usercontrol1 cmbProcessSearch 
         Height          =   285
         Left            =   90
         TabIndex        =   32
         Top             =   225
         Width           =   7575
         _ExtentX        =   13361
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
      Begin VB.CheckBox chkShowUrdu 
         Caption         =   "Show Urdu Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   30
         Top             =   6495
         Width           =   2430
      End
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   105
         ScaleHeight     =   255
         ScaleWidth      =   5295
         TabIndex        =   15
         Top             =   1110
         Visible         =   0   'False
         Width           =   5325
         Begin MSForms.ComboBox cmbAuthReqE 
            Height          =   300
            Left            =   3840
            TabIndex        =   24
            Top             =   -30
            Width           =   1455
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2566;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbOperationE 
            Height          =   300
            Left            =   1830
            TabIndex        =   22
            Top             =   -30
            Width           =   2055
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3625;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSupervisor 
            Height          =   300
            Left            =   1215
            TabIndex        =   18
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDescription 
            Height          =   300
            Left            =   600
            TabIndex        =   17
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSNo 
            Height          =   300
            Left            =   -15
            TabIndex        =   16
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5775
         Left            =   90
         TabIndex        =   19
         Top             =   585
         Width           =   11985
         _ExtentX        =   21140
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Supervisor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Operation"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fix Contractor"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Auth. Req."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Insp. Process"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Process Name Other"
            Object.Width           =   0
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   9150
         TabIndex        =   14
         Top             =   6495
         Width           =   1425
         ForeColor       =   0
         Caption         =   "Print         "
         PicturePosition =   327683
         Size            =   "2514;635"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   360
         Left            =   10650
         TabIndex        =   11
         Top             =   6495
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close     "
         PicturePosition =   327683
         Size            =   "2514;635"
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
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1275
      TabIndex        =   1
      Top             =   285
      Width           =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Process Detailed Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   4830
      TabIndex        =   33
      Top             =   705
      Width           =   2025
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "*you are currently in the Editing Mode. Click on Add or Clear to Continue."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   29
      Top             =   1875
      Visible         =   0   'False
      Width           =   5280
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   11100
      TabIndex        =   10
      Top             =   1920
      Width           =   1125
      Caption         =   "Clear"
      PicturePosition =   65543
      Size            =   "1984;661"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Billing Process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   6480
      TabIndex        =   28
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fix Contractor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   2160
      TabIndex        =   27
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÑÇÓíÓ ˜Ç äÇã "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   19
      Left            =   11040
      TabIndex        =   26
      Top             =   15
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   25
      Top             =   60
      Width           =   375
   End
   Begin MSForms.ComboBox cmbAuthReq 
      Height          =   300
      Left            =   9675
      TabIndex        =   7
      Top             =   1530
      Width           =   2535
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4471;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Authorization Req."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   9705
      TabIndex        =   23
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Operation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   21
      Top             =   1320
      Width           =   720
   End
   Begin MSForms.ComboBox cmbOperation 
      Height          =   300
      Left            =   60
      TabIndex        =   5
      Top             =   1530
      Width           =   2055
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3625;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Supervisor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   705
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Process Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   1305
      TabIndex        =   13
      Top             =   60
      Width           =   1395
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   375
      Left            =   9945
      TabIndex        =   9
      Top             =   1920
      Width           =   1125
      Caption         =   "Save"
      PicturePosition =   65543
      Size            =   "1984;661"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bInspectionProcesses As Boolean

Private Sub chkShowUrdu_Click()

    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If chkShowUrdu.Value = vbChecked Then
            LV.ListItems(i).ListSubItems(1) = LV.ListItems(i).ListSubItems(2).Tag
            LV.ListItems(i).ListSubItems(6) = LV.ListItems(i).ListSubItems(5).Tag
        Else
            LV.ListItems(i).ListSubItems(1) = LV.ListItems(i).ListSubItems(1).Tag
            LV.ListItems(i).ListSubItems(6) = LV.ListItems(i).ListSubItems(3).Tag
        End If
    Next
    If chkShowUrdu.Value = vbChecked Then
        LV.ColumnHeaders(2).Text = "ÑÇÓíÓ ˜Ç äÇã "
        LV.ColumnHeaders(7).Text = "ÇäÓ˜Ôä ÑÇÓíÓ"
    Else
        LV.ColumnHeaders(2).Text = "Process"
        LV.ColumnHeaders(7).Text = "Insp. Process"
    End If
    
End Sub

Private Sub cmbAuthReq_Change()
    Call cmbAuthReq_Click
End Sub

Private Sub cmbAuthReq_Click()

    If cmbAuthReq.ListIndex = 1 Then
        cmbInspectionProcess.Enabled = True
        cmbInspectionProcess.ClearVals
        cmbInspectionProcess.AddVals con, "'{' + Code + '} ' + ProcessName", "InspectionProcesses", "EntryID"
    Else
        cmbInspectionProcess.ClearVals
        cmbInspectionProcess.Enabled = True
    End If
    
End Sub

Private Sub cmbAuthReqE_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
    
        con.Execute "UPDATE Processes Set SNo=" & txtSNo & ",Description='" & txtDescription & _
         "',Supervisor='" & txtSupervisor & "',Operation=" & cmbOperationE.ListIndex & ",AuthRequired=" & _
         cmbAuthReqE.ListIndex & " Where ProcessID=" & Val(LV.SelectedItem.key), iAffect
         
        If iAffect = 1 Then
            PicEdit.Visible = False
            Call FillList
        End If
        
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub cmbOperation_Change()
    Call cmbOperation_Click
End Sub

Private Sub cmbOperation_Click()

    If cmbOperation.ListIndex = 1 Then
        
        cmbMaker.Enabled = True
        cmbMaker.ClearVals
        cmbMaker.AddVals con, "VenderName", "Makers", "VendID", " WHERE Active=1"
    Else
        cmbMaker.ClearVals
        cmbMaker.Enabled = False
    End If
    
End Sub

Private Sub cmbOperationE_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim iAffect As Integer
    If KeyCode.Value = 13 Then
        cmbAuthReqE.SetFocus
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub cmbProcessSearch_matched()

    LV.ListItems(cmbProcessSearch.ID & "'").Selected = True
    LV.SelectedItem.EnsureVisible
    LV.SetFocus
    
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    If txtDesc = "" Then
        MsgBox "Please Enter Process Name", vbInformation
        Exit Sub
    ElseIf cmbOperation.MatchFound = False Then
        MsgBox "Please Select Operation.", vbInformation
        Exit Sub
    ElseIf cmbAuthReq.MatchFound = False Then
        MsgBox "Please Select Authorization.", vbInformation
        Exit Sub
    End If
    
    LV.Sorted = False
    
    Dim iAffect As Long
        
    If Val(lblMsg.Tag) = 0 Then
    
        Dim MaxSNo As Integer
        MaxSNo = Val(con.Execute("SELECT Max(SNo) FROM Processes").Fields(0) & "")
        MaxSNo = MaxSNo + 1
        
        con.Execute "INSERT INTO Processes(SNo,Description,Supervisor,Operation,AuthRequired,Code," & _
        "ProcessNameUrdu,Insp_RefID,Fix_Maker_RefID,InspectionProcess,ProcessNameUrduOther,BillingProcessID) VALUES(" & MaxSNo & ",'" & txtDesc & "','" & _
        Text1 & "'," & cmbOperation.ListIndex & "," & cmbAuthReq.ListIndex & ",'" & txtCode & _
        "',N'" & txtUrduFields(0) & "'," & Val(cmbInspectionProcess.ID) & "," & Val(cmbMaker.ID) & "," & Abs(bInspectionProcesses) & ",N'" & txtUrduFields(0) & "'," & Val(cmbBillingProcess.ID) & ")", iAffect
    Else    'Editing
        con.Execute "UPDATE Processes SET Description='" & txtDesc & "',Supervisor='" & Text1 & _
        "',Operation=" & cmbOperation.ListIndex & ",AuthRequired=" & cmbAuthReq.ListIndex & _
        ",Code='" & txtCode & "',ProcessNameUrdu=N'" & txtUrduFields(0) & "',Insp_RefID=" & _
        Val(cmbInspectionProcess.ID) & ",Fix_Maker_RefID=" & Val(cmbMaker.ID) & ",ProcessNameUrduOther=N'" & txtUrduFields(1) & "',BillingProcessID=" & Val(cmbBillingProcess.ID) & " WHERE ProcessID=" & Val(lblMsg.Tag), iAffect
    End If
    
    If iAffect = 1 Then
        Call FillList
        txtDesc = ""
        txtCode = ""
        txtUrduFields(0) = ""
        txtUrduFields(1) = ""
        Text1 = ""
    End If
    
    lblMsg.Visible = False
    lblMsg.Tag = "0"
    
    LV.Sorted = True
    txtCode.SetFocus
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    lblMsg.Visible = False
    lblMsg.Tag = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Processes.rpt")
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "{VProcesses.InspectionProcess}=" & bInspectionProcesses, rpt
    Unload Me
End Sub

Private Sub FillList()

    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    Dim strOperation As String
    
    SQL = "SELECT * FROM VProcesses WHERE InspectionProcess=" & Abs(bInspectionProcesses) & " ORDER BY Code"
    
    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, ![ProcessID] & "'", !code & "")
            ITM.ListSubItems.add(, , ![Description] & "").Tag = !Description & ""
            ITM.ListSubItems.add(, , ![Supervisor] & "").Tag = !ProcessNameUrdu & ""
            strOperation = ""
            If Val(!Operation & "") = 0 Then
                strOperation = "Employee"
            ElseIf Val(!Operation & "") = 1 Then
                strOperation = "Contractor"
            ElseIf Val(!Operation & "") = 2 Then
                strOperation = "Both"
            End If
            ITM.ListSubItems.add(, , strOperation).Tag = !Insp_ProcessName & ""
            ITM.ListSubItems.add(, , !VenderName & "").Tag = Val(!Fix_Maker_RefID & "")
            
            If IsNull(!AuthRequired) Then
                ITM.ListSubItems.add(, , "No").Tag = !Insp_ProcessNameUrdu & ""
            Else
                ITM.ListSubItems.add(, , cmbAuthReq.List(Abs(!AuthRequired))).Tag = !Insp_ProcessNameUrdu & ""
            End If
            
            ITM.ListSubItems.add(, , !Insp_ProcessName & "").Tag = Val(!Insp_RefID & "")
            ITM.ListSubItems.add(, , !ProcessNameUrduOther & "" & "").Tag = Val(!BillingProcessID & "")
            
            .MoveNEXT
        Next
        .Close
    End With
    Set rsList = Nothing
    SQL = ""
    
    cmbProcessSearch.ClearVals
    cmbProcessSearch.AddVals con, "Description", "Processes", "ProcessID", " WHERE InspectionProcess=" & Abs(bInspectionProcesses)
    
End Sub

Private Sub txtDedE_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    Dim theKey As String
    Dim Affect As Long
    theKey = LV.SelectedItem.key
    theKey = Right(theKey, Len(theKey) - 1)
    
    If KeyAscii = 13 Then
        con.Execute "Update TaxRanges Set SalFrom=" & _
         txtFromE & ",SalTo=" & txtToE & ",PerDeduct=" & _
         txtPerE & ",Reduction=" & txtDedE & _
         " Where SalFrom=" & theKey, Affect
        If Affect = 1 Then
            With LV.SelectedItem
                .key = "'" & txtFromE
                .Text = txtFromE
                .SubItems(1) = txtToE
                .SubItems(2) = txtPerE
                .SubItems(3) = txtDedE
            End With
        End If
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        txtFromE = ""
        txtToE = ""
        txtPerE = ""
        txtDedE = ""
        Pic.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub Form_Load()

    With cmbOperation
        .AddItem "Employee"
        .AddItem "Contractor"
        .AddItem "Both"
    End With
    With cmbOperationE
        .AddItem "Employee"
        .AddItem "Contractor"
        .AddItem "Both"
    End With
    With cmbAuthReq
        .AddItem "No"
        .AddItem "Yes"
    End With
    With cmbAuthReqE
        .AddItem "No"
        .AddItem "Yes"
    End With
    
    cmbMaker.ListHeight = 4000
    cmbInspectionProcess.ListHeight = 4000
    cmbBillingProcess.ListHeight = 4000
    
    cmbBillingProcess.ClearVals
    cmbBillingProcess.AddVals con, "'{' + Code + '} ' + Description", "Processes", "ProcessID"
    
End Sub

Private Sub Form_Resize()
    Me.Top = 1700 'MainForm.Toolbar1.Top + MainForm.Toolbar1.Height + 100
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub
    With LV
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
                ColumnHeader.Icon = 2
            Else
                .SortOrder = lvwAscending
                ColumnHeader.Icon = 1
            End If
        Else
            LV.ColumnHeaderIcons = Nothing
            Set LV.ColumnHeaderIcons = ImageList1
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            ColumnHeader.Icon = 1
        End If
        .Sorted = True
    End With
End Sub

Private Sub LV_DblClick()

    lblMsg.Visible = True
    lblMsg.Tag = Val(LV.SelectedItem.key)
    '---------------------------------------
    With LV.SelectedItem
        txtCode = .Text
        txtDesc = .ListSubItems(1).Tag
        txtUrduFields(0) = .ListSubItems(2).Tag
        txtUrduFields(1) = .SubItems(7)
        Text1 = .SubItems(2)
        cmbOperation.Text = .SubItems(3)
   
        cmbAuthReq.ListIndex = IIf(.SubItems(5) = "Yes", 1, 0)
        cmbMaker.ID = .ListSubItems(5).Tag
        cmbInspectionProcess.ID = .ListSubItems(3).Tag
        cmbBillingProcess.ID = Val(.ListSubItems(7).Tag)
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PicEdit.Visible = False
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If LV.ListItems.count = 0 Then Exit Sub
    'On Error Resume Next
    
    If KeyCode = 46 Then
        If MsgBox("Do You Want To Delete This Process?", vbQuestion + vbYesNo) = vbYes Then
            'Check If Assigned to Maker...
            Dim lCount As Long, lProcessID As Long
            lProcessID = Val(LV.SelectedItem.key)
            lCount = 0
            lCount = GetSingleLongValue("COUNT(ProcessID)", "MakerProcesses", " WHERE ProcessID=" & lProcessID)
            If lCount > 0 Then
                MsgBox "Can't Delete, It's Assigned to Makers.", vbInformation
                Exit Sub
            End If
            'Now Check If Assigned to Item
            lCount = 0
            lCount = GetSingleLongValue("COUNT(ProcessID)", "ItemProcesses", " WHERE ProcessID=" & lProcessID)
            If lCount > 0 Then
                MsgBox "Can't Delete, It's Assigned to Items.", vbInformation
                Exit Sub
            End If
            
            Dim Affect As Integer
            con.Execute "DELETE FROM Processes WHERE ProcessID=" & Val(LV.SelectedItem.key), Affect
            If Affect = 1 Then
                LV.ListItems.Remove (LV.SelectedItem.Index)
            End If
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Inspection Points")
    
    If strRet = "Inspection Points" Then
        Dim f As New frmProcesses_Inspection_Points
        Load f
        f.ShowMe (Val(LV.SelectedItem.key))
    End If
End Sub

Private Sub txtDescription_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    
    If KeyCode.Value = 13 Then
        With txtSupervisor
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
    
End Sub

Private Sub txtSNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode.Value = 13 Then
        With txtDescription
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSupervisor_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    Dim iAffect As Integer
    If KeyCode.Value = 13 Then
        cmbOperationE.SetFocus
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtUrduFields_GotFocus(Index As Integer)
    Call SetUrduLanguage
End Sub

Private Sub txtUrduFields_LostFocus(Index As Integer)
    Call SetEnglishLanguage
End Sub

Public Sub ShowMe(Optional p_bInspectionProcesses As Boolean = False)

    bInspectionProcesses = p_bInspectionProcesses
    Call FillList
    Me.Show 1
    
End Sub
