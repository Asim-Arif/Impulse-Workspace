VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCVList 
   ClientHeight    =   9540
   ClientLeft      =   1110
   ClientTop       =   15
   ClientWidth     =   11955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin VB.Frame FraAdvSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Advanced Search :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   900
      TabIndex        =   8
      Top             =   225
      Visible         =   0   'False
      Width           =   10830
      Begin VB.ComboBox cmbCriteria 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         ItemData        =   "frmCVList.frx":0000
         Left            =   105
         List            =   "frmCVList.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   4530
         Width           =   2055
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Designation :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   105
         TabIndex        =   32
         Top             =   4950
         Width           =   5265
         Begin VB.ListBox lstDesignations 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2430
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   33
            Top             =   315
            Width           =   5070
         End
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7860
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Computer Literacy :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   5415
         TabIndex        =   28
         Top             =   3885
         Width           =   5265
         Begin VB.ListBox lstComputer 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2130
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   29
            Top             =   300
            Width           =   5070
         End
      End
      Begin VB.ComboBox cmbCriteria 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         ItemData        =   "frmCVList.frx":0017
         Left            =   5415
         List            =   "frmCVList.frx":0021
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3465
         Width           =   2055
      End
      Begin VB.ComboBox cmbCriteria 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         ItemData        =   "frmCVList.frx":002E
         Left            =   5430
         List            =   "frmCVList.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   420
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Other :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   5415
         TabIndex        =   23
         Top             =   6885
         Width           =   5265
         Begin VB.ComboBox cmbShortList 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "frmCVList.frx":0045
            Left            =   3555
            List            =   "frmCVList.frx":0055
            TabIndex        =   36
            Text            =   "Combo1"
            Top             =   300
            Width           =   1605
         End
         Begin VB.ComboBox cmbGender 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "frmCVList.frx":008A
            Left            =   915
            List            =   "frmCVList.frx":0097
            TabIndex        =   24
            Text            =   "Combo1"
            Top             =   300
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Short List :"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Index           =   2
            Left            =   2535
            TabIndex        =   35
            Top             =   375
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender :"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Index           =   3
            Left            =   90
            TabIndex        =   25
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.ComboBox cmbCriteria 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         ItemData        =   "frmCVList.frx":00B3
         Left            =   5415
         List            =   "frmCVList.frx":00BD
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   6495
         Width           =   2055
      End
      Begin VB.CommandButton cmdCloseAdvSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7860
         Width           =   1875
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Academic Qualification :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   5415
         TabIndex        =   19
         Top             =   810
         Width           =   5265
         Begin VB.ListBox lstAca 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2130
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   20
            Top             =   300
            Width           =   5070
         End
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6870
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7860
         Width           =   1875
      End
      Begin VB.CommandButton cmdAdvSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8790
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   7860
         Width           =   1890
      End
      Begin VB.ComboBox cmbCriteria 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         ItemData        =   "frmCVList.frx":00CA
         Left            =   105
         List            =   "frmCVList.frx":00D4
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Area :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   105
         TabIndex        =   14
         Top             =   1635
         Width           =   5265
         Begin VB.ListBox lstArea 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2430
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   315
            Width           =   5070
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Age :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   105
         TabIndex        =   9
         Top             =   285
         Width           =   5265
         Begin VB.TextBox txtAge 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2835
            TabIndex        =   12
            Top             =   300
            Width           =   1785
         End
         Begin VB.ComboBox cmbAge 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "frmCVList.frx":00E1
            Left            =   750
            List            =   "frmCVList.frx":00F1
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   300
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Years"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Index           =   1
            Left            =   4665
            TabIndex        =   13
            Top             =   375
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age :"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Index           =   0
            Left            =   255
            TabIndex        =   10
            Top             =   330
            Width           =   465
         End
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
      Height          =   795
      Left            =   30
      TabIndex        =   3
      Top             =   8715
      Width           =   9105
      Begin VB.TextBox txtValue 
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
         Left            =   2910
         TabIndex        =   5
         Top             =   390
         Width           =   2745
      End
      Begin VB.ComboBox cmbField 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmCVList.frx":011C
         Left            =   60
         List            =   "frmCVList.frx":012F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   390
         Width           =   2835
      End
      Begin MSForms.CommandButton cmdShowAdvSearch 
         Height          =   360
         Left            =   7155
         TabIndex        =   30
         Top             =   360
         Width           =   1830
         Caption         =   "Advance Search"
         PicturePosition =   327683
         Size            =   "3228;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdSearch 
         Height          =   360
         Left            =   5700
         TabIndex        =   7
         Top             =   360
         Width           =   1380
         Caption         =   "Search      "
         PicturePosition =   327683
         Size            =   "2434;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By :                                           Search For:"
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
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   5235
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   8610
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   15187
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CV ID"
         Object.Width           =   2013
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Interview Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Age"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Education"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total Exp."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "C/O"
         Object.Width           =   2822
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   10170
      TabIndex        =   2
      Top             =   8730
      Width           =   1665
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2937;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   375
      Left            =   8490
      TabIndex        =   1
      Top             =   8730
      Width           =   1680
      Caption         =   "Options         "
      PicturePosition =   327683
      Size            =   "2963;661"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenuEmp"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuempdash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewEmp 
         Caption         =   "Add New &CV"
      End
      Begin VB.Menu mnuempdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDashPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "Print List"
      End
      Begin VB.Menu mnuPrintCV 
         Caption         =   "Print CV"
      End
      Begin VB.Menu mnuDashInterview 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInterviewQuestions 
         Caption         =   "Interview Questions"
      End
      Begin VB.Menu mnuDashConvert 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvertToEmployee 
         Caption         =   "Convert to Employee"
      End
   End
End
Attribute VB_Name = "frmCVList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevIndex As Long
Public add As Boolean, edit As Boolean
Dim showAllDept As Boolean
Dim strRptCondition As String

Private Sub cmbAge_Change()
    Call cmbAge_Click
End Sub

Private Sub cmbAge_Click()
    If cmbAge.ListIndex = 0 Then
        txtAge = ""
        txtAge.Locked = True
    Else
        txtAge.Locked = False
    End If
End Sub

Private Sub cmbField_Change()
    txtValue = ""
End Sub

Private Sub cmdAdvSearch_Click()

    Dim strCondition As String
    
    Dim i As Integer
    con.Execute "DELETE FROM PrintCVs"
    strRptCondition = ""
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    'First Check Age....
    '0=Anything;1=Equal;2=More than;3=Less than
    If cmbAge.ListIndex = 0 Then
        'Do Nothing....
    ElseIf cmbAge.ListIndex = 1 Then
        strCondition = " WHERE (DATEDIFF(yyyy,DOB,getDate())=" & Val(txtAge) & ")"
        strRptCondition = "(DateDiff('yyyy',{CVs.DOB},CurrentDate)=" & Val(txtAge) & ")"
    ElseIf cmbAge.ListIndex = 2 Then
        strCondition = " WHERE (DATEDIFF(yyyy,DOB,getDate())>" & Val(txtAge) & ")"
        strRptCondition = "(DateDiff('yyyy',{CVs.DOB},CurrentDate)>" & Val(txtAge) & ")"
    ElseIf cmbAge.ListIndex = 3 Then
        strCondition = " WHERE (DATEDIFF(yyyy,DOB,getDate())<" & Val(txtAge) & ")"
        strRptCondition = "(DateDiff('yyyy',{CVs.DOB},CurrentDate)<" & Val(txtAge) & ")"
    End If
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    'Second Area....
    Dim strAreaList As String
    If lstArea.Selected(0) Then
        'No Criteria for Area....
    Else
        For i = 1 To lstArea.ListCount - 1
            If lstArea.Selected(i) Then
                strAreaList = strAreaList & "'" & lstArea.List(i) & "',"
            End If
        Next
        If strAreaList = "" Then
            'No Criterial
        Else
            strAreaList = Left(strAreaList, Len(strAreaList) - 1)
            If strCondition <> "" Then
                strCondition = strCondition & " " & cmbCriteria(0).Text & " (Area IN(" & strAreaList & "))"
                strRptCondition = strRptCondition & " " & cmbCriteria(0).Text & " ({CVs.Area} IN[" & strAreaList & "])"
            Else
                strCondition = " WHERE (Area IN(" & strAreaList & "))"
                strRptCondition = " ({CVs.Area} IN[" & strAreaList & "])"
            End If
        End If
    End If
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    'Third Gender....
     If cmbGender.ListIndex = 0 Then
        'No Criteria for Gender
    Else
        If strCondition <> "" Then
            strCondition = strCondition & " " & cmbCriteria(1).Text & " (Sex='" & cmbGender.Text & "')"
            strRptCondition = strRptCondition & " " & cmbCriteria(1).Text & " ({CVs.Sex}='" & cmbGender.Text & "')"
        Else
            strCondition = " WHERE (Sex='" & cmbGender.Text & "')"
            strRptCondition = " ({CVs.Sex}='" & cmbGender.Text & "')"
        End If
    End If
    
    'Short List
    If cmbShortList.ListIndex = 0 Then
        'No Criteria for Gender
    Else
        If strCondition <> "" Then
            strCondition = strCondition & " " & cmbCriteria(1).Text & " (SelectionType=" & cmbShortList.ListIndex & ")"
            strRptCondition = strRptCondition & " " & cmbCriteria(1).Text & " ({CVs.SelectionType}=" & cmbShortList.ListIndex & "')"
        Else
            strCondition = " WHERE (SelectionType=" & cmbShortList.ListIndex & ")"
            strRptCondition = " ({CVs.SelectionType}=" & cmbShortList.ListIndex & ")"
        End If
    End If
    '/\/\/\/\/\/
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    'Fourth Academic ....
    Dim strAcaList As String
    If lstAca.Selected(0) Then
        'No Criteria
    Else
        For i = 1 To lstAca.ListCount - 1
            If lstAca.Selected(i) Then
                strAcaList = strAcaList & "'" & lstAca.List(i) & "',"
            End If
        Next
        If strAcaList = "" Then
            'No Criteria
        Else
            strAcaList = Left(strAcaList, Len(strAcaList) - 1)
            If strCondition <> "" Then
                strCondition = strCondition & " " & cmbCriteria(2).Text & " (CVID IN(SELECT CVID FROM CVAcademicQualifications WHERE Diploma IN(" & strAcaList & ")))"
            Else
                strCondition = " WHERE (CVID IN(SELECT CVID FROM CVAcademicQualifications WHERE Diploma IN(" & strAcaList & ")))"
            End If
            con.Execute "INSERT INTO PrintCVs SELECT CVID FROM CVAcademicQualifications WHERE Diploma IN(" & strAcaList & ")"
            If strRptCondition <> "" Then
                strRptCondition = strRptCondition & cmbCriteria(2).Text & " {PrintCVs.CVID}<>''"
            End If
        End If
    End If
    
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    'Fifth Computer....
    Dim strCompList As String
    If lstComputer.Selected(0) Then
        'No Criteria
    Else
        For i = 1 To lstComputer.ListCount - 1
            If lstComputer.Selected(i) Then
                strCompList = strCompList & "'" & lstComputer.List(i) & "',"
            End If
        Next
        If strCompList = "" Then
            'No Criteria
        Else
            strCompList = Left(strCompList, Len(strCompList) - 1)
            If strCondition <> "" Then
                strCondition = strCondition & " " & cmbCriteria(3).Text & " (CVID IN(SELECT CVID FROM CVComputerLiteracy WHERE Diploma IN(" & strCompList & ")))"
            Else
                strCondition = " WHERE (CVID IN(SELECT CVID FROM CVComputerLiteracy WHERE Diploma IN(" & strCompList & ")))"
            End If
            'con.Execute "INSERT INTO PrintCVs SELECT CVID FROM CVComputerLiteracy WHERE Diploma IN(" & strAcaList & ") AND CVID NOT IN(SELECT CVID FROM PrintCVs)"
        End If
    End If
    
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    'Designations
    Dim strDesigList As String
    If lstDesignations.Selected(0) Then
        'No Criteria
    Else
        For i = 1 To lstDesignations.ListCount - 1
            If lstDesignations.Selected(i) Then
                strDesigList = strDesigList & "'" & lstDesignations.List(i) & "',"
            End If
        Next
        If strDesigList = "" Then
            'No Criteria
        Else
            strDesigList = Left(strDesigList, Len(strDesigList) - 1)
            If strCondition <> "" Then
                strCondition = strCondition & " " & cmbCriteria(4).Text & " (CVID IN(SELECT CVID FROM CVDesignations WHERE Designation IN(" & strDesigList & ")))"
            Else
                strCondition = " WHERE (CVID IN(SELECT CVID FROM CVDesignations WHERE Designation IN(" & strDesigList & ")))"
            End If
            'con.Execute "INSERT INTO PrintCVs SELECT CVID FROM CVComputerLiteracy WHERE Diploma IN(" & strAcaList & ") AND CVID NOT IN(SELECT CVID FROM PrintCVs)"
        End If
    End If
        
    Call RefreshList(strCondition)
    FraAdvSearch.Visible = False
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub ShowInactive(m As Menu)

    Dim Cond As String
    If m.Checked Then
        Cond = " Where Active=1"
    Else
        Cond = ""
    End If
    
    m.Checked = Not m.Checked
    Call RefreshList(Cond)
    
End Sub

Private Sub cmdCloseAdvSearch_Click()
    FraAdvSearch.Visible = False
End Sub

Private Sub cmdOpt_Click()

    With cmdOpt
        Call CheckMenu
        Me.PopUpMenu mnuPop, , .Left, .Top
    End With
    
End Sub

Private Sub cmdReset_Click()
    Call ResetAdvanceSearchOptions
End Sub

Private Sub cmdShowAdvSearch_Click()
    FraAdvSearch.Visible = True
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    'MainForm.mnuShowInactive.Checked = False
End Sub

Private Sub LV_ColumnClick1(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
       If LV.SortOrder = lvwAscending Then
          LV.SortOrder = lvwDescending
          ColumnHeader.Icon = "Up"
       Else
          LV.SortOrder = lvwAscending
          ColumnHeader.Icon = "Down"
       End If
    Else
     LV.ColumnHeaderIcons = Nothing
     Set LV.ColumnHeaderIcons = ImageList2
     LV.SortKey = ColumnHeader.Index - 1
     LV.SortOrder = lvwAscending
     ColumnHeader.Icon = "Down"
    End If
    
    On Error Resume Next
    LV.SelectedItem.EnsureVisible

End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    'If UserHasAccess("EditEmployee") Then
        Call EditItem(IIf(Left$(LV.SelectedItem.key, 3) = "DPT", "DPT", "EMP"))
    'End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Form_Load()

    mnuPop.Visible = False
    cmbfield.ListIndex = 0
    Call RefreshList
    
    Call ResetAdvanceSearchOptions
    
End Sub

Private Sub ResetAdvanceSearchOptions()
    
    'Advanced Search.....
    Dim i As Integer
    cmbAge.ListIndex = 0
    For i = 0 To cmbCriteria.count - 1
        cmbCriteria(i).ListIndex = 0
    Next
    
    lstArea.Clear
    lstArea.AddItem "Any Area"
    Call AddToCombo(lstArea, "DISTINCT Area", "CVs", , True)
    lstArea.Selected(0) = True
    cmbGender.ListIndex = 0
    
    lstAca.Clear
    lstAca.AddItem "Any Education"
    Call AddToCombo(lstAca, "Description", "AcaCourses", , True)
    lstAca.Selected(0) = True
    
    lstComputer.Clear
    lstComputer.AddItem "Any Course"
    Call AddToCombo(lstComputer, "Description", "ComputerCourses", , True)
    lstComputer.Selected(0) = True
    
    
    lstDesignations.Clear
    lstDesignations.AddItem "Any Designation"
    Call AddToCombo(lstDesignations, "DISTINCT Designation", "CVDesignations", , True)
    lstDesignations.Selected(0) = True
    
    cmbGender.ListIndex = 0
    cmbShortList.ListIndex = 0
    
End Sub

Private Sub Form_Resize()
     On Error Resume Next
     'fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
     
     Dim TitleBottom As Long
    
     TitleBottom = 20 'fTitle.Top + fTitle.Height
     
     LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 450)
     
     cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
     cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
     Frame1.Top = Me.ScaleHeight - Frame1.Height - 50
     
     FraAdvSearch.Move (LV.Width - FraAdvSearch.Width) / 2, (LV.Height - FraAdvSearch.Height) / 2
 
End Sub

Private Sub cmdSearch_Click()

    Dim StartIdx As Integer
    
    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
        If LV.SelectedItem.Index = LV.ListItems.count Then
            StartIdx = 1
        Else
            StartIdx = LV.SelectedItem.Index
        End If
    End If

    Dim fldValue As String

    For i = StartIdx To LV.ListItems.count
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
       End If
    
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtValue)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
       
    Next

    If i = LV.ListItems.count + 1 Then
        MsgBox "Search Information not FOUND...", vbInformation
        Set LV.SelectedItem = Nothing
    End If
    
Exit Sub

    On Error Resume Next

    If cmbfield.ListIndex = 0 Then
        For i = 1 To LV.ListItems.count
            If LV.ListItems(i).key Like "*" & txtValue & "*" Then
                LV.ListItems(i).Selected = True
                Exit For
            End If
        Next
    Else
        For i = 1 To LV.ListItems.count
            If LCase(LV.ListItems(i).SubItems(cmbfield.ListIndex) & "") Like "*" & LCase(txtValue) & "*" Then
                LV.ListItems(i).Selected = True
                Exit For
            End If
        Next
    End If
        

End Sub



Public Sub EditItem(Cond As String)

    Dim f As Form
    If Cond = "DPT" Then
        Set f = New NewDepartment
        f.TableName = "Departments"
        f.Tag = Right$(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3)
        f.add = False
        f.Editrec
        f.Show 1
        
    Else
'        If getDBPassword("EditEmployee") Then
            Set f = New frmNewCV
            Load f
            Call f.EditCV(LV.SelectedItem.key)
'        End If
    End If
    'Call RefreshList
    
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If showAllDept = True Then
            showAllDept = False
        Else
            showAllDept = True
        End If
        Call RefreshList
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   
    If Button <> 2 Then Exit Sub
   
'    Call CheckMenu
    PopUpMenu mnuPop
    
End Sub

Private Sub mnuappForm_Click()
    Call PrintAppForm
End Sub

Private Sub mnuConvertToEmployee_Click()
    Dim f As New NewEmployee
    f.TableName = "Employees"
    f.add = True
    Load f
    Call f.AddNewEmployee("", False, LV.SelectedItem.key)
End Sub

Private Sub mnuedit_Click()
    
    Call EditItem(mnuEdit.Tag)
    
End Sub

Private Sub mnuEmpLeaveBalance_Click()

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalance.rpt")
    SelFormula = "{VEmpLeaveBalance.DeptID}<>'VDRS' and {VEmpLeaveBalance.EmpID}<>''"
    SelFormula = SelFormula & IIf(mnuShowInactive.Checked, "", " and {VEmpLeaveBalance.Active}=True ")
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub mnuF1_Click()
    PrintEmpListDeptWise ("{Vemp.AttendanceGate}='F1' AND {VEmp.Active}=TRUE")
End Sub

Private Sub mnuF2_Click()
    PrintEmpListDeptWise ("{Vemp.AttendanceGate}='F2' AND {VEmp.Active}=TRUE")
End Sub

Private Sub mnuF3_Click()
    PrintEmpListDeptWise ("{Vemp.AttendanceGate}='F3' AND {VEmp.Active}=TRUE")
End Sub

Public Sub mnuNewDept_Click()

    Dim f As New NewDepartment
    f.TableName = "Departments"
    f.lblCatID = GetNextDeptID("")
    f.add = True
    f.Show 1
    Call RefreshList
    
End Sub

Private Sub mnuInterviewQuestions_Click()

    Dim f As New frmCVInterviewQuestions
    Load f
    f.ShowMe (LV.SelectedItem.key)
    
End Sub

Public Sub mnunewEmp_Click()

    Dim f As New frmNewCV
    Load f
    f.AddNewCV
    Call RefreshList
    
End Sub

Public Sub RefreshList(Optional strCondition As String)

    Dim rs As New ADODB.Recordset
   
    Dim ITM As ListItem
   
    Dim age As String, strStatus As String
     
    With rs
        .Open "SELECT CVID,Name,FName,Religion,Sex,DOB,InterviewDT,Diploma,TotalExperience,SelectionType,strfullname FROM VCVs " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![CVID], "  • " & ![CVID])
        
             ITM.ListSubItems.add(, , !Name).ToolTipText = !Name
             ITM.ListSubItems.add(, , Format(![InterviewDT], "dd-MMM-yyyy")).ForeColor = TheColor
             age = DateDiffInText2(![DOB])
             ITM.ListSubItems.add(, , age).ForeColor = TheColor
             
             ITM.ListSubItems.add(, , ![Diploma] & "").ToolTipText = ""
             ITM.ListSubItems.add(, , Round(Val(![TotalExperience] & ""), 2)).ForeColor = TheColor
            If IsNull(!SelectionType) Then
                strStatus = "-"
            ElseIf Val(!SelectionType & "") = 1 Then
                strStatus = "Reject"
            ElseIf Val(!SelectionType & "") = 2 Then
                strStatus = "Short List"
            ElseIf Val(!SelectionType & "") = 3 Then
                strStatus = "Future Short"
            End If
             ITM.ListSubItems.add(, , strStatus).ForeColor = TheColor
             ITM.ListSubItems.add(, , ![strFullName] & "").ForeColor = TheColor
    
             
                     
            .MoveNEXT
        Loop
        .Close

    End With
    Set rs = Nothing

End Sub
 
Public Sub MakeInactive(Optional NoMessage As Boolean = False)

    Dim Act As Boolean
    Dim SQL As String
    Dim theKey As String
    Dim DCDate As Date
    theKey = LV.SelectedItem.key
        
    If mnuMake.Caption = "&Make This Employee InActive" Then
        con.Execute "UPDATE Employees SET active=0 WHERE EmpID='" & theKey
    Else
        con.Execute "UPDATE Employees SET active=1 WHERE EmpID='" & theKey
    End If
    
'    If LV.SelectedItem.ForeColor = vbBlack Then Act = True Else Act = False
'
'    Dim DisContOn As String
'    If Act Then DisContOn = "NULL" Else DisContOn = "'" & Format(Date, "dd-MMM-yyyy") & "'"
'
'
'    If Left$(theKey, 3) = "DPT" Then
'
'        Dim Childs As Integer
'        Dim rs As New ADODB.Recordset
'        With rs
'            .Open "Select Count(EmpID) From Employees Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'", con, adOpenForwardOnly, adLockReadOnly
'            If IsNull(.Fields(0)) Then Childs = 0 Else Childs = .Fields(0)
'            .Close
'        End With
'        Set rs = Nothing
'
'        If MsgBox("This Will Affect All The Employees Of This Department." & _
'         vbNewLine & "There Are " & Childs & _
'         " Employees In This Department." & vbNewLine & "Do You Want To Continue?", vbQuestion + vbYesNo) = vbNo Then
'            Exit Sub
'        End If
'        con.BeginTrans
'
'
'        con.Execute "Update Departments set Active=" & Abs(Act) & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
'        con.Execute "Update Employees Set Discontinuedon=" & DisContOn & ",Active=" & Abs(Act) & " Where DeptID='" & Right(theKey, Len(theKey) - 3) & "'"
'        con.CommitTrans
'    Else
'        If Not Act Then
'            If MsgBox("This Is Usually Done When An Employee Discontinues. Do You Want To Continue ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
'
'            DCDate = InputBox("Plz Enter Date[mm/dd/yy] To Discontinue.")
'            If Not IsDate(DCDate) Then
'                MsgBox "Invalid Date.", vbCritical
'                Exit Sub
'            End If
'            con.Execute "Update Employees Set DiscontinuedOn='" & DCDate & "',Active=" & Abs(Act) & " Where EmpID='" & theKey & "'"
'        Else
'
'            DCDate = InputBox("Plz Enter Date[mm/dd/yy] To ReJoin.", "Re Joining After Discontinuation.")
'            If Not IsDate(DCDate) Then
'                MsgBox "Invalid Date.", vbCritical
'                Exit Sub
'            End If
'            con.Execute "Update Employees Set JoinDate='" & DCDate & "',DiscontinuedOn=Null,Active=" & Abs(Act) & " Where EmpID='" & theKey & "'"
'
      '  End If
        
        
'       con.Execute "INSERT INTO Employees(EmpID,DisjoinDT) VALUES()"
'        If Not Act Then
'            With Payslip
'                .cmbEmp.Text = LV.SelectedItem.SubItems(1)
'                .cmbDept.Enabled = False
'                .cmbEmp.Enabled = False
'                .Show
'            End With
'        End If
        
        
   ' End If
    
    Call RefreshList
    Dim Cond As String

End Sub

Private Sub CheckMenu()
    
    
    
   With Me
         
        If LV.ListItems.count = 0 Then
            .mnuEdit.Enabled = False
            
             .mnuNewEmp.Enabled = True
            
        Else
           
        End If
    End With

End Sub


Public Sub PrintAppForm()

    'With cr1
    '    .ReportFileName = RptPath & "\ApplicationForm.rpt"
    '    .Connect = ""
    '    .SelectionFormula = ""
    '    .Formulas(0) = ""
    '    .Formulas(1) = "Company='" & company & "'"
    '    .Action = 1
    'End With
    
    Screen.MousePointer = vbHourglass
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAppForm.rpt")
    rpt.FormulaFields(1).Text = "'" & company & "'"
    f.ShowReport " ", rpt
    Screen.MousePointer = vbDefault

End Sub

Public Sub PrintEmpList(Optional WithExternalSalary As Boolean = False, Optional Transport As Boolean = False)

    '    With cr1
    '        .ReportFileName = RptPath & "\EmpList.rpt"
    '        .Connect = ConnectStr
    '
    '        Dim SelFormula As String
            If mnuShowInactive.Checked Then
                SelFormula = IIf(Transport, " {VEmp.Transport}=true", "")
            Else
                SelFormula = "{VEmp.Active}=true " & IIf(Transport, " and {VEmp.Transport}=true", "")
            End If
        
    '        .SelectionFormula = SelFormula
    '
    '        .Formulas(0) = "External=" & WithExternalSalary
    '
    '        .Action = 1
        
    '    End With

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpList.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        End If
    Next
    
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault


End Sub

Public Sub PrintEmpListWithoutSal(Optional WithDOB As Boolean = False, Optional SelFormula As String, Optional ReportCaption As String = "Employee List")

        Dim TransCondition As String

        If mnuShowInactive.Checked Then
            TransCondition = IIf(Transport, " {VEmp.Transport}=true", "")
        Else
            TransCondition = "{VEmp.Active}=true " & IIf(Transport, " and {VEmp.Transport}=true", "")
        End If
        
        If TransCondition <> "" Then
            If SelFormula <> "" Then
                SelFormula = TransCondition & " and " & SelFormula
            Else
                SelFormula = TransCondition
            End If
        End If
        
        
'
'        .SelectionFormula = SelFormula
'
'        .Formulas(0) = "Company='" & company & "'"
'
'        .Action = 1
'
'    End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpListWithoutSal.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@BirthDate}" Then
        FormulaField.Text = WithDOB
    ElseIf FormulaField.Name = "{@RptCaption}" Then
        FormulaField.Text = "'" & ReportCaption & "'"
    End If
Next


f.ShowReport SelFormula & " ", rpt

Screen.MousePointer = vbDefault

End Sub


Public Sub PrintEmpCard(Optional EmpId As String)

    Dim Cond As String
    
    If EmpId = "" Then
        Cond = "{Employees.Active}=True"
    Else
        Cond = "{Employees.EmpID}='" & Trim(Replace(EmpId, "»", "")) & "' and {Employees.Active}=True"
    End If
        
    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpCards.rpt")
    f.ShowReport Cond & " ", rpt
    
    Screen.MousePointer = vbDefault
        
End Sub


Private Sub mnuempRefresh_Click()
    Call RefreshList
End Sub

Private Sub mnuMake_Click()
    
    Dim rs As New ADODB.Recordset
    Dim TempEmp As Boolean
    With rs
        .Open "Select TempDept  From Departments Where DeptID=(Select DeptID From Employees Where EmpID='" & LV.SelectedItem.key & ")", con, adOpenForwardOnly, adLockReadOnly
        
        TempEmp = .Fields(0)
        
        .Close
    End With
    Set rs = Nothing
    
'    If TempEmp = False Then
'        MsgBox "This Option Is Not Available." & vbNewLine & "Use 'Discontinue Employee' From Employees.", vbInformation
'        Exit Sub
'    End If
    
    If getPassword Then
        Call MakeInactive
    End If
    
End Sub

Private Sub mnuPrint_Click()
    Call PrintEmpList
End Sub

Private Sub mnuPrintAllBarcode_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    frmPrint.Show 1

End Sub

Private Sub mnuPrintAllCard_Click()
    Call PrintEmpCard
End Sub

Private Sub mnuPrintEmList_Click()
    
'    If Len(Lv.SelectedItem.Text) = 2 Then
'        PrintEmpListDeptWise ("{VEmp.Deptid} like '" & Lv.SelectedItem & "*'")
'
'    ElseIf Len(Lv.SelectedItem.Text) = 5 Then
'        PrintEmpListDeptWise ("{Vemp.Deptid} like '" & Lv.SelectedItem & "*'")
'    End If
    Call PrintEmpListDeptWise("{VEmp.DeptID}='" & LV.SelectedItem.ListSubItems(3).Tag & "' AND {VEmp.Active}")
    
End Sub

Private Sub PrintEmpListDeptWise(SelFormula As String)

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpList.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        End If
    Next
    
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault

End Sub


Private Sub mnuPrintEmpListCF_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpCardFields.rpt")
    Dim f As New frmPrevRpt
    
    Dim SelFormula As String
    
    SelFormula = IIf(mnuShowInactive.Checked, " ", "{VEmp.Active}=True ")
    
    f.ShowReport SelFormula, rpt
    
End Sub

Private Sub mnuPrintEmpListSal_Click()

    PrintEmpListWithoutSal

End Sub

Private Sub mnuPrintEmpListSalWDOB_Click()
    PrintEmpListWithoutSal True, , "Employee List (With DOB)"
End Sub

Private Sub mnuPrintEmpTrans_Click()
    Call PrintEmpList(False, True)
End Sub

Private Sub mnuPrintEmpWithLunch_Click()
    PrintEmpListWithoutSal False, "{Vemp.Lunch}=True", "Mess Members"
End Sub

Private Sub mnuPrintUnionMembers_Click()
    PrintEmpListWithoutSal False, "{Vemp.UnionFund}=True", "Union Members"
End Sub



Private Sub mnuPrintEmployeeListcodewise_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmpListCodewise.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VEmp.Active}=TRUE", rpt
    
End Sub

Private Sub mnuPrintCV_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\CVWithInterviewQA.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VCVs.CVID}='" & LV.SelectedItem.key & "'", rpt
End Sub

Private Sub mnuPrintList_Click()
    
    Dim i As Integer
    con.Execute "DELETE FROM PrintCVs"
    For i = 1 To LV.ListItems.count
        con.Execute "INSERT INTO PrintCVs(CVID) VALUES('" & LV.ListItems(i).key & "')"
    Next
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CVs.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt

End Sub

Private Sub mnuRefresh_Click()
    RefreshList
End Sub


Private Sub txtValue_GotFocus()

    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue)

End Sub
