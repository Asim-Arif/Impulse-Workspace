VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form EmpLedger 
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   10762
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   6165
      TabIndex        =   12
      Top             =   285
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
      TextForeColor   =   12582912
      ListIndex       =   -1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9330
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EmpLedger.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EmpLedger.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   3225
      TabIndex        =   3
      Top             =   285
      Width           =   2925
      _ExtentX        =   5159
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
      TextForeColor   =   12582912
      ListIndex       =   -1
   End
   Begin VB.Frame FLV 
      Height          =   7725
      Left            =   30
      TabIndex        =   0
      Top             =   615
      Width           =   11910
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   195
         TabIndex        =   14
         Top             =   7215
         Width           =   11670
         Begin MSForms.CommandButton cmdSettlement 
            Height          =   375
            Left            =   390
            TabIndex        =   23
            Top             =   60
            Visible         =   0   'False
            Width           =   1395
            ForeColor       =   0
            Caption         =   "Settlement"
            PicturePosition =   327683
            Size            =   "2461;661"
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdClose 
            Height          =   375
            Left            =   10635
            TabIndex        =   22
            Top             =   60
            Width           =   1035
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "1826;661"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrint 
            Height          =   375
            Left            =   9570
            TabIndex        =   21
            Top             =   60
            Width           =   1035
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "1826;661"
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintForSA 
            Height          =   375
            Left            =   6105
            TabIndex        =   20
            Top             =   60
            Visible         =   0   'False
            Width           =   1395
            ForeColor       =   0
            Caption         =   "Print Loans"
            PicturePosition =   327683
            Size            =   "2461;661"
            Accelerator     =   110
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintLeaves 
            Height          =   375
            Left            =   7545
            TabIndex        =   19
            Top             =   60
            Visible         =   0   'False
            Width           =   1395
            ForeColor       =   0
            Caption         =   "Print Leaves"
            PicturePosition =   327683
            Size            =   "2461;661"
            Accelerator     =   76
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintLeaveLedger 
            Height          =   375
            Left            =   7530
            TabIndex        =   18
            Top             =   60
            Visible         =   0   'False
            Width           =   1395
            ForeColor       =   0
            Caption         =   "Print L. Ledger"
            PicturePosition =   327683
            Size            =   "2461;661"
            Accelerator     =   103
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintEmpList 
            Height          =   375
            Left            =   4680
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   1395
            ForeColor       =   0
            Caption         =   "Print Emps"
            PicturePosition =   327683
            Size            =   "2461;661"
            Accelerator     =   69
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintTerm 
            Height          =   375
            Left            =   1830
            TabIndex        =   16
            Top             =   60
            Visible         =   0   'False
            Width           =   1395
            ForeColor       =   0
            Caption         =   "Print Term."
            PicturePosition =   327683
            Size            =   "2461;661"
            Accelerator     =   109
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintEntrants 
            Height          =   375
            Left            =   3255
            TabIndex        =   15
            Top             =   60
            Visible         =   0   'False
            Width           =   1395
            ForeColor       =   0
            Caption         =   "Print Entr."
            PicturePosition =   327683
            Size            =   "2461;661"
            Accelerator     =   116
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   450
         Left            =   30
         TabIndex        =   6
         Top             =   7185
         Visible         =   0   'False
         Width           =   4410
         Begin VB.Label lblTotalPresents 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1395
            TabIndex        =   10
            Top             =   150
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Presents :"
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
            Left            =   45
            TabIndex        =   9
            Top             =   165
            Width           =   1320
         End
         Begin VB.Label lblTotalAbs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3645
            TabIndex        =   8
            Top             =   150
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total Absents :"
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
            Left            =   2340
            TabIndex        =   7
            Top             =   165
            Width           =   1260
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7050
         Left            =   75
         TabIndex        =   11
         Top             =   180
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   12435
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Day"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "In Time"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Out Time"
            Object.Width           =   1535
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Hours"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Late Hours"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Over Time"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Short Leaves"
            Object.Width           =   4628
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Short Loan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Long Loan"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   7050
         Left            =   75
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   12435
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1958
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Day"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   1139
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Chk-In"
            Object.Width           =   1296
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Brk-Out"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Brk-In"
            Object.Width           =   1085
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Chk-Out"
            Object.Width           =   1429
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Hrs"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "OT In"
            Object.Width           =   1191
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "OT Out"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Late Hours"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Over Time"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Short Leaves"
            Object.Width           =   4366
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "Short Loan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "Long Loan"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   105
      TabIndex        =   5
      Top             =   285
      Width           =   1440
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2540;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   1575
      TabIndex        =   4
      Top             =   285
      Width           =   1620
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2857;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   105
      TabIndex        =   2
      Top             =   45
      Width           =   10230
      BackColor       =   11517387
      VariousPropertyBits=   8388627
      Caption         =   $"EmpLedger.frx":0564
      Size            =   "18045;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   10410
      TabIndex        =   1
      Top             =   255
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuAbsentLeave 
         Caption         =   "Change Attendance To Absent /Leave"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewSTLoan 
         Caption         =   "New Short Term Advance"
      End
      Begin VB.Menu mnuNewLTLoan 
         Caption         =   "New Long Term Advance"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuManualOT 
         Caption         =   "Manual Over Time Feeding"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "EmpLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalPresents As Double, TotalAbsents As Double, TotalLeaves As Double, TotalDays As Double, TotalHolidays As Double
Dim PrintForSA As Boolean

Public m_bShowFake As Boolean

Private Sub cmbMonth_Change()
    Call cmbMonth_Click
End Sub

Private Sub cmbMonth_Click()
    AddEmployees
End Sub

Private Sub cmbYear_Change()
    Call cmbYear_Click
End Sub

Private Sub cmbYear_Click()
    AddEmployees
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    PrintForSA = True
'    On Error GoTo ERR
'    Screen.MousePointer = vbHourglass
    Dim DTAtt As Date
    
    con.Execute "DELETE FROM PrintEmpLedger"
    
    If LV.Visible Then
'        For i = 1 To LV.ListItems.Count
'            With LV.ListItems(i)
'
'                If .Text <> "" Then
'                    DT = .Text
'                End If
'
'                con.Execute "Insert into  PrintEmpLedger (DT,Status,InTime,OutTime,Hrs,LateHours,OverTime,ShortLeave,EmpID)" & _
'                "values ('" & DT & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" _
'                & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & .SubItems(8) & "','" & cmbEmp.ID & "')"
'
'            End With
'        Next

        '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
        'Fazooooooooooooooool Coding Starts Here
        '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
        Dim SecondIn As String
        Dim SecondOut As String
        Dim dHrs As Double
        Dim dPHrs As Double
        For i = 1 To LV.ListItems.count
            dHrs = 0
            dPHrs = 0
            With LV.ListItems(i)
                If .Text <> "" Then
                    DT = .Text
                End If
            
                SecondIn = ""
                SecondOut = ""
                                
                If .SubItems(5) <> "" Then
                    dHrs = Val(.SubItems(5)) + (Val(Right(.SubItems(5), 2)) / 60)
                End If
                If .ListSubItems(3).Tag <> "" Then
                    dPHrs = Val(.ListSubItems(3).Tag) '+ (Val(Right(.ListSubItems(3).Tag, 2)) / 60)
                End If
                
                If i < LV.ListItems.count Then
                    If LV.ListItems(i + 1).Text = "" Then
                
                        SecondIn = LV.ListItems(i + 1).SubItems(3)
                        SecondOut = LV.ListItems(i + 1).SubItems(4)
                        
                        If LV.ListItems(i).SubItems(5) <> "" Then
                            dHrs = dHrs + Val(LV.ListItems(i + 1).SubItems(5)) + Val(Right(LV.ListItems(i + 1).SubItems(5), 2)) / 60
                        End If
                        dPHrs = dPHrs + Val(LV.ListItems(i + 1).ListSubItems(3).Tag)
                    End If
                End If
            
            
                dHrs = Round(dHrs, 2)
                
            
                con.Execute "INSERT INTO PrintEmpLedger(DT,Status,InTime,OutTime,SecondInTime,SecondOutTime,Hrs,LateHours,OverTime,ShortLeave,EmpID,PayableHrs,OTHrsOther)" & _
                 "Values ('" & DT & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" & SecondIn & "','" & SecondOut & "','" _
                 & getHrsMin(dHrs, True) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & .SubItems(8) & "','" & cmbEmp.ID & "','" & getHrsMin(dPHrs, True) & "'," & Val(.ListSubItems(4).Tag) & ")"
             
                If SecondIn <> "" Then i = i + 1
            End With
        Next i
        
        '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
        'Fazooooooooooooooool Coding Ends Here
        '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
    Else
        For i = 1 To LV2.ListItems.count
            dPHrs = 0
            With LV2.ListItems(i)
                If .Text <> "" Then
                    DT = .Text
                End If
                If .ListSubItems(3).Tag <> "" Then
                    dPHrs = Val(.ListSubItems(3).Tag) '+ (Val(Right(.ListSubItems(3).Tag, 2)) / 60)
                End If
'                con.Execute "Insert into  PrintEmpLedger (DT,Status,InTime,OutTime,SecondInTime," & _
                "SecondOutTime,Hrs,OTInTime,OTOutTime,LateHours,OverTime,ShortLeave,EmpID,PayableHrs)" & _
                " Values ('" & DT & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" _
                & .SubItems(7) & "','" & .SubItems(8) & "','" & .SubItems(9) & "','" & .SubItems(10) & _
                "','" & .SubItems(11) & "','" & .SubItems(12) & "','" & .SubItems(13) & "','" & .SubItems(14) & "','" & cmbEmp.ID & "','" & getHrsMin(dPHrs, True) & "')"
                
                 con.Execute "INSERT INTO PrintEmpLedger (DT,Status,InTime,OutTime,SecondInTime," & _
                "SecondOutTime,Hrs,OTInTime,OTOutTime,LateHours,OverTime,ShortLeave,EmpID,PayableHrs,OTHrsOther)" & _
                " VALUES ('" & DT & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" _
                & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & .SubItems(8) & _
                "','" & .SubItems(9) & "','" & .SubItems(10) & "','" & .SubItems(11) & "','" & .SubItems(12) & "','" & cmbEmp.ID & "','" & getHrsMin(dPHrs, True) & "'," & Val(.ListSubItems(4).Tag) & ")"
            End With
        Next
    End If
    Dim myDT_Last As Date
    myDT_Last = GetSingleDateValue("MAX(DT)", "PrintEmpLedger")
    
    con.Execute "UPDATE T1 SET T1.OTHrs=T2.OTHrs,T1.LateHrs=T2.LateHrs FROM PrintEmpLedger T1 INNER JOIN dbo.OverTime_Datewise_Fn('" & myDT_Last & "','" & cmbEmp.ID & "') T2 ON T1.EmpID=T2.EmpID AND T1.DT=T2.DT"
    con.Execute "UPDATE T1 SET T1.LateHrs_Original=T2.LateComingHrs FROM PrintEmpLedger T1 INNER JOIN dbo.LateComingHrs_Datewise_Fn('" & myDT_Last & "') T2 ON T1.EmpID=T2.EmpID AND T1.DT=T2.DT"
    
    'With cr1
    '    .ReportFileName = RptPath & "\EmpLedger.rpt"
    '    '.DataFiles(0) = rptpath & "\Payroll.mdb"
    '    '.Connect = ConnectStr
    '    .Formulas(0) = "FromTo='For " & cmbMonth & ", " & cmbYear & ".'"
    '    .Formulas(1) = "EmpNAme='Ledger Of " & cmbEmp.Text & "'"
    '    .Formulas(2) = "TotalLateHrs='" & LV.ListItems(LV.ListItems.Count).ListSubItems(5) & "'"
    '    .Formulas(3) = "TotalOverTime='" & LV.ListItems(LV.ListItems.Count).ListSubItems(6) & "'"
    '    .Formulas(4) = "TotalAbsents=" & TotalAbsents & ""
    '    .Formulas(5) = "TotalPresents=" & TotalPresents & ""
    '    .Formulas(6) = "Company='" & strCompany & "'"
    '    .Action = 1
    '    .PageZoomNext
    'End With
    
    
    

    '''''''''''''''''''''''''' Print The Ledger For Production Employees '''''''''''''''''''''''''''''''''
    
    If LV2.Visible Then
        PrintEmpLedgerWithLunchTime (PrintForSA)
        Exit Sub
    End If


    '''''''''''''''''''''''''' Print The Ledger For Staff Employees '''''''''''''''''''''''''''''''''


    Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLedger1.rpt")

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & strCompany & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "'For " & cmbMonth & ", " & cmbYear & ".'"
    ElseIf FormulaField.Name = "{@EmpName}" Then
        FormulaField.Text = "'" & cmbEmp.Text & "'"
    ElseIf FormulaField.Name = "{@TotalLateHrs}" Then
        FormulaField.Text = "'" & LV.ListItems(LV.ListItems.count).SubItems(5) & "'"
    ElseIf FormulaField.Name = "{@TotalOverTime}" Then
        FormulaField.Text = "'" & LV.ListItems(LV.ListItems.count).SubItems(6) & "'"

    ' Use Total Present Formula as Total Hrs at the moment
    ElseIf FormulaField.Name = "{@TotalPresents}" Then
        FormulaField.Text = "'" & getHrsMin(GetTotalHrs) & "'"
    ElseIf FormulaField.Name = "{@TotalPayableHrs}" Then
        FormulaField.Text = GetTotalPHrs
    ElseIf FormulaField.Name = "{@TotalAbsents}" Then
        FormulaField.Text = TotalAbsents
    End If
Next


'rpt.SQLQueryString = Sql
rpt.RecordSelectionFormula = "{VEmp.EmpID}='" & cmbEmp.ID & "'"
'rpt.PrintOut

f.ShowReport "{VEmp.EmpID}='" & cmbEmp.ID & "' ", rpt

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrintEmpList_Click()
Dim LastDT As Date
    Dim FirstDT As Date
    
    FirstDT = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
    LastDT = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0)
    
    'SelFormula = "dateVar DT; {VEmp.JoinDate}<=#" & LastDT & "# AND LEFT({VEmp.empid},4)<>'VDRS' AND ({VEmp.DiscontinuedOn}>=#" & FirstDT & "# OR ToText({VEmp.DiscontinuedOn})='')"
    SelFormula = "{VEmp.JoinDate}<=#" & LastDT & "# AND LEFT({VEmp.empid},4)<>'VDRS' AND ({VEmp.DiscontinuedOn}>=#" & FirstDT & "# OR ToText({VEmp.DiscontinuedOn})='')"
    
    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpListNew.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@ForMonth}" Then
            FormulaField.Text = "'" & cmbMonth & "-" & cmbYear & "'"
        ElseIf FormulaField.Name = "{@ForDT}" Then
            FormulaField.Text = "#" & LastDT & "#"
        End If
    Next
    
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrintForSA_Click()
    
    Dim SvrDt As Date
    SvrDt = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0)
    
    Screen.MousePointer = vbHourglass
    con.Execute "Delete From PrintSTBalance"
    con.Execute "Delete From PrintLTBalance"
    con.Execute "Insert Into PrintSTBalance Select EmpID,Sum(Amount) as STBalance From Advances " & _
     "Where type=0 and DT<='" & SvrDt & "' and Month(DT)=" & Month(SvrDt) & " and Year(DT)=" & year(SvrDt) & " Group By EmpID"
     
     
    'con.Execute "Insert Into PrintLTBalance Select EmpID,Sum(Amount)," & _
     "(Select Sum(AmtClrd) From AmtCleared Where EmpID=Advances.EmpID AND DT<='" & SvrDt & _
     "') From Advances Where Type=1 AND DT<='" & SvrDt & "' Group By EmpID"
     
     con.Execute "Insert Into PrintLTBalance Select EmpID,Sum(Amount)," & _
     "(Select PrevLTLoan-LongTerm From MonthlySalaries Where EmpID=Advances.EmpID AND Month(DT)=" & Month(SvrDt) & _
     " AND Year(DT)=" & year(SvrDt) & " AND IsForSA=0) From Advances Where Type=1 AND DT<='" & SvrDt & "' Group By EmpID"
     
     Call Sleep(1000)
    
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLoanBalance.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@ForMonth}" Then
            FormulaField.Text = "'" & cmbMonth & " - " & cmbYear & "'"
        End If
    Next

    f.ShowReport Cond & "{VLongTermBalance.empid}<>'' AND {VLongTermBalance.deptid}<>'VDRS' ", rpt
    Screen.MousePointer = vbDefault

'    PrintForSA = True
'    cmdPrint_Click
'    PrintForSA = False
End Sub

Private Sub cmdPrintLeaveLedger_Click()
    On Error GoTo err
    
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpMonthlyLeaveLedger.rpt")
    rpt.FormulaFields(1).Text = "' For " & cmbMonth & " - " & cmbYear & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "Month({Leaves.DT})=" & cmbMonth.ListIndex + 1 & " AND Year({Leaves.DT})=" & cmbYear & " AND ({Leaves.CL}=1 or {Leaves.SL}=1 or {Leaves.AL}=1 or {Leaves.CPL}=1 or {Leaves.LWP}=1 or {Leaves.ML}=1) ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrintLeaves_Click()

    On Error GoTo err
    Dim cmd As New ADODB.Command
    With cmd
        .CommandTimeout = 120
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_EmpLeaveBalances"
        .Parameters("@DT").Value = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
        .Execute
    End With
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalanceMonthly.rpt")
    rpt.FormulaFields(16).Text = "'" & cmbMonth & " - " & cmbYear & "'"
    Dim f As New frmPrevRpt
    f.ShowReport " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub cmdRefresh_Click()
    
    On Error GoTo err
        
'    If Left(cmbEmp.ID, 4) = "VMIS" And command$ <> "FromVBIDE" Then
'       Dim cmd As New ADODB.command
'        With cmd
'            .ActiveConnection = con
'            .CommandType = adCmdStoredProc
'            .CommandText = "SP_SaveMyViewer"
'            .Parameters("@UserName").value = CurrentUserName
'            .Parameters("@ComputerName").value = strComputerName
'            .Parameters("@DT").value = GetServerDate(True)
'            .Parameters("@TypeOfView").value = "Emp Ledger"
'            .Execute
'        End With
'        MsgBox "Run Time Error.", vbCritical
'        Exit Sub
'    End If
    
    
    If m_bShowFake Then
    
        Dim rs As New ADODB.Recordset
        Dim bDataExist As Boolean
        Dim bSalarySheetFinal As Boolean
        With rs
            .Open "Select Count(*) From EmpTimesFake Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear.Text, con, adOpenForwardOnly, adLockReadOnly
            If Val(.Fields(0)) = 0 Then
                bDataExist = False
            Else
                bDataExist = True
            End If
            .Close
            If bDataExist = False Then
                .Open "Select Count(*) From MonthlySalaries Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear, con, adOpenForwardOnly
                If Val(.Fields(0)) = 0 Then
                    bSalarySheetFinal = False
                Else
                    bSalarySheetFinal = True
                End If
            End If
        End With
        Set rs = Nothing
        
        If bDataExist = False Then
            If bSalarySheetFinal Then
                MsgBox "Fake Attendance will be Calculated Now.", vbInformation
                Call AddFakeAttendance
            Else
                MsgBox "Salary Sheet need to be Finalized for this Feature.", vbInformation
                Exit Sub
            End If
        End If
        
        
        
    End If
    
    If cmbEmp.MatchFound Then
        
        Emptype = con.Execute("Select EmpType From Employees Where EmpID='" & cmbEmp.ID & "'").Fields(0)
        Call Refresh2ShiftLedger
'        If Emptype = 2 Then
'            Call Refresh2ShiftLedger
'        Else
'            Call RefreshLedger
'        End If
'
    End If
        
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub AddFakeAttendance()

    On Error GoTo err
    
    con.Execute "Insert Into EmpTimesFake Select EmpID,DT,InTime,OutTime,Hrs,PayableHrs,ManualEntry From EmpTimes Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear.Text
    
    Dim rs As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim dHrsDiff As Double
    Dim dHrsLeft As Double
    Dim dHrs1 As Double, dHrs2 As Double
    dHrs1 = 1.5
    Dim EmpIDs As String
    Dim InTime As String
    
    With rs
        rsMain.Open "Select EmpID,SDays,FakeWorkingHrs From MonthlySalaries Where BSal<>FakeSalary AND Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear, con, adOpenForwardOnly, adLockReadOnly
        
        Do Until rsMain.EOF
        
            dHrsDiff = Val(rsMain!SDays & "") - Val(rsMain!FakeWorkingHrs & "")
            dHrsLeft = dHrsDiff
            .Open "Select NewID() As Row,* From EmpTimesFake Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear & " AND EmpID='" & rsMain!EmpId & "' AND Round(PayableHrs,1)>3.5 Order By Row", con, adOpenDynamic, adLockOptimistic
            
            If dHrs1 * .RecordCount < dHrsDiff Then
                dHrs1 = 2
                If dHrs1 * .RecordCount < dHrsDiff Then
                    dHrs1 = 2.5
                    If dHrs1 * .RecordCount < dHrsDiff Then
                        EmpIDs = EmpIDs & rsMain!EmpId & ","
                        GoTo ComeHere
                    End If
                End If
            End If
            
            Do Until .EOF
                InTime = !InTime
                InTime = DateAdd("H", dHrs1, InTime)
                
                dHrsLeft = dHrsLeft - dHrs1
                !InTime = Format(InTime, "HH:NN")
                !Hrs = !Hrs - dHrs1
                !PayableHrs = !PayableHrs - dHrs1
                .MoveNEXT
                If Round(dHrsLeft, 1) <= 0 Then GoTo ComeHere
            Loop
ComeHere:
            .Close
            rsMain.MoveNEXT
        Loop
        rsMain.Close
    End With
    Set rsMain = Nothing
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSettlement_Click()

    On Error GoTo err
    Dim SFor As String
    SFor = " "
    If cmbEmp.MatchFound = False Then
        SFor = " "
    Else
        If cmbEmp.ID <> "0" Then
            SFor = " AND {Employees.EmpID}='" & cmbEmp.ID & "' "
        End If
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptFinalSett1.rpt")
    Dim frm As New frmPrevRpt
    frm.ShowReport "{Employees.DiscontinuedOn} >= #10/01/2004# " & SFor, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()

    m_bShowFake = False
    
    mnuPop.Visible = False

    cmbMonth.AddItem "Jan"
    cmbMonth.List(0, 1) = 1
    cmbMonth.AddItem "Feb"
    cmbMonth.List(1, 1) = 2
    cmbMonth.AddItem "Mar"
    cmbMonth.List(2, 1) = 3
    cmbMonth.AddItem "Apr"
    cmbMonth.List(3, 1) = 4
    cmbMonth.AddItem "May"
    cmbMonth.List(4, 1) = 5
    cmbMonth.AddItem "Jun"
    cmbMonth.List(5, 1) = 6
    cmbMonth.AddItem "Jul"
    cmbMonth.List(6, 1) = 7
    cmbMonth.AddItem "Aug"
    cmbMonth.List(7, 1) = 8
    cmbMonth.AddItem "Sep"
    cmbMonth.List(8, 1) = 9
    cmbMonth.AddItem "Oct"
    cmbMonth.List(9, 1) = "10"
    cmbMonth.AddItem "Nov"
    cmbMonth.List(10, 1) = 11
    cmbMonth.AddItem "Dec"
    cmbMonth.List(11, 1) = 12
    
    Dim CYear As Long
    CYear = Format(GetServerDate, "yyyy")
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i


    cmbMonth.ListIndex = Val(Format$(GetServerDate, "MM")) - 1
    cmbYear = Format$(GetServerDate, "yyyy")



    DtTo = GetServerDate
    DtFrom = DateAdd("M", -1, GetServerDate)
    cmbDept.ListHeight = 6500
    cmbEmp.ListHeight = 6500
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    'cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    
    
End Sub
Private Sub cmbdept_matched()
    Call AddEmployees
End Sub

Private Sub AddEmployees()
    
    cmbEmp.ClearVals
    If cmbYear.MatchFound = False Then
        'MsgBox "Select Year.", vbInformation
        Exit Sub
    End If

    If cmbMonth.MatchFound = False Then
        'MsgBox "Select Year.", vbInformation
        Exit Sub
    End If
    
    'cmbEmp.AddItem "<All Employees>", "0"
    Dim LastDT As Date
    Dim FirstDT As Date
    
    FirstDT = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
    LastDT = DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0)
    
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", " Where JOINDate<='" & LastDT & "' AND (DiscontinuedOn>'" & FirstDT & "' OR DiscontinuedOn Is Null)"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and JOINDate<='" & LastDT & "' AND (DiscontinuedOn>'" & FirstDT & "' OR DiscontinuedOn Is Null)"
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    'fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    Dim TopMargin As Integer
    TopMargin = cmbEmp.Top + cmbEmp.Height + 30

    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
 
    LV.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
    LV2.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
 
    FButs1.Move LV.Left + LV.Width - FButs1.Width, LV.Top + LV.Height
 
    fButs2.Move LV.Left, LV.Top + LV.Height
 
End Sub



Private Sub RefreshLedger()

    LV2.Visible = False
    LV.Visible = True

    TotalAbsents = 0
    TotalPresents = 0
    TotalLeaves = 0
    TotalHolidays = 0

    'On Error GoTo ERR
    
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim rsShifts As New ADODB.Recordset
    Dim ITM As ListItem
    Dim pos As Integer
    Dim LateHours1 As Double
    Dim LateHours2 As Double
    
    Dim DT As Date
    Dim DTs As New Collection
    Dim LastDay As Integer
    Dim TheColor As Long
    Dim iDay As Integer
    Dim OverTime As String
    Dim ShortTerm As String
    Dim LongTerm As String, str As String
    
    Dim rs As New ADODB.Recordset, rsTime As New ADODB.Recordset
    
    'Get Holidays Into DTs Collection Object.
    
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    LastDay = DateAdd("m", 1, DT) - DT
    LV.ListItems.Clear
    


    '''''''''''''''''' This Will Update The LateHrs & OverTime ''''''''''
    'con.Execute "EXEC SP_CalcOvertimeEmpWise '" & DT & "','" & cmbEmp.ID & "'"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    With rs
        .Open "Select DAY(DT),Description from Holidays Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " Order by DT", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            'DTs.Add .Fields(0).value, CStr(.Fields(0).value)
            DTs.add .Fields(0).Value & "  {" & .Fields(1).Value & "}", .Fields(0).Value & ""
            
            .MoveNEXT
        Loop
        TotalHolidays = .RecordCount
        .Close
            
        'Get Records
        .Open "Select * from VEmpLedger Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and EmpID='" & cmbEmp.ID & "' Order by DT", con, adOpenStatic, adLockReadOnly
    
    TotalDays = LastDay
    
    Dim i As Long
    
    
    For o = 1 To LastDay
                
        If o - 1 = Day(Date) And Month(DT) = Month(Date) Then
            TotalDays = Day(Date)
            Exit For
        End If

        'On Error Resume Next 'To Trap the Error
        
        iDay = Day(DT)
        
        
        
        'If iDay = 9 Then Stop
        
        TheColor = &H800000
        Desc = ""
        
        If Weekday(DT) = 1 Then
            TheColor = vbRed
            TotalHolidays = TotalHolidays + 1
        Else
            For j = 1 To DTs.count
                If Val(DTs(j)) = iDay Then
                    TheColor = vbRed
                    Desc = Right(DTs(j), Len(DTs(j)) - 2)
                    TheColor = RGB(20, 150, 20)
                    TotalHol = TotalHol + 1
                    Bold = True
                    DTs.Remove (j)
                    Exit For
                End If
            Next j
        End If
        
        
        Set ITM = LV.ListItems.add(, Format(DT, "dd-MMM-yyyy"), Format(DT, "dd-MMM-yyyy"))
        
        i = i + 1
        
        ITM.ForeColor = TheColor
        
        ITM.ListSubItems.add(, , Format(DT, "dddd") & Desc).ForeColor = TheColor
        DT = DateAdd("d", 1, DT)
        rs.Find "AttDay=" & iDay, , adSearchForward, 1
        
        
            If .EOF Then
                For j = 2 To 7
                    ITM.SubItems(j) = "" '"N/A"
                Next j
            Else

                    If Format(![DT], "ddd") = "Sun" Or TheColor = RGB(20, 150, 20) Then
                        'itm.ListSubItems.Add(, , ![Attend] & "").ForeColor = TheColor
                        ITM.ListSubItems.add(, , "P" & "").ForeColor = TheColor
                        
                    Else
                        ITM.ListSubItems.add(, , IIf(Val(!Attendance & "") = 0, "A", IIf(Val(!Attendance & "") = 1, "P", IIf(Val(!Attendance & "") = 4, "HL", "L")))).ForeColor = TheColor
                        
                        If Val(!Attendance & "") = 1 Then
                            TotalPresents = TotalPresents + 1
                        ElseIf Val(!Attendance & "") = 2 Then
                            TotalLeaves = TotalLeaves + 1

                        End If
                    End If
                    
                    'Debug.Assert iDay <> 19
                    
                    If (LV.ListItems(i).ListSubItems(2) = "A" Or LV.ListItems(i).ListSubItems(2) = "L") And Format(![DT], "ddd") <> "Sun" Then
                    
                        ITM.ListSubItems.add , , ""
                        ITM.ListSubItems.add , , ""
                    Else


''''''''''''''''''''''''''''''''''''' To Get All The Time Intervals ''''''''''
                        
                        'rsTime.Open "SELECT * FROM EmpTimes WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & !DT & "'", con, adOpenStatic, adLockReadOnly
                        rsTime.Open "SELECT * FROM VEmpTimes1 WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & !DT & "'", con, adOpenStatic, adLockReadOnly
                        If rsTime.RecordCount > 0 Then
                            ITM.ListSubItems.add(, , rsTime![InTime] & "").ForeColor = TheColor
                            ITM.ListSubItems(3).Tag = rsTime!PayableHrs1 & ""
                            ITM.ListSubItems.add(, , rsTime![OutTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.add(, , getHrsMin(Val(rsTime!Hrs & ""))).ForeColor = TheColor
                            ITM.Tag = rsTime!Hrs & ""
                            rsTime.MoveNEXT
                        Else
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                        End If
                        
                        ITM.ListSubItems.add(, , getHrsMin(Val(!LateHours & ""), True)).ForeColor = TheColor
                        ITM.ListSubItems.add(, , getHrsMin(Val(!OverTimeHours & ""), True)).ForeColor = TheColor
                        str = con.Execute("Select dbo.FShortLeaves ('" & DateSerial(cmbYear, cmbMonth.ListIndex + 1, i) & "','" & cmbEmp.ID & "')").Fields(0) & ""
                        ITM.SubItems(8) = str

                        For j = 1 To rsTime.RecordCount - 1
                            Set ITM = LV.ListItems.add(, , "")
                            i = i + 1
'                            ITM.SubItems(3) = rsTime!InTime & ""
'                            ITM.SubItems(4) = rsTime!OutTime & ""
'                            ITM.SubItems(5) = getHrsMin(Val(rsTime!Hrs & ""))
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , rsTime!InTime & ""
                            ITM.ListSubItems(3).Tag = rsTime!PayableHrs1 & ""
                            ITM.ListSubItems.add , , rsTime!OutTime & ""
                            ITM.ListSubItems.add , , getHrsMin(Val(rsTime!Hrs & ""))
                            ITM.Tag = rsTime!Hrs & ""
                            
                
                            rsTime.MoveNEXT
                        Next
                        rsTime.Close
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    End If
            End If
        Next o
    End With
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
    ICO = ""
    
    'TotalAbsents = TotalDays - (TotalPresents + TotalLeaves + TotalHolidays)
    
    'Replace the TotalAbsents and TotalPresents with TotalDays and Hours
    
    TotalPresents = TotalDays - TotalHolidays
    TotalAbsents = TotalPresents * 8
    
    Exit Sub
err:
        MsgBox err.Number & vbNewLine & err.Description
        Exit Sub
        
End Sub



Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Or LV.ListItems.count = 0 Then Exit Sub
    PopUpMenu mnuPop
End Sub

Private Sub mnuAbsentLeave_Click()
    With frmLeaves
        .cmbEmp.Text = cmbEmp.Text
        .DTPicker1 = CDate(LV.SelectedItem.Text)
        .Show 1
    End With
End Sub

Private Sub mnuManualOT_Click()
    With frmManualOTFeeding
        .cmbEmp.Text = cmbEmp.Text
        .Show 1
    End With
End Sub

Private Sub mnuNewSTLoan_Click()

    With NewAdvanceShort
        .cmbDept.Text = cmbDept.Text
        .cmbEmp.ID = cmbEmp.ID
        .LongTerm = False
        .Show 1
    End With
    
    
End Sub

Private Sub mnuNewLTLoan_Click()
    
    With NewAdvanceLong
        .cmbDept.Text = cmbDept.Text
        .cmbEmp.ID = cmbEmp.ID
        .LongTerm = True
        .DTPicker1 = #6/30/2004#
        .txtDesc = "Previous Balance Of Fixed Loan"
        .Show 1
    End With
    
End Sub

Private Sub Refresh2ShiftLedger()

    LV2.Visible = True
    LV.Visible = False

    TotalAbsents = 0
    TotalPresents = 0
    TotalLeaves = 0
    'On Error GoTo ERR
    
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim rsShifts As New ADODB.Recordset
    Dim ITM As ListItem
    Dim pos As Integer
    Dim LateHours1 As Double
    Dim LateHours2 As Double
    
    Dim DT As Date
    Dim DTs As New Collection
    Dim LastDay As Integer
    Dim TheColor As Long
    Dim iDay As Integer
    Dim OverTime As String
    Dim ShortTerm As String
    Dim LongTerm As String, str As String
    
    Dim rs As New ADODB.Recordset, rsTime As New ADODB.Recordset
    
    Dim LeaveType As String
    
    
    TotalPresents = 0
    TotalHolidays = 0
    TotalAbsents = 0
    
    'Get Holidays Into DTs Collection Object.
    
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    LastDay = DateAdd("m", 1, DT) - DT
    LV2.ListItems.Clear


    '''''''''''''''''' This Will Update The LateHrs & OverTime ''''''''''
'        con.Execute "EXEC SP_CalcOvertimeEmpWise '" & DT & "','" & cmbEmp.ID & "'"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With rs
        .Open "Select DAY(DT),Description from Holidays Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " Order by DT", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            'DTs.Add .Fields(0).value, CStr(.Fields(0).value)
            DTs.add .Fields(0).Value & "  {" & .Fields(1).Value & "}", .Fields(0).Value & ""
            
            .MoveNEXT
        Loop
        .Close

        'Get Records
        .Open "Select * from VEmpLedger Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and EmpID='" & cmbEmp.ID & "' Order by DT", con, adOpenStatic, adLockReadOnly
        
    For i = 1 To LastDay
    
        TotalDays = i
        
'        If i - 1 = Day(Date) And Month(DT) = Month(Date) Then
'            TotalDays = i
'            Exit For
'        End If

        'On Error Resume Next 'To Trap the Error
        
        iDay = Day(DT)
        
        'If iDay = 9 Then Stop
        
        TheColor = &H800000
        Desc = ""
        
        If Weekday(DT) = 1 Then
            TheColor = RGB(20, 150, 20)
            TotalHolidays = TotalHolidays + 1
        Else
            For j = 1 To DTs.count
                If Val(DTs(j)) = iDay Then
                    TheColor = vbRed
                    Desc = Right(DTs(j), Len(DTs(j)) - 2)
                    TheColor = RGB(20, 150, 20)
                    TotalHol = TotalHol + 1
                    Bold = True
                    DTs.Remove (j)
                    Exit For
                End If
            Next j
        End If
        
        Set ITM = LV2.ListItems.add(, Format(DT, "dd-MMM-yyyy"), Format(DT, "dd-MMM-yyyy"))
        ITM.ForeColor = TheColor
        
        ITM.ListSubItems.add(, , Format(DT, "dddd") & Desc).ForeColor = TheColor
        DT = DateAdd("d", 1, DT)
        rs.Find "AttDay=" & iDay, , adSearchForward, 1
                
            If .EOF Then
            
                For j = 2 To 7
                    ITM.SubItems(j) = "" '"N/A"
                Next j
                
            Else
            
                    ' Determine The Leave Type
                    If ![CL] & "" Then
                        LeaveType = "( Casual Leave )"
                    ElseIf ![SL] & "" Then
                        LeaveType = "( Sick Leave )"
                    ElseIf ![AL] & "" Then
                        LeaveType = "( Annual Leave )"
                    ElseIf ![CPL] & "" Then
                        LeaveType = "( Compensatory Leave )"
                    ElseIf ![LWP] & "" Then
                        LeaveType = "( Leave Without Pay )"
                    ElseIf ![ML] & "" Then
                        LeaveType = "( Maternity Leave )"
                    ElseIf ![Sal] & "" Then
                        LeaveType = "( Special Approved Leave )"
                    End If
                                        

                    If Format(![DT], "ddd") = "Sun" Or TheColor = RGB(20, 150, 20) Then
                        'itm.ListSubItems.Add(, , ![Attend] & "").ForeColor = TheColor
                        ITM.ListSubItems.add(, , "P" & "").ForeColor = TheColor
                    Else
                        ITM.ListSubItems.add(, , IIf(Val(!Attendance & "") = 0, "A", IIf(Val(!Attendance & "") = 1, "P", IIf(Val(!Attendance & "") = 4, "HL", "L")))).ForeColor = TheColor
                        If Val(!Attendance & "") = 1 Then
                            TotalPresents = TotalPresents + 1
                        ElseIf Val(!Attendance & "") = 2 Then
                            TotalLeaves = TotalLeaves + 1
                        End If
                    End If
                    
                    
                    If (LV2.ListItems(i).ListSubItems(2) = "A" Or LV2.ListItems(i).ListSubItems(2) = "L") And Format(![DT], "ddd") <> "Sun" Then
                    
                        ITM.ListSubItems.add , , ""
                        ITM.ListSubItems.add , , ""
                        ITM.SubItems(14) = LeaveType
                        
                    Else


                    ''''''''''''''''''''''' To Get All The Time Intervals '''''''''''''''''''''''
                        If m_bShowFake Then
                            rsTime.Open "SELECT * FROM VEmpTimes2Fake WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & !DT & "' ORder By DT,EmpID,FirstEntryID ", con, adOpenForwardOnly, adLockReadOnly
                        Else
                            rsTime.Open "SELECT * FROM VEmpTimes2 WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & !DT & "' ORder By DT,EmpID,FirstEntryID ", con, adOpenForwardOnly, adLockReadOnly
                        End If
                        If rsTime.RecordCount > 0 Then
                            ITM.ListSubItems.add(, , rsTime![FirstInTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.add(, , rsTime![FirstOutTime] & "").ForeColor = TheColor
                            'ITM.ListSubItems.add(, , rsTime![FirstOutTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.add(, , rsTime![SecondInTime] & "").ForeColor = TheColor
                            'ITM.ListSubItems.add(, , rsTime![SecondInTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.add(, , rsTime![SecondOutTime] & "").ForeColor = TheColor
                            
                            ITM.ListSubItems.add(, , getHrsMin(Val(rsTime!Hrs & ""))).ForeColor = TheColor
                            
                            ITM.ListSubItems.add(, , rsTime![OTInTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.add(, , rsTime![OTOutTime] & "").ForeColor = TheColor
                            ITM.ListSubItems.add().ForeColor = TheColor
                            ITM.ListSubItems.add(, , getHrsMin(Val(rsTime!OTHrs & ""))).ForeColor = TheColor
                            
                            ITM.Tag = rsTime!Hrs & ""
                            ITM.ListSubItems(3).Tag = IIf(Val(rsTime!PayableHrs & "") > 8, 8, rsTime!PayableHrs & "")
                            ITM.ListSubItems(4).Tag = Val(!OTHrsOther & "")
                            rsTime.MoveNEXT
                        Else
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            ITM.ListSubItems.add , , ""
                            
                        End If
                        
                        'ITM.ListSubItems.add(, , getHrsMin(Val(!LateHours & ""), True)).ForeColor = TheColor
                        'ITM.ListSubItems.add(, , getHrsMin(Val(!OverTimeHours & ""), True)).ForeColor = TheColor
                        str = con.Execute("Select dbo.FShortLeaves ('" & DateSerial(cmbYear, cmbMonth.ListIndex + 1, i) & "','" & cmbEmp.ID & "')").Fields(0) & ""
                        
                        
                        
                        ITM.SubItems(12) = str

                        'For j = 1 To rsTime.RecordCount - 1
                        '    Set ITM = LV2.ListItems.Add(, , "")
                        '    ITM.SubItems(3) = rsTime!InTime & ""
                        '    ITM.SubItems(4) = rsTime!OutTime & ""
                        '    ITM.SubItems(5) = getHrsMin(Val(rsTime!Hrs & ""))
                        '    rsTime.MoveNext
                        'Next
                        
                        rsTime.Close
                        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    End If
            End If
        Next i
    End With
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
    ICO = ""
    
    
'    For i = 1 To LV2.ListItems.Count
'        If LV2.ListItems(i).Text <> "" And (LV2.ListItems(i).SubItems(2) = "" Or LV2.ListItems(i).SubItems(2) = "A") Then
'            TotalAbsents = TotalAbsents + 1
'        End If
'    Next
    
    'TotalAbsents = TotalDays - (TotalPresents + TotalLeaves + TotalHolidays)
    
    'Replace the TotalAbsents and TotalPresents with TotalDays and Hours
    
    ' at The moment -1 is applied to remove extra day ...
    ' to be  Checked after
    
    TotalPresents = TotalDays - TotalHolidays - 1
    TotalAbsents = TotalPresents * 8
    
    Dim lForeColor As Long
    For i = 1 To LV2.ListItems.count
        With LV2.ListItems(i)
            If .SubItems(2) = "" And .ForeColor <> RGB(20, 150, 20) Then
                lForeColor = vbRed
            ElseIf .SubItems(2) = "L" Then
                lForeColor = RGB(20, 220, 20)
            ElseIf .ForeColor = RGB(20, 150, 20) Then
                lForeColor = RGB(20, 150, 20)
            Else
                lForeColor = LV.ForeColor
            End If
            Set ITM = LV2.ListItems(i)
            ITM.ForeColor = lForeColor
            For j = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(j).ForeColor = lForeColor
            Next
        End With
    Next
    
    Exit Sub
err:
        MsgBox err.Number & vbNewLine & err.Description
        Exit Sub

End Sub

Private Sub PrintEmpLedgerWithLunchTime(Optional ForSA As Boolean)


    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
'    If MsgBox("Print Over Time Hours?", vbYesNo + vbQuestion) = vbYes Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLedgerWLTOT.rpt")
'    Else
'        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLedgerWithLunchTime.rpt")
'    End If
    'Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLedgerWithLunchTime.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'For " & cmbMonth & ", " & cmbYear & ".'"
        ElseIf FormulaField.Name = "{@EmpName}" Then
            FormulaField.Text = "'" & cmbEmp.Text & "'"
        ElseIf FormulaField.Name = "{@TotalLateHrs}" Then
            FormulaField.Text = "'" & LV2.ListItems(LV2.ListItems.count).SubItems(5) & "'"
        ElseIf FormulaField.Name = "{@TotalOverTime}" Then
            FormulaField.Text = "'" & LV2.ListItems(LV2.ListItems.count).SubItems(6) & "'"
        ElseIf FormulaField.Name = "{@TotalPresents}" Then
            FormulaField.Text = "'" & getHrsMin(GetTotalHrs) & "'"
        ElseIf FormulaField.Name = "{@TotalPayableHrs}" Then
            FormulaField.Text = GetTotalPHrs
        ElseIf FormulaField.Name = "{@TotalAbsents}" Then
            FormulaField.Text = TotalAbsents
        ElseIf FormulaField.Name = "{@ForSA}" Then
            FormulaField.Text = ForSA
        End If
    Next

    'SelFormula = "Year({Holidays.DT})=" & cmbYear
    'On Error Resume Next
    'SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
    'rpt.SQLQueryString = SQL
    'rpt.PrintOut
    
    f.ShowReport Cond & " ", rpt
    
    Screen.MousePointer = vbDefault

End Sub

Private Function GetTotalPHrs() As Double

    Dim total As Double
    If LV.Visible Then
        For i = 1 To LV.ListItems.count
            total = total + Val(LV.ListItems(i).ListSubItems(3).Tag)
        Next
    Else
        For i = 1 To LV2.ListItems.count
            Debug.Print Round(Val(LV2.ListItems(i).ListSubItems(3).Tag), 1) & "<" & Val(LV2.ListItems(i).ListSubItems(3).Tag)
            
            total = total + Round(Val(LV2.ListItems(i).ListSubItems(3).Tag), 2)
            
        Next
    End If
    
    GetTotalPHrs = total
    
End Function


Private Function GetTotalHrs() As Double
    Dim total As Double
    If LV.Visible Then
        For i = 1 To LV.ListItems.count
            total = total + Val(LV.ListItems(i).Tag)
        Next
    Else
        For i = 1 To LV2.ListItems.count
            total = total + Val(LV2.ListItems(i).Tag)
        Next
    End If
    
    GetTotalHrs = total
End Function


Private Sub cmdPrintTerm_Click()

    Dim DT As Date
    Dim DT1 As Date
    
    DT1 = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
    DT = DateAdd("M", -1, DT1)
    
    'SelFormula = "dateVar DT; {VEmp.JoinDate}<=#" & LastDT & "# AND LEFT({VEmp.empid},4)<>'VDRS' AND ({VEmp.DiscontinuedOn}>=#" & FirstDT & "# OR ToText({VEmp.DiscontinuedOn})='')"
    SelFormula = "Month({Employees.DiscontinuedOn})=" & Month(DT) & " AND Year({Employees.DiscontinuedOn})=" & year(DT) & " AND LEFT({Employees.empid},4)<>'VDRS'"
    
    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptTerminationsList.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        
        If FormulaField.Name = "{@ForDT}" Then
            FormulaField.Text = "#" & DT1 & "#"
        End If
    Next
    
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrintEntrants_Click()

    Dim DT As Date
    Dim DT1 As Date
    
    DT1 = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
    DT = DateAdd("M", -3, DT1)
    
    'SelFormula = "dateVar DT; {VEmp.JoinDate}<=#" & LastDT & "# AND LEFT({VEmp.empid},4)<>'VDRS' AND ({VEmp.DiscontinuedOn}>=#" & FirstDT & "# OR ToText({VEmp.DiscontinuedOn})='')"
    SelFormula = "Month({Employees.JoinDate})=" & Month(DT) & " AND Year({Employees.JoinDate})=" & year(DT) & " AND LEFT({Employees.empid},4)<>'VDRS'"
    
    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEntrantsList.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        
        If FormulaField.Name = "{@ForDT}" Then
            FormulaField.Text = "#" & DT1 & "#"
        End If
    Next
    
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
    
End Sub
