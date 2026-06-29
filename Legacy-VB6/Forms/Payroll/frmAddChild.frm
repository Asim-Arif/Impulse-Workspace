VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddChild 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Family Member"
   ClientHeight    =   4155
   ClientLeft      =   7170
   ClientTop       =   5820
   ClientWidth     =   4965
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txtInstitute 
         Height          =   330
         Left            =   90
         TabIndex        =   24
         Top             =   2835
         Width           =   4560
      End
      Begin VB.CheckBox chkDependent 
         Caption         =   "Dependent"
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   3465
         Width           =   1125
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   1890
         TabIndex        =   5
         Top             =   1650
         Width           =   2760
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   3
         Top             =   990
         Width           =   1440
      End
      Begin VB.CommandButton cmdSavenNew 
         Caption         =   "Save && &New"
         Height          =   405
         Left            =   1260
         TabIndex        =   7
         Top             =   3465
         Width           =   1110
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   0
         Top             =   375
         Width           =   2340
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1650
         Width           =   1740
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save && Close"
         Height          =   405
         Left            =   2430
         TabIndex        =   8
         Top             =   3465
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   3555
         TabIndex        =   9
         Top             =   3465
         Width           =   1110
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   330
         Left            =   3315
         TabIndex        =   2
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   87752707
         CurrentDate     =   36891
      End
      Begin MSForms.ComboBox cmbClasses 
         Height          =   330
         Left            =   2430
         TabIndex        =   23
         Top             =   2205
         Width           =   2220
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3916;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label8 
         Caption         =   "Name of the Institute"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   90
         TabIndex        =   22
         Top             =   2565
         Width           =   1950
      End
      Begin VB.Label Label7 
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2430
         MouseIcon       =   "frmAddChild.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1980
         Width           =   465
      End
      Begin MSForms.ComboBox cmbEducationLevel 
         Height          =   330
         Left            =   90
         TabIndex        =   20
         Top             =   2205
         Width           =   2310
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4075;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbMaritalStatus 
         Height          =   315
         Left            =   1530
         TabIndex        =   4
         Tag             =   "MaritalStatus"
         Top             =   990
         Width           =   1785
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3149;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Marital Status "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   1530
         TabIndex        =   19
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1845
         TabIndex        =   18
         Top             =   1395
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Education Level :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   1980
         Width           =   1230
      End
      Begin MSForms.ComboBox cmbRelation 
         Height          =   330
         Left            =   2460
         TabIndex        =   1
         Top             =   375
         Width           =   2220
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3916;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   12164479
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   28
         Left            =   3330
         TabIndex        =   16
         Top             =   765
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIC # :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   15
         Top             =   795
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2460
         TabIndex        =   14
         Top             =   165
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   13
         Top             =   165
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   12
         Top             =   1440
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmAddChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal As Integer
Dim sName As String, sRelation As String, sNICNo As String
Dim DTBirth As Date, sAge As String, sMaritalStatus As String
Dim sClassName As String, sInstitute As String
Dim sEducation As String, sOccupation As String
Dim strNumber As String
Dim bDependent As Boolean

Public Function ShowMe(ByRef strName As String, ByRef Relation As String, ByRef NICNo As String, ByRef MaritalStatus As String, ByRef Birth As Date, strAge As String, Education As String, Occupation As String, Dependent As Boolean, ClassName As String, Institute As String) As Integer
     
    txts(0) = strName
    cmbRelation = Relation
    txts(1) = NICNo
    cmbMaritalStatus = MaritalStatus
    txts(4) = Occupation
    txts(3) = strAge
    DT = Birth
    cmbEducationLevel = Education
    cmbClasses = ClassName
    txtInstitute = Institute
    Me.Show 1
   
    strName = sName
    Relation = sRelation
    NICNo = sNICNo
    Birth = DTBirth
    MaritalStatus = sMaritalStatus
    strAge = sAge
    Education = sEducation
    Occupation = sOccupation
    Dependent = bDependent
    ClassName = sClassName
    Institute = sInstitute
    
    ShowMe = RetVal

End Function

Private Sub cmbEducationLevel_Change()
    
    Dim rs As New ADODB.Recordset
    If rs.State Then rs.Close
    rs.Open "select * from ClassesDetail where EducationLevel='" & cmbEducationLevel.Text & "'", con, adOpenForwardOnly, adLockReadOnly
    cmbClasses.Clear
    While Not rs.EOF
        cmbClasses.AddItem (rs.Fields("className"))
        rs.MoveNEXT
    
    Wend
     
End Sub

Private Sub cmbRelation_Change()
    Call cmbRelation_Click
End Sub

Private Sub cmbRelation_Click()
    
    If cmbRelation.ListIndex = -1 Then Exit Sub
    If cmbRelation.ListIndex > 3 Then
        chkDependent.Visible = True
        chkDependent.Value = vbChecked
    Else
        chkDependent.Visible = False
        chkDependent.Value = vbUnchecked
    End If
    If cmbMaritalStatus.ListIndex = -1 Then Exit Sub
    Select Case cmbRelation.ListIndex
        Case 2, 3, 4, 5
            cmbMaritalStatus.ListIndex = 1
        Case Else
            cmbMaritalStatus.ListIndex = 0
    End Select
End Sub

Private Sub cmdCancel_Click()
    RetVal = 0
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    If Not ValidateData Then
        Exit Sub
    End If
    RetVal = 1
    
    sName = txts(0)
    sRelation = cmbRelation
    sNICNo = txts(1)
    sMaritalStatus = cmbMaritalStatus
    sEducation = cmbEducationLevel.Text
    sClassName = cmbClasses.Text
    sInstitute = txtInstitute.Text
    sOccupation = txts(4)
    DTBirth = DT
    sAge = txts(3)
    bDependent = IIf(chkDependent.Value = vbChecked, True, False)
    
    Unload Me
    
End Sub

Private Function ValidateData() As Boolean

    If txts(0) = "" Then
        MsgBox "Invalid Name.", vbInformation
        Exit Function
    ElseIf cmbRelation.ListIndex = -1 Then
        MsgBox "Select Relation.", vbInformation
        Exit Function
    ElseIf cmbMaritalStatus.ListIndex = -1 Then
        MsgBox "Select Marital Status.", vbInformation
        Exit Function
    End If
    ValidateData = True
    
End Function

Private Sub cmdSavenNew_Click()
        
    If Not ValidateData Then
        Exit Sub
    End If
    
    RetVal = 2
    
    sName = txts(0)
    sRelation = cmbRelation
    sNICNo = txts(1)
    sMaritalStatus = cmbMaritalStatus
    sEducation = cmbEducationLevel.Text
    sClassName = cmbClasses.Text
    sInstitute = txtInstitute.Text
    sOccupation = txts(4)
    DTBirth = DT
    sAge = txts(3)
    bDependent = IIf(chkDependent.Value = vbChecked, True, False)
    
    Unload Me
    
End Sub

Private Sub Dt_Change()
    
    txts(3) = DateDiffInText(DT.Value)
'    Dim myDT As Date
'    Dim iYears As Integer, iMonths As Integer, iDays As Integer
'    myDT = dt.value
'    iYears = DateDiff("m", myDT, Date)
'    myDT = DateAdd("M", iYears, myDT)
'    iMonths = iYears Mod 12
'
'    iDays = DateDiff("d", myDT, Date)
'    If iDays < 0 Then
'        myDT = DateAdd("M", -1, myDT)
'        iDays = DateDiff("d", myDT, Date)
'
'        iMonths = iMonths - 1
'    End If
'
'    iYears = Int(iYears / 12)
'
'    txts(3) = iYears & " Years, " & iMonths & " Months, " & iDays & " Days."
    
End Sub

Private Sub Form_Load()
    
    DT = #1/1/1995#
    Call Dt_Change
    With cmbRelation
        .AddItem "Son"
        .AddItem "Daughter"
        .AddItem "Wife"
        .AddItem "Husband"
        .AddItem "Father"
        .AddItem "Mother"
        .ListIndex = 0
    End With
    
    With cmbMaritalStatus
        .AddItem "Single"
        .AddItem "Married"
        .AddItem "Widow"
        .AddItem "Widower"
        .AddItem "Divorced"
        .ListIndex = 0
    End With
    
    With cmbEducationLevel
        .AddItem ("Primary")
        .AddItem ("Middle")
        .AddItem ("Matric")
        .AddItem ("Intermediate")
        .AddItem ("Graducation")
        .AddItem ("Post Graducate")
    End With
    
End Sub

Private Sub Label7_Click()
    frmAddClasses.Show 1
End Sub

Private Sub txtInstitute_LostFocus()
    txtInstitute = StrConv(txtInstitute, vbProperCase)
End Sub

Private Sub txts_LostFocus(Index As Integer)
    txts(Index) = StrConv(txts(Index), vbProperCase)
End Sub
