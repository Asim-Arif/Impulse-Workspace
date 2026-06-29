VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddBroSis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Non Dependant Family Members"
   ClientHeight    =   3150
   ClientLeft      =   7170
   ClientTop       =   5820
   ClientWidth     =   4905
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   90
      TabIndex        =   6
      Top             =   0
      Width           =   4755
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   2430
         TabIndex        =   17
         Top             =   2205
         Width           =   2265
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   15
         Top             =   2205
         Width           =   2340
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Text            =   "34603-2275342-5"
         Top             =   990
         Width           =   1440
      End
      Begin VB.CommandButton cmdSavenNew 
         Caption         =   "&Save && New"
         Height          =   405
         Left            =   540
         TabIndex        =   12
         Top             =   2580
         Width           =   1335
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
         TabIndex        =   3
         Top             =   1650
         Width           =   4620
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Save && &Close"
         Height          =   405
         Left            =   1935
         TabIndex        =   4
         Top             =   2580
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "C&lose"
         Height          =   405
         Left            =   3330
         TabIndex        =   5
         Top             =   2580
         Width           =   1335
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
         Format          =   41615363
         CurrentDate     =   36891
      End
      Begin MSForms.ComboBox cmbMaritalStatus 
         Height          =   315
         Left            =   1530
         TabIndex        =   19
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   1530
         TabIndex        =   18
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2460
         TabIndex        =   16
         Top             =   2010
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Education :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   75
         TabIndex        =   14
         Top             =   2010
         Width           =   810
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   28
         Left            =   3330
         TabIndex        =   11
         Top             =   765
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIC # :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   75
         TabIndex        =   10
         Top             =   795
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2460
         TabIndex        =   9
         Top             =   165
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   8
         Top             =   165
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   75
         TabIndex        =   7
         Top             =   1440
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmAddBroSis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal As Integer
Dim sName As String, sRelation As String, sNICNo As String
Dim DTBirth As Date, sAge As String, sMaritalStatus As String
Dim sEducation As String, sOccupation As String
Dim strNumber As String

Public Function ShowMe(ByRef strName As String, ByRef Relation As String, ByRef NICNo As String, ByRef MaritalStatus As String, ByRef Birth As Date, strAge As String, Education As String, Occupation As String) As Integer
    
    Me.Show 1
    
    strName = sName
    Relation = sRelation
    NICNo = sNICNo
    Birth = DTBirth
    MaritalStatus = sMaritalStatus
    strAge = sAge
    Education = sEducation
    Occupation = sOccupation
    
    ShowMe = RetVal
        
End Function

Private Sub cmdCancel_Click()
    RetVal = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()

    RetVal = 1
    
    sName = txts(0)
    sRelation = cmbRelation
    sNICNo = txts(1)
    sMaritalStatus = cmbMaritalStatus
    sEducation = txts(2)
    sOccupation = txts(4)
    DTBirth = DT
    sAge = txts(3)
    Unload Me
    
End Sub


Private Sub cmdSavenNew_Click()
    RetVal = 2
    
    sName = txts(0)
    sRelation = cmbRelation
    sMother = cmbMother
    DTBirth = DT
    sAge = txts(3)
    Unload Me
End Sub

Private Sub DT_Change()
    
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

    With cmbRelation
        .AddItem "Brother"
        .AddItem "Sister"
        .ListIndex = 0
    End With
    With cmbMaritalStatus
        .AddItem "Single"
        .AddItem "Married"
        .ListIndex = 0
    End With
    
End Sub
