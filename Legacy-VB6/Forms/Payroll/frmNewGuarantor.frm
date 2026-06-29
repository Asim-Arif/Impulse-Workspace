VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmNewGuarantor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guarantor"
   ClientHeight    =   6855
   ClientLeft      =   -18000
   ClientTop       =   1395
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FOptions 
      Height          =   3975
      Index           =   0
      Left            =   90
      TabIndex        =   23
      Top             =   1170
      Width           =   9150
      Begin VB.Frame Frame1 
         Caption         =   "Comments :"
         Height          =   2415
         Left            =   4560
         TabIndex        =   29
         Top             =   1425
         Width           =   4365
         Begin MSForms.TextBox TBox 
            Height          =   2085
            Index           =   7
            Left            =   60
            TabIndex        =   8
            Tag             =   "strComments"
            Top             =   240
            Width           =   4215
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "7435;3678"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Address :"
         Height          =   2415
         Left            =   165
         TabIndex        =   24
         Top             =   1425
         Width           =   4365
         Begin MSForms.TextBox TBox 
            Height          =   2085
            Index           =   8
            Left            =   60
            TabIndex        =   7
            Tag             =   "strAddress"
            Top             =   240
            Width           =   4215
            VariousPropertyBits=   -1467987941
            ScrollBars      =   2
            Size            =   "7435;3678"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Known By "
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   30
         Top             =   810
         Width           =   750
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   6
         Tag             =   "strKnownBy"
         Top             =   1035
         Width           =   8775
         VariousPropertyBits=   679495707
         Size            =   "15478;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cast"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   2385
         TabIndex        =   28
         Top             =   240
         Width           =   330
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   3
         Left            =   2340
         TabIndex        =   3
         Tag             =   "strCast"
         Top             =   465
         Width           =   2145
         VariousPropertyBits=   679495707
         Size            =   "3784;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N.I.C. # : (Nadra Format)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   27
         Top             =   225
         Width           =   1875
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   4
         Left            =   165
         TabIndex        =   2
         Tag             =   "strNICNo"
         Top             =   465
         Width           =   2145
         VariousPropertyBits=   679495707
         Size            =   "3784;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Phone :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   4530
         TabIndex        =   26
         Top             =   255
         Width           =   555
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   5
         Left            =   4515
         TabIndex        =   4
         Tag             =   "strPhoneNo"
         Top             =   465
         Width           =   2145
         VariousPropertyBits=   679495707
         Size            =   "3784;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cell :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   6705
         TabIndex        =   25
         Top             =   255
         Width           =   360
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   6
         Left            =   6690
         TabIndex        =   5
         Tag             =   "strCellNo"
         Top             =   465
         Width           =   2265
         VariousPropertyBits=   679495707
         Size            =   "3995;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.PictureBox PicSearch 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   105
      ScaleHeight     =   600
      ScaleWidth      =   7125
      TabIndex        =   17
      Top             =   6135
      Visible         =   0   'False
      Width           =   7125
      Begin VB.ComboBox cmbField 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmNewGuarantor.frx":0000
         Left            =   30
         List            =   "frmNewGuarantor.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   210
         Width           =   2835
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Left            =   2880
         TabIndex        =   18
         Top             =   210
         Width           =   2745
      End
      Begin VB.Label lblSearch 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By :                                           Search For:"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   21
         Top             =   0
         Width           =   5235
      End
      Begin MSForms.CommandButton cmdSearch 
         Height          =   360
         Left            =   5670
         TabIndex        =   20
         Top             =   180
         Width           =   1380
         Caption         =   "Search      "
         PicturePosition =   327683
         Size            =   "2434;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame fraNav 
      Caption         =   " Showing :   0 / 0 "
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
      Height          =   540
      Left            =   105
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   2505
      Begin VB.CommandButton cmdNav 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   705
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1815
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   200
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Father Name :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   10
      Left            =   4170
      TabIndex        =   22
      Top             =   240
      Width           =   1035
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   405
      Left            =   5415
      TabIndex        =   11
      Top             =   6225
      Width           =   1620
      Caption         =   "Save"
      PicturePosition =   327683
      Size            =   "2857;714"
      Accelerator     =   118
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TBox 
      Height          =   390
      Index           =   2
      Left            =   4140
      TabIndex        =   1
      Tag             =   "strFatherName"
      Top             =   450
      Width           =   4485
      VariousPropertyBits=   679495707
      Size            =   "7911;688"
      FontName        =   "Tahoma"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   390
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Tag             =   "strFullName"
      Top             =   450
      Width           =   4020
      VariousPropertyBits=   679495707
      Size            =   "7091;688"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   225
      Width           =   510
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   405
      Left            =   7110
      TabIndex        =   9
      Top             =   6225
      Width           =   1620
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2857;714"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Job"
      End
   End
   Begin VB.Menu mnupopup1 
      Caption         =   "popup1"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteQual 
         Caption         =   "Delete This Entry"
      End
   End
End
Attribute VB_Name = "frmNewGuarantor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelOption As Integer
Dim LastPoint As Integer

Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Dim dontShowList As Boolean

Public SqlSettings As String

Dim PicFileName As String
Dim ThumbFileName As String
Dim SigFileName As String

Public TableName As String
Dim EditEmp As Boolean

Dim EmpId As String
Dim WithEvents rsView As ADODB.Recordset
Attribute rsView.VB_VarHelpID = -1
Dim strName As String, Relation As String, NICNo As String, Birth As Date, strAge As String
Dim MaritalStatus As String, Education As String, ClassName As String, Institute As String, Occupation As String

Dim myVal As Integer
Dim LvChildIndex As Integer
Dim lEntryID As Long


Public Sub ShowMe()

    Me.Show 1
    
End Sub

Public Sub ViewRecords(Optional bShowInactive As Boolean = False)

    On Error GoTo err
   
    Call RefreshRecordSet(bShowInactive)
    
    cmdSave.Visible = False
    cmdOK.Visible = False
    cmdNext.Visible = False
    fraNav.Visible = True
    PicSearch.Visible = True
    
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub RefreshRecordSet(bShowInactive As Boolean)

    Set rsView = New ADODB.Recordset
    
    If bShowInactive Then
        rsView.Open "Select * From GuarantorList_Imported Order By intGuarantorId", con, adOpenForwardOnly, adLockReadOnly
    Else
        rsView.Open "Select * From GuarantorList_Imported Order By intGuarantorId", con, adOpenForwardOnly, adLockReadOnly
    End If
    If rsView.EOF = False Then rsView.MoveFirst
    
End Sub

Private Sub FillCmbs()

'    With cmbBloodGroup
'        .AddItem "--"
'        .AddItem "A+"
'        .AddItem "B+"
'        .AddItem "AB+"
'        .AddItem "O+"
'        .AddItem "A-"
'        .AddItem "B-"
'        .AddItem "AB-"
'        .AddItem "O-"
'    End With
'
    
End Sub


Private Sub chkShowInactive_Click()
    
End Sub

Private Sub cmbAKIWCML_Change()
    Call cmbAKIWCML_Click
End Sub

Private Sub cmbAKIWCML_Click()
    If cmbAKIWCML.ListIndex = 0 Then
        TBox(59) = ""
        TBox(59).Locked = True
    Else
        TBox(59) = ""
        TBox(59).Locked = False
    End If
End Sub

Private Sub cmbAllergy_Change()
    Call cmbAllergy_Click
End Sub

Private Sub cmbAllergy_Click()

    If cmbAllergy.ListIndex = 0 Then
        Label1(101).Visible = False
        TBox(17).Visible = False
        TBox(17) = ""
    Else
        Label1(101).Visible = True
        TBox(17).Visible = True
    End If
    
End Sub

Private Sub cmbDesignation_Change()
    Call cmbDesignation_Click
End Sub

Private Sub cmbDesignation_Click()

'    If cmbDesignation = "Driver" Then
'        TBox(12).Locked = False
'        TBox(13).Locked = False
'        DTLicenseUpto.Enabled = True
'        cmbLicenseTypes.Locked = False
'    Else
'        TBox(12).Locked = True
'        TBox(13).Locked = True
'        DTLicenseUpto.Enabled = False
'        cmbLicenseTypes.Locked = True
'    End If
'
'    TBox(12) = ""
'    TBox(13) = ""
'    DTLicenseUpto.value = Null
'    cmbLicenseTypes.ListIndex = -1
    
End Sub

Private Sub cmbEOBIL_Change()
    Call cmbEOBIL_Click
End Sub

Private Sub cmbEOBIL_Click()

    TBox(52) = ""
    TBox(10) = ""
    DTEOBIIss = Null
    DTMaturity = Null
    
    If cmbEOBIL.ListIndex = 0 Then
        cmbPensionableL.Visible = True
        cmbPensionableL.ListIndex = 0
        'TBox(52).Locked = True
        TBox(52).Left = DTEOBIIss.Left
        Label1(93).Caption = "Status"
        Label1(97).Caption = "EOBI #"
        Label1(98).Caption = "Card #"
        DTMaturity.Enabled = False
        DTEOBIIss.Enabled = False
        DTMaturity.Visible = False
        DTEOBIIss.Visible = False
        TBox(10).Visible = True
        If cmbPensionableL.Visible Then cmbPensionableL.SetFocus
    Else
        cmbPensionableL.Visible = False
        'TBox(52).Locked = False
        TBox(52).Left = cmbPensionableL.Left
        Label1(93).Caption = "EOBI #"
        Label1(97).Caption = "Issue Date"
        Label1(98).Caption = "Date Of Maturity"
        
        
        TBox(10).Visible = False
        If TBox(52).Visible Then TBox(52).SetFocus
    End If
    
End Sub

Private Sub cmbFlatL_Change()
    Call cmbFlatL_Click
End Sub

Private Sub cmbFlatL_Click()

    Dim bLocked As Boolean
    Dim StrText As String
    If cmbFlatL.ListIndex = 0 Then
        StrText = "N/A"
        bLocked = True
    Else
        StrText = ""
        bLocked = False
    End If
    
    For i = 60 To 62
        TBox(i) = StrText
        TBox(i).Locked = bLocked
    Next
    TBox(60) = 0
    
End Sub

Private Sub cmbGroupInsuranceL_Change()
    Call cmbGroupInsuranceL_Click
End Sub

Private Sub cmbGroupInsuranceL_Click()
    
    Dim bLocked As Boolean
    If cmbGroupInsuranceL.ListIndex = 0 Then
        bLocked = True
    Else
        bLocked = False
    End If
    For i = 56 To 56
        TBox(i).Locked = bLocked
        TBox(i) = ""
    Next
    lblNominee.Enabled = Not bLocked
    cmbNominees.Locked = bLocked
    cmbNominees.Clear
    TBox(11) = ""
    TBox(14) = ""
   
End Sub

Private Sub cmbMainUnit_matched()
    cmbDepartment.ClearVals
    cmbDepartment.AddVals con, "Name", "Departments", "DeptID", " Where SubDeptOf='" & cmbMainUnit.ID & "'"
End Sub

Private Sub cmbNominees_Change()
    Call cmbNominees_Click
End Sub

Private Sub cmbNominees_Click()
    If cmbNominees.ListIndex <> -1 Then
        TBox(14) = cmbNominees.List(cmbNominees.ListIndex, 1)
        TBox(11) = cmbNominees.List(cmbNominees.ListIndex, 2)
    End If
End Sub

Private Sub cmbPensionableL_Change()
    Call cmbPensionableL_Click
End Sub

Private Sub cmbPensionableL_Click()
    If cmbPensionableL.ListIndex = 0 Then
        Label1(98).Caption = "Remarks"
    Else
        Label1(98).Caption = "Card #"
    End If
End Sub

Private Sub cmbPymtmode_Change()
    Call cmbPymtmode_Click
End Sub

Private Sub cmbPymtmode_Click()

    If cmbPymtmode.ListIndex = 2 Then
        TBox(47) = ""
        TBox(5) = ""
        TBox(47).Locked = False
        TBox(5).Locked = False
    Else
        TBox(47) = "N/A"
        TBox(5) = "N/A"
        TBox(47).Locked = True
        TBox(5).Locked = True
    End If
    
End Sub

Private Sub cmbShare1_Change()
    cmbShare1_Click
End Sub

Private Sub cmbShare1_Click()
    If cmbShare1.ListIndex <> -1 Then
        cmbShare2.ListIndex = 100 - cmbShare1.ListIndex
    End If
End Sub

Private Sub cmbShare2_Change()
    Call cmbShare2_Click
End Sub

Private Sub cmbShare2_Click()
    If cmbShare2.ListIndex <> -1 Then
        cmbShare1.ListIndex = 100 - cmbShare2.ListIndex
    End If
End Sub

Private Sub cmbSSecurityL_Change()
    Call cmbSSecurityL_Click
End Sub

Private Sub cmbSSecurityL_Click()
    
    Dim StrText As String
    DTSSIss = Null
    If cmbSSecurityL.ListIndex = 0 Then
        StrText = "N/A"
        TBox(51).Locked = True
        TBox(53).Locked = True
        DTSSIss.Enabled = False
    Else
        StrText = ""
        TBox(51).Locked = False
        TBox(53).Locked = False
        DTSSIss.Enabled = True
    End If
    TBox(51) = StrText
    TBox(53) = StrText
    
    
End Sub

Private Sub cmbTaxDeductionL_Change()
    Call cmbTaxDeductionL_Click
End Sub

Private Sub cmbTaxDeductionL_Click()

    Dim bLocked As Boolean
    Dim StrText As String
    If cmbTaxDeductionL.ListIndex = 0 Then
        StrText = "N/A"
        bLocked = True
    Else
        StrText = ""
        bLocked = False
    End If
    
    For i = 48 To 49
        TBox(i) = StrText
        TBox(i).Locked = bLocked
    Next
    TBox(16) = 0
    TBox(16).Locked = bLocked
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
        
End Sub

Private Sub cmdCopyAddress_Click()
    
    TBox(40) = TBox(8)
         
End Sub

Private Sub cmdNav_Click(Index As Integer)

    With rsView
        Select Case Index
            Case 0          '''Move to First Record
                .MoveFirst
            Case 1          '''Move to Previous Record
                .MovePrevious
                If .BOF Then .MoveFirst
            Case 2   '''Move to Next Record
                If .EOF Then .MoveLast
                .MoveNEXT
               
            Case 3          '''Move to Last Record
                .MoveLast
        End Select
    End With
    
End Sub

Private Sub cmdNext_Click()

    If saveCV Then
        'StrTemp = cmbDept.Text
        Unload Me
        Load frmNewCV
        frmNewCV.AddNewCV
    End If
    
End Sub

Private Sub cmdOK_Click()
    If saveCV Then
        Unload Me
        If SqlSettings <> "" Then Unload frmSettings
    End If
End Sub

Private Sub cmdSave_Click()

    If saveGuarantor Then
        'EmpId = lblEmpID
        If lEntryID = 0 Then
            lEntryID = GetSingleLongValue("MAX(intGuarantorId)", "GuarantorList_Imported")
        End If
    End If

End Sub

Private Sub cmdSearch_Click()

    If cmbField.ListIndex = -1 Then Exit Sub
    If txtValue = "" Then Exit Sub
    
    Dim iCurrPos As Integer
    
    With rsView
        If .EOF Or .BOF Then
            
            .Find cmbField.Text & " like '%" & txtValue & "%'", , adSearchForward, 1
        Else
            iCurrPos = .AbsolutePosition
            .Find cmbField.Text & " like '%" & txtValue & "%'", iCurrPos, adSearchForward, 1
        End If
    End With
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    'Me.KeyPreview = True
    'Call FillCmbs
    EmpId = ""
    'Tbox(2).Text = "S/O"
    
End Sub

Private Function saveGuarantor() As Boolean

    On Error GoTo err

    Dim OpBalance As Double
    
    If Trim(TBox(0)) = "" Then
        MsgBox "Invalid Name.", vbInformation
        saveCV = False
        Exit Function
'    ElseIf TBox(4) = "" Then
'        MsgBox "Invalid NIC #", vbInformation
'        saveCV = False
'        Exit Function
    End If
    
    Call StartTrans(con)
    Dim StrSQL As String
    If lEntryID = 0 Then
        StrSQL = CreateInsertSQL("GuarantorList_Imported")
    Else
        StrSQL = CreateUpdateSQL("GuarantorList_Imported") & " WHERE intGuarantorId=" & lEntryID
    End If
       
    con.Execute StrSQL
     
    'con.Execute "UPDATE Employees SET"
    
    

TheEnd:
    saveGuarantor = True
    con.CommitTrans
    Exit Function
err:
    'If con.BeginTrans Then con.CommitTrans
    con.RollbackTrans
    MsgBox err.Description
    
End Function

Private Sub AddEmpPic(Optional iOption As Integer)

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        If iOption = 0 Then
            Image1.Picture = LoadPicture(Picfile)
            PicFileName = Picfile
            lblPic.Visible = False
        End If
    End If
    
End Sub

Private Sub Image1_DblClick()
    Call AddEmpPic
End Sub

Private Sub JoinDate_Change()
    TBox(9) = DateDiffInText(JoinDate.Value)
End Sub


Private Sub lblPic_DblClick()
    AddEmpPic
End Sub

Private Sub LVComputer_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVComputer.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVComputer.ListItems.Remove LVComputer.SelectedItem.Index
        End If
    End If
End Sub

Private Sub LVExp_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVExp.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVExp.ListItems.Remove LVExp.SelectedItem.Index
        End If
    End If
End Sub

Private Sub LVQ_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVQ.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVQ.ListItems.Remove LVQ.SelectedItem.Index
        End If
    End If

End Sub

Private Sub LVRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVRef.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
            LVRef.ListItems.Remove LVRef.SelectedItem.Index
        End If
    End If
End Sub

Private Sub rsView_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    With rsView
        If .EOF = False And .BOF = False Then
            fraNav.Caption = " Showing :   " & CStr(.AbsolutePosition) & " / " & CStr(.RecordCount) & " "
            'EmpId = !CVID & ""
            lEntryID = !intGuarantorId
            Call ShowGuarantorData
        End If
    End With
    
End Sub

Private Sub SO_Click(Index As Integer)

    On Error GoTo err
    Dim sSex As String
    If Index = 0 Then
        
        TBox(15) = "Male"
        TBox(2) = "S/O"
        TBox(1).SetFocus
        
    ElseIf Index = 1 Then
        
        TBox(15) = "Female"
        TBox(1).SetFocus
        TBox(2) = "D/O"
    ElseIf Index = 2 Then
        
        TBox(15) = "Female"
        TBox(1).SetFocus
        TBox(2) = "W/O"

    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub TBox_GotFocus(Index As Integer)
    
    TBox(Index).SelStart = 0
    TBox(Index).SelLength = Len(TBox(Index))
    
End Sub

Private Function GetNextItemID(CatID As String) As String
    Dim NextVal As String
    NextVal = GetMax("cast(Right(ItemID,2) as int)", "Items", "Where CatID ='" & CatID & "'")
    GetNextItemID = CatID & "-" & Format(Val(NextVal) + 1, "000")
End Function

Private Sub TBox_LostFocus(Index As Integer)
    If Index = 26 Then
        TBox(Index) = UCase(TBox(26))
    Else
        TBox(Index) = StrConv(TBox(Index), vbProperCase)
    End If
End Sub

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)
    If Index = 41 Then
        If Val(TBox(41)) < 5000 Then
            cmbSSecurityL.ListIndex = 1
        Else
            cmbSSecurityL.ListIndex = 0
        End If
    End If
End Sub



Public Sub EditGuarantor(p_lEntryID As Long, Optional bShowInactive As Boolean = False)

    On Error GoTo err
    lEntryID = p_lEntryID
    
    Dim rs As New ADODB.Recordset
    Dim iCurPos As Integer
    
    EditEmp = True
   
    Set rsView = New ADODB.Recordset
    
'    If bShowInactive Then
'        chkShowInactive.Value = vbChecked
'        'rsView.Open "SELECT EmpID,PrevCode,Name,FName,Designation,Religion,Sex FROM Employees ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
'    Else
'        chkShowInactive.Value = vbUnchecked
'        'rsView.Open "SELECT EmpID,PrevCode,Name,FName,Designation,Religion,Sex FROM Employees WHERE Active=1 ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
'    End If
    rsView.Open "SELECT * FROM GuarantorList_Imported ORDER BY intGuarantorId", con, adOpenForwardOnly, adLockReadOnly
    Call RefreshRecordSet(True)
    
    'If rsView.EOF = False Then rsView.MoveFirst   '
    iCurPos = rsView.AbsolutePosition
    'EmpId = strEmpID
    rsView.Find "intGuarantorId=" & p_lEntryID, iCurPos, adSearchForward, 1
    
    'Call ShowEmpData
    fraNav.Visible = True
    'cmdNext.Visible = False
    'cmdOK.Visible = False
    cmdSave.Move cmdCancel.Left - 2000, cmdCancel.Top
    Me.Show 1
   
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub ShowGuarantorData()
    
    On Error Resume Next
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM GuarantorList_Imported WHERE intGuarantorId=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each C In Controls
            If C.Tag <> "" Then
                'Debug.Assert C.Tag <> "CustCode"
                If TypeOf C Is MSForms.CheckBox Then
                    C.Value = Not .Fields(C.Tag)
                ElseIf TypeOf C Is ComboList.Usercontrol1 Then
                    C.ID = .Fields(C.Tag) & ""
              
                    'C.ID = .Fields(C.Tag) & ""
                ElseIf TypeOf C Is MSForms.Label Or TypeOf C Is VB.Label Then
                    C.Caption = .Fields(C.Tag) & ""
                ElseIf TypeOf C Is DTPicker Then
                    C.Value = .Fields(C.Tag) & ""
                ElseIf (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is VB.ComboBox) Then
                
                    If Right(C.Name, 1) = "L" Then
                        If Not IsNull(.Fields(C.Tag)) Then
                            C.ListIndex = Abs(.Fields(C.Tag))
                        Else
                            C.ListIndex = -1
                        End If
                    Else
                        If C.Name = "cmbShare1" Then
                            C.Text = .Fields(C.Tag) & "%"
                        ElseIf C.Name = "cmbArea" Then
                            C.Text = .Fields(C.Tag) & ""
                        Else
                            'C.Text = .Fields(C.Tag) & ""
                            For i = 0 To C.ListCount - 1
                                If C.List(i) = .Fields(C.Tag) & "" Then
                                    C.ListIndex = i
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                ElseIf (TypeOf C Is MSForms.OptionButton) Then
                    ''''
                ElseIf (TypeOf C Is VB.CheckBox) Then
                    C.Value = Abs(.Fields(C.Tag))
                Else
                    C.Text = .Fields(C.Tag) & ""
                End If
            End If
        Next
                
        .Close
       
    
    End With
    
    Set rs = Nothing
    
End Sub



Private Function CreateInsertSQL(strTable As String, Optional With_ As Boolean = False) As String
    
    Dim C As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    
    For Each C In Controls
        If C.Tag <> "" Then
            'If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Or _
            (TypeOf C Is MSForms.CheckBox) Or (TypeOf C Is ComboList.Usercontrol1) Or _
            (TypeOf C Is VB.TextBox) Or (TypeOf C Is VB.ComboBox) Then
            
            RetString = RetString & C.Tag & ","
            
            'End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now The Second Part...
    '---------------------------------------------------------------------------
    
    For Each C In Controls
        
        If C.Tag <> "" Then
            If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is VB.TextBox) Then
                RetString = RetString & "'" & C.Text & "',"
            ElseIf (TypeOf C Is MSForms.CheckBox) Or (TypeOf C Is VB.CheckBox) Then
                RetString = RetString & Abs(C.Value) & ","
            ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & C.ID & "',"
            ElseIf (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is VB.ComboBox) Then
                If Right(C.Name, 1) = "L" Then
                    RetString = RetString & C.ListIndex & ","
                Else
                    If C.Name = "cmbShare1" Then
                        RetString = RetString & Val(C.Text) & ","
                    Else
                        RetString = RetString & "'" & C.Text & "',"
                    End If
                End If
            ElseIf (TypeOf C Is MSForms.OptionButton) Then
                If C.Value = True Then
                    RetString = RetString & "'" & C.Caption & "',"
                End If
            ElseIf (TypeOf C Is DTPicker) Then
                If C.Value & "" = "" Then
                    RetString = RetString & "Null,"
                Else
                    RetString = RetString & "'" & Format(C.Value, "dd-MMM-yyyy") & "',"
                End If
            Else
                RetString = RetString & "'" & C & "',"
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

Private Function CreateUpdateSQL(strTable As String, Optional With_ As Boolean = False) As String
    
    Dim C As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each C In Controls
    
        If C.Tag <> "" And C.Name <> "lblEmpID" Then
            If (TypeOf C Is MSForms.TextBox) Then   '
                If C.Tag = "name" Then
                    RetString = RetString & "[Name]" & "='" & C.Text & "',"
                ElseIf LCase(C.Name) Like "nbox*" Then
                         RetString = RetString & C.Tag & "=" & C.Text & ","
                Else
                     RetString = RetString & C.Tag & "='" & C.Text & "',"
                End If
                
            ElseIf (TypeOf C Is MSForms.CheckBox) Or (TypeOf C Is VB.CheckBox) Then
                RetString = RetString & C.Tag & "=" & Abs(C.Value) & ","
            ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                RetString = RetString & C.Tag & "='" & C.ID & "',"
            ElseIf (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is VB.ComboBox) Then
                If Right(C.Name, 1) = "L" Then
                    RetString = RetString & C.Tag & "=" & C.ListIndex & ","
                Else
                    If C.Name = "cmbShare1" Then
                        RetString = RetString & C.Tag & "=" & Val(C) & ","
                    Else
                        RetString = RetString & C.Tag & "='" & C & "',"
                    End If
                End If
            ElseIf (TypeOf C Is DTPicker) Then
                If C.Value & "" = "" Then
                    RetString = RetString & C.Tag & "=Null,"
                Else
                    RetString = RetString & C.Tag & "='" & Format(C.Value, "dd-MMM-yyyy") & "',"
                End If
     
            Else
                RetString = RetString & C.Tag & "='" & C & "',"
            End If
        End If
    Next
    
    'Get For ListView
    'For i = 1 To LV2.ListItems.Count
    '    RetString = RetString & LV2.ListItems(i).Key & "=" & Abs(LV2.ListItems(i).Checked) & ","
    'Next i
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Private Sub SaveExpnReference(EmpId As String)
   
    con.Execute "DELETE FROM CVExperience WHERE CVID='" & EmpId & "'"
    
    For i = 1 To LVExp.ListItems.count
        With LVExp.ListItems(i)
            con.Execute "INSERT INTO CVExperience(CVID,SrNo,Designation,Institute,Experience) VALUES('" & _
             EmpId & "'," & Val(.Text) & ",'" & .SubItems(1) & "','" & .SubItems(2) & "'," & _
             Val(.SubItems(3)) & ")"
        End With
    Next
    
    
    con.Execute "DELETE FROM CVReference WHERE CVID='" & EmpId & "'"
    For i = 1 To LVRef.ListItems.count
        With LVRef.ListItems(i)
            con.Execute "INSERT INTO CVReference(CVID,SrNo,Name,Relation,Organization,ContactNo) VALUES('" & _
             EmpId & "'," & Val(.Text) & ",'" & .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & _
             .SubItems(4) & "')"
        End With
    Next
    
End Sub


Private Sub SaveQualifications(EmpId As String)

    con.Execute "DELETE FROM CVAcademicQualifications WHERE CVID='" & EmpId & "'"

    For i = 1 To LVQ.ListItems.count
        With LVQ.ListItems(i)
            con.Execute "INSERT INTO CVAcademicQualifications (CVID,SrNo,Diploma,Institute," & _
             "PassingYear,MarksObtained,TotalMarks)VALUES('" & EmpId & "'," & .Text & ",'" & .SubItems(1) & _
             "','" & .SubItems(3) & "'," & .SubItems(2) & "," & Val(.SubItems(4)) & "," & Val(.SubItems(5)) & ")"
        End With
    Next
    
    con.Execute "DELETE FROM CVComputerLiteracy WHERE CVID='" & EmpId & "'"

    For i = 1 To LVComputer.ListItems.count
        With LVComputer.ListItems(i)
            con.Execute "INSERT INTO CVComputerLiteracy(CVID,SrNo,Diploma,Institute," & _
             "Period,MarksObtained,TotalMarks)VALUES('" & EmpId & "'," & .Text & ",'" & .SubItems(1) & _
             "','" & .SubItems(3) & "','" & .SubItems(2) & "'," & Val(.SubItems(4)) & "," & Val(.SubItems(5)) & ")"
        End With
    Next
    
End Sub

Private Sub LoadQualifications(EnrollNo As String)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * From CVAcademicQualifications Where CVID='" & EmpId & "' Order By SrNo", con, adOpenForwardOnly, adLockReadOnly
        LVQ.ListItems.Clear
        Do Until .EOF
            Set ITM = LVQ.ListItems.add(, , ![SrNo] & "")
            ITM.SubItems(1) = !Diploma & ""
            ITM.SubItems(2) = !PassingYear & ""
            ITM.SubItems(3) = !Institute & ""
            
            ITM.SubItems(4) = Val(!MarksObtained & "")
            ITM.SubItems(5) = Val(!TotalMarks & "")
            ITM.SubItems(6) = Round(((Val(ITM.SubItems(4)) / Val(ITM.SubItems(5))) * 100), 2) & " %"
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select * From CVComputerLiteracy Where CVID='" & EmpId & "' Order By SrNo", con, adOpenForwardOnly, adLockReadOnly
        LVComputer.ListItems.Clear
        Do Until .EOF
            Set ITM = LVComputer.ListItems.add(, , ![SrNo] & "")
            ITM.SubItems(1) = !Diploma & ""
            ITM.SubItems(2) = !Period & ""
            ITM.SubItems(3) = !Institute & ""
            
            ITM.SubItems(4) = Val(!MarksObtained & "")
            ITM.SubItems(5) = Val(!TotalMarks & "")
            If Val(ITM.SubItems(5)) = 0 Then
                ITM.SubItems(6) = "0"
            Else
                ITM.SubItems(6) = Round(((Val(ITM.SubItems(4)) / Val(ITM.SubItems(5))) * 100), 2) & " %"
            End If
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Private Sub LoadExpnReference(EmpId As String)

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM CVExperience WHERE CVID='" & EmpId & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVExp.ListItems.Clear
        Do Until .EOF
           Set ITM = LVExp.ListItems.add(, , LVExp.ListItems.count + 1)
           ITM.ListSubItems.add , , !Designation & ""
           ITM.ListSubItems.add , , !Institute & ""
           ITM.ListSubItems.add , , Val(!Experience & "")
           .MoveNEXT
        Loop
        .Close
         
        .Open "SELECT * FROM CVReference WHERE CVID='" & EmpId & "' ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVRef.ListItems.Clear
        Do Until .EOF
            Set ITM = LVRef.ListItems.add(, , LVRef.ListItems.count + 1)
            ITM.ListSubItems.add , , !Name & ""
            ITM.ListSubItems.add , , !Relation & ""
            ITM.ListSubItems.add , , !Organization & ""
            ITM.ListSubItems.add , , !ContactNo & ""
           .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
End Sub

Public Function AddNewGuarantor()
    'lblEmpID = GetNextCVNo
    Me.Show
End Function

Private Function GetNextCVNo()

   
    
    Dim NextVal As String
    
    NextVal = GetMax("CAST(RIGHT(CVID,2) AS INT)", "CVs", "WHERE LEFT(CVID,3)='CV-'")
    GetNextCVNo = "CV-" & Format(Val(NextVal) + 1, "000")
   
End Function
