VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form NewDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departments"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbAccruedAcc 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Tag             =   "PRTAccNo"
      Top             =   420
      Width           =   2715
      _ExtentX        =   4789
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
   Begin ComboList.Usercontrol1 cmbUnionAcc 
      Height          =   285
      Left            =   2880
      TabIndex        =   20
      Tag             =   "PRTAccNo"
      Top             =   990
      Width           =   2715
      _ExtentX        =   4789
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
   Begin ComboList.Usercontrol1 cmbEOBIAcc 
      Height          =   285
      Left            =   2880
      TabIndex        =   25
      Tag             =   "PchAccNo"
      Top             =   1575
      Width           =   2715
      _ExtentX        =   4789
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
   Begin ComboList.Usercontrol1 cmbOnePercentAcc 
      Height          =   285
      Left            =   2880
      TabIndex        =   29
      Tag             =   "PchAccNo"
      Top             =   2130
      Width           =   2715
      _ExtentX        =   4789
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
   Begin ComboList.Usercontrol1 cmbSalaryAcc 
      Height          =   285
      Left            =   135
      TabIndex        =   15
      Tag             =   "PchAccNo"
      Top             =   420
      Width           =   2715
      _ExtentX        =   4789
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
   Begin ComboList.Usercontrol1 cmbFoodAcc 
      Height          =   285
      Left            =   135
      TabIndex        =   19
      Tag             =   "PchAccNo"
      Top             =   990
      Width           =   2715
      _ExtentX        =   4789
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
   Begin ComboList.Usercontrol1 cmbSchAcc 
      Height          =   285
      Left            =   135
      TabIndex        =   23
      Tag             =   "PchAccNo"
      Top             =   1575
      Width           =   2715
      _ExtentX        =   4789
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
   Begin ComboList.Usercontrol1 cmbTaxAcc 
      Height          =   285
      Left            =   135
      TabIndex        =   27
      Tag             =   "PchAccNo"
      Top             =   2130
      Width           =   2715
      _ExtentX        =   4789
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
   Begin VB.CheckBox chkTemp 
      Caption         =   "Dept. of Temporary Employees"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   14
      Top             =   5205
      Visible         =   0   'False
      Width           =   2565
   End
   Begin MSComDlg.CommonDialog cdColor 
      Left            =   6345
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2700
      Left            =   135
      TabIndex        =   5
      Top             =   2460
      Width           =   5475
      Begin VB.PictureBox picColor 
         BackColor       =   &H00800000&
         Height          =   315
         Left            =   3690
         ScaleHeight     =   255
         ScaleWidth      =   1635
         TabIndex        =   12
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Double Click To Select)"
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
         Left            =   1860
         TabIndex        =   13
         Top             =   2340
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Color :"
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
         Left            =   210
         TabIndex        =   11
         Top             =   2340
         Width           =   1605
      End
      Begin MSForms.TextBox lblCatID 
         Height          =   300
         Left            =   105
         TabIndex        =   10
         Top             =   435
         Width           =   1695
         VariousPropertyBits=   679495707
         Size            =   "2990;529"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   1170
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   1050
         Width           =   5280
         VariousPropertyBits=   -1467987941
         MaxLength       =   255
         ScrollBars      =   2
         Size            =   "9313;2064"
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
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department ID"
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
         Left            =   135
         TabIndex        =   8
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Name"
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
         Left            =   1845
         TabIndex        =   7
         Top             =   210
         Width           =   1560
      End
      Begin MSForms.TextBox TBox 
         Height          =   300
         Index           =   0
         Left            =   1845
         TabIndex        =   0
         Top             =   450
         Width           =   3420
         VariousPropertyBits=   679495707
         Size            =   "6032;529"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCatID1 
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
         Height          =   300
         Left            =   150
         TabIndex        =   6
         Top             =   1305
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fine Account"
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
      Left            =   2880
      TabIndex        =   30
      Top             =   1905
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tax"
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
      TabIndex        =   28
      Top             =   1905
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Accrued EOBI Account"
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
      Left            =   2880
      TabIndex        =   26
      Top             =   1365
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bonus"
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
      Left            =   135
      TabIndex        =   24
      Top             =   1350
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Accrued Food Account"
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
      Left            =   135
      TabIndex        =   22
      Top             =   765
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Over Time Account"
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
      Left            =   2865
      TabIndex        =   21
      Top             =   780
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Accrued Account"
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
      Left            =   2865
      TabIndex        =   18
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salary Account"
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
      Left            =   135
      TabIndex        =   17
      Top             =   195
      Width           =   1275
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5685
      TabIndex        =   4
      Top             =   4770
      Width           =   1665
      Caption         =   "Cancel            "
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
   Begin MSForms.CommandButton cmdNext 
      Height          =   375
      Left            =   5685
      TabIndex        =   2
      Top             =   3840
      Width           =   1665
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "NewDepartment.frx":0000
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   5685
      TabIndex        =   3
      Top             =   4305
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;661"
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
Attribute VB_Name = "NewDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public add As Boolean, edit As Boolean
Dim Loaded As Boolean
Public TableName As String

Private Sub Chksubof_Click()
    If chkSubof = vbChecked Then
        cmbCatID.Enabled = True
        cmbCat.Enabled = True
        'Call AddToCombo(cmbCatID, "DeptID", "Departments", " Order By DeptID")
        'Call AddToCombo(cmbCat, "Name", "Departments", " Order By deptID")
    Else
        cmbCatID.Enabled = False
        cmbCat.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        Unload Me
        
        Load NewDepartment
        NewDepartment.TableName = TableName
        NewDepartment.lblCatID = GetNextDeptID(TableName)
        NewDepartment.add = True
        NewDepartment.edit = False
        NewDepartment.Show 1
    End If

End Sub

Public Sub Editrec()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
   
        .Open "SELECT * FROM " & TableName & " WHERE Deptid='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        lblCatID = ![DeptID] & ""
        TBox(0) = ![Name] & ""
        TBox(1) = ![Description] & ""
        PicColor.BackColor = Val(![Color] & "")
        chkTemp.Value = IIf(IsNull(!TempDept) = True, 0, Abs(!TempDept))
        If chkTemp.Value = vbUnchecked Then
            cmbSalaryAcc.ID = !SalaryAccNo & ""
            cmbAccruedAcc.ID = !AccruedAccNo & ""
            cmbFoodAcc.ID = !AccruedFoodAccNo & ""
            cmbUnionAcc.ID = !AccruedUnionAccNo & ""
            cmbSchAcc.ID = !SchAccNo & ""
            cmbEOBIAcc.ID = !AccruedEOBIAccNo & ""
            cmbTaxAcc.ID = !TaxAccNo & ""
            cmbOnePercentAcc.ID = !OnePercentAccNo & ""
        End If
        .Close
    End With

    Set rs = Nothing

    Loaded = True
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub cmbCatID_Validate(cancel As Boolean)
    If Not cmbCatID.MatchFound Then cmbCatID.ListIndex = cmbCat.ListIndex
End Sub

Private Sub cmbCat_Change()
    If cmbCat.MatchFound Then Call cmbCat_Click Else If cmbCat <> "" Then cmbCat.DropDown
End Sub

Private Sub cmbCat_Click()
    If cmbCatID.ListIndex <> cmbCat.ListIndex Then cmbCatID.ListIndex = cmbCat.ListIndex
End Sub

Private Sub cmbCat_Validate(cancel As Boolean)
   If Not cmbCat.MatchFound Then cmbCat.ListIndex = cmbCatID.ListIndex
End Sub

Private Function Saved() As Boolean

    On Error GoTo err

    If Trim(TBox(0)) = "" Then
        MsgBox "Invalid Department Name", vbInformation
        SaveItem = False
        Exit Function
    End If

    If add Then
        Call StartTrans(con)
        If chkTemp.Value = vbChecked Then
            con.Execute "INSERT INTO Departments (DeptID,Name,Description,Active,Color,TempDept) " & _
             "Values('" & lblCatID & "','" & TBox(0).Text & "','" & TBox(1) & "',1," & _
              PicColor.BackColor & "," & Abs(chkTemp.Value) & ")"
        Else
            con.Execute "Insert Into Departments (DeptID,Name,Description,Active,Color,TempDept,SalaryAccNo,AccruedAccNo,AccruedFoodAccNo,AccruedUnionAccNo,SchAccNo,AccruedEOBIAccNo,TaxAccNo,OnePercentAccNo) " & _
             "Values('" & lblCatID & "','" & TBox(0).Text & "','" & TBox(1) & "',1," & _
              PicColor.BackColor & "," & Abs(chkTemp.Value) & ",'" & cmbSalaryAcc.ID & "','" & cmbAccruedAcc.ID & "','" & cmbFoodAcc.ID & "','" & cmbUnionAcc.ID & "','" & cmbSchAcc.ID & "','" & cmbEOBIAcc.ID & "','" & cmbTaxAcc.ID & "','" & cmbOnePercentAcc.ID & "')"
        End If
        con.CommitTrans
    Else
        If chkTemp.Value = vbChecked Then
            con.Execute "Update Departments Set Name='" & TBox(0) & "',Description='" & TBox(1) & _
            "',Color=" & PicColor.BackColor & ",TempDept=" & Abs(chkTemp.Value) & " WHERE DeptID='" & Me.Tag & "'"
        Else
            con.Execute "UPDATE Departments SET Name='" & TBox(0) & "',Description='" & _
             TBox(1) & "',Color=" & PicColor.BackColor & ",TempDept=" & Abs(chkTemp.Value) & _
             ",SalaryAccNo='" & cmbSalaryAcc.ID & "',AccruedAccNo='" & cmbAccruedAcc.ID & _
             "',AccruedFoodAccNo='" & cmbFoodAcc.ID & "',AccruedUnionAccNo='" & _
             cmbUnionAcc.ID & "',SchAccNo='" & cmbSchAcc.ID & "',AccruedEOBIAccNo='" & _
             cmbEOBIAcc.ID & "',TaxAccNo='" & cmbTaxAcc.ID & "',OnePercentAccNo='" & cmbOnePercentAcc.ID & "' WHERE DeptID='" & Me.Tag & "'"
        End If
        con.Execute "UPDATE Departments SET Active=1 WHERE DeptID='" & Me.Tag & "'"
    End If

    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub cmdOK_Click()

    If Saved Then Unload Me

End Sub


Private Sub Form_Activate()

    If edit = False Then
        lblCatID = GetNextDeptID("Departments")
    End If

End Sub

Private Sub Form_Load()

    cmbSalaryAcc.ListHeight = 2400
    cmbAccruedAcc.ListHeight = 2400
    cmbFoodAcc.ListHeight = 2400
    cmbUnionAcc.ListHeight = 2400
    cmbSchAcc.ListHeight = 2400
    cmbEOBIAcc.ListHeight = 2400
    cmbTaxAcc.ListHeight = 2400
    
    Call cmbSalaryAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where  Parent=0 AND Active=1 Order By ACCTitle")
    Call cmbAccruedAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where Parent=0 AND Active=1 Order By ACCTitle")
    Call cmbFoodAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where  Parent=0 AND Active=1 Order By ACCTitle")
    Call cmbUnionAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where Parent=0 AND Active=1 Order By ACCTitle")
    Call cmbSchAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where Parent=0 AND Active=1 Order By ACCTitle")
    Call cmbEOBIAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where Parent=0 AND Active=1 Order By ACCTitle")
    Call cmbTaxAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where Parent=0 AND Active=1 Order By ACCTitle")
    Call cmbOnePercentAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where Parent=0 AND Active=1 Order By ACCTitle")
    
End Sub

Private Sub PicColor_DblClick()

    cdColor.Color = PicColor.BackColor
    
    cdColor.ShowColor
    
    If cdColor.Color <> 0 Then
        PicColor.BackColor = cdColor.Color
    End If

End Sub

