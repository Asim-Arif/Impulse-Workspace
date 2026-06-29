VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form NewDepartmentint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Catagories"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2955
      Left            =   135
      TabIndex        =   8
      Top             =   165
      Width           =   5475
      Begin VB.Frame fsubcat 
         Height          =   870
         Left            =   60
         TabIndex        =   13
         Top             =   930
         Width           =   5355
         Begin VB.CheckBox chkSubof 
            Caption         =   "Sub Catagory Of"
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
            Left            =   105
            TabIndex        =   1
            Top             =   0
            Width           =   1695
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   300
            Left            =   90
            TabIndex        =   2
            Top             =   450
            Width           =   1680
            VariousPropertyBits=   679495705
            DisplayStyle    =   3
            Size            =   "2963;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Catagory ID "
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
            Left            =   105
            TabIndex        =   15
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Catagory Title "
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
            Left            =   1830
            TabIndex        =   14
            Top             =   225
            Width           =   1245
         End
         Begin MSForms.ComboBox cmbCat 
            Height          =   300
            Left            =   1830
            TabIndex        =   3
            Top             =   450
            Width           =   3420
            VariousPropertyBits=   679495705
            DisplayStyle    =   3
            Size            =   "6032;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSForms.TextBox TBox 
         Height          =   810
         Index           =   1
         Left            =   105
         TabIndex        =   4
         Top             =   2070
         Width           =   5280
         VariousPropertyBits=   -1467987941
         MaxLength       =   255
         ScrollBars      =   2
         Size            =   "9313;1429"
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
         TabIndex        =   12
         Top             =   1830
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Catagory ID "
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
         TabIndex        =   11
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Catagory Title "
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
         TabIndex        =   10
         Top             =   210
         Width           =   1245
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
      Begin VB.Label lblCatID 
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
         Left            =   120
         TabIndex        =   9
         Top             =   450
         Width           =   1680
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5685
      TabIndex        =   7
      Top             =   1185
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "NewDepartmentint.frx":0000
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
      Left            =   5670
      TabIndex        =   5
      Top             =   255
      Width           =   1665
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2937;661"
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
      TabIndex        =   6
      Top             =   720
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "NewDepartmentint.frx":0112
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
Attribute VB_Name = "NewDepartmentint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean
Dim Loaded As Boolean

Private Sub Chksubof_Click()
If chkSubof = vbChecked Then
    cmbCatID.Enabled = True
    cmbCat.Enabled = True
    Call AddToCombo(cmbCatID, "CatID", "ItemCatagories", " Order By CatID")
    Call AddToCombo(cmbCat, "Catagory", "ItemCatagories", " Order By CatID")
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
    Me.Add = True
    Me.Edit = False
    Me.Show 1
End If

End Sub

Private Sub showdata()

Dim rs As New ADODB.Recordset
With rs
   fsubcat.Enabled = False
    .Open "Select * From Vitemcatagories Where CatID='" & Me.Tag & "'", Con, adOpenForwardOnly, adLockReadOnly
    lblItemID = ![CatID] & ""
    
    If IsNull(![SubCatof]) Or ![SubCatof] = "" Then chkSubof = vbUnchecked Else chkSubof = vbChecked
    
    lblCatID = ![CatID]
    TBox(0) = ![Catagory] & ""
    TBox(1) = ![Description] & ""
    cmbCatID = ![SubCatof] & ""
    .Close
End With

Set rs = Nothing


Loaded = True
End Sub


Private Sub cmbCatID_Change()
    If cmbCatID.MatchFound Then Call cmbCatID_Click Else If Not cmbCatID = "" Then cmbCatID.DropDown
End Sub
Private Sub cmbCatID_Click()
    If cmbCat.ListIndex <> cmbCatID.ListIndex Then cmbCat.ListIndex = cmbCatID.ListIndex
    If Add Then lblCatID = GetNextCatID()
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

Private Function GetNextCatID() As String
    Dim NextVal As String
    NextVal = GetMax("val(Right(CatID,2))", "ItemCatagories", "Where SubCatOf ='" & IIf(chkSubof = vbChecked, cmbCatID, "") & "'")
    GetNextCatID = cmbCatID & "-" & Format(Val(NextVal) + 1, "00")
End Function



Private Function Saved() As Boolean

On Error GoTo err

    If Not cmbCatID.MatchFound And chkSubof = vbChecked Then
        MsgBox "Invalid Catagory", vbInformation
        SaveItem = False
        Exit Function
    ElseIf Trim(TBox(0)) = "" Then
        MsgBox "Invalid Catagory Name", vbInformation
        SaveItem = False
        Exit Function
    End If


If Add Then
   Con.Execute "Insert into ItemCatagories(CatID,Catagory,SubCatOf,Description) " & _
        "Values('" & lblCatID & "','" & TBox(0).Text & "','" & cmbCatID & "','" & TBox(1) & "')"
Else
   Con.Execute "Update Itemcatagories Set Catagory='" & TBox(0) & "',Description='" & TBox(1) & "' Where CatID='" & Me.Tag & "'"
End If
 Saved = True
Exit Function
err:
    MsgBox err.Description
End Function

Private Sub cmdOk_Click()
If Saved Then Unload Me
End Sub

Private Sub Form_Activate()
If Edit And Not Loaded Then Call showdata
End Sub

