VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmMatCatagories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Catagories"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2955
      Left            =   135
      TabIndex        =   6
      Top             =   165
      Width           =   5475
      Begin ComboList.Usercontrol1 cmbCat 
         Height          =   285
         Left            =   1860
         TabIndex        =   14
         Top             =   1395
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Enabled         =   0   'False
      End
      Begin ComboList.Usercontrol1 cmbCatID 
         Height          =   285
         Left            =   135
         TabIndex        =   15
         Top             =   1395
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Enabled         =   0   'False
      End
      Begin VB.Frame fsubcat 
         Height          =   870
         Left            =   60
         TabIndex        =   11
         Top             =   915
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
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   225
            Width           =   1245
         End
      End
      Begin MSForms.TextBox TBox 
         Height          =   810
         Index           =   1
         Left            =   105
         TabIndex        =   2
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   450
         Width           =   1680
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5685
      TabIndex        =   5
      Top             =   1185
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
      Left            =   5670
      TabIndex        =   3
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
      TabIndex        =   4
      Top             =   720
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
Attribute VB_Name = "frmMatCatagories"
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
        'Call AddToCombo(cmbCatID, "CatID", "ItemCatagories", " Order By CatID")
        'Call AddToCombo(cmbCat, "Catagory", "ItemCatagories", " Order By CatID")
        Call cmbCat.ClearVals
        Call cmbCat.AddVals(con, "Catagory", "MaterialCatagories", "CatID", " Order By CatID")
        
        Call cmbCatID.ClearVals
        Call cmbCatID.AddVals(con, "CatID", "MaterialCatagories", "Catagory", " Order By CatID")
        
    Else
        cmbCatID.Enabled = False
        cmbCat.Enabled = False
    End If

End Sub

Private Sub cmbCatID_matched()
    cmbCat.Text = cmbCatID.ID
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
    .Open "Select * From VMaterialcatagories Where CatID='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
    lblItemID = ![CatID] & ""
    
    If IsNull(![SubCatof]) Or ![SubCatof] = "" Then chkSubof = vbUnchecked Else chkSubof = vbChecked
    
    lblCatID = ![CatID]
    TBox(0) = ![Catagory] & ""
    TBox(1) = ![Description] & ""
    cmbCatID.Text = ![SubCatof] & ""
    cmbCatID.SetFocus
    cmbCat.SetFocus
    TBox(0).SetFocus
    .Close
End With

Set rs = Nothing


Loaded = True
End Sub


Private Sub cmbCatID_Validate(Cancel As Boolean)
    If Not cmbCatID.MatchFound Then cmbCatID.Text = cmbCat.ID
    
End Sub
Private Sub cmbCat_Validate(Cancel As Boolean)
   If Not cmbCat.MatchFound Then cmbCat.Text = cmbCatID.ID
End Sub

Private Function GetNextCatID() As String
    Dim NextVal As String
    
    
    
    NextVal = GetMax("Cast(Right(CatID,2) as Int)", "MaterialCatagories", "Where SubCatOf ='" & IIf(chkSubof = vbChecked, cmbCatID.Text, "") & "'")
    GetNextCatID = IIf(chkSubof = vbChecked, cmbCatID.Text & "-", "") & Format(Val(NextVal) + 1, "00")
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
        con.Execute "Insert into MaterialCatagories(CatID,Catagory,SubCatOf,Description) " & _
         "Values('" & lblCatID & "','" & TBox(0).Text & "','" & cmbCatID.Text & "','" & TBox(1) & "')"
    Else
        con.Execute "Update MaterialCatagories Set Catagory='" & TBox(0) & "',Description='" & TBox(1) & "' Where CatID='" & Me.Tag & "'"
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
If Edit And Not Loaded Then Call showdata
End Sub

Private Sub Form_Load()
    cmbCat.ListHeight = 1200
    cmbCatID.ListHeight = 1200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub Tbox_Change(Index As Integer)
   If Index = 0 Then If Add Then lblCatID = GetNextCatID()
End Sub
