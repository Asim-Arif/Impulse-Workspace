VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmNewItemCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Catagories"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3270
      Left            =   135
      TabIndex        =   11
      Top             =   45
      Width           =   5475
      Begin VB.TextBox txtMaxLotQty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4155
         TabIndex        =   1
         Top             =   180
         Width           =   1230
      End
      Begin VB.TextBox txtHSCode 
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
         Left            =   960
         TabIndex        =   0
         Top             =   180
         Width           =   1815
      End
      Begin MSForms.TextBox txtFDADescription 
         Height          =   315
         Left            =   2820
         TabIndex        =   7
         Tag             =   "FDAProductCode"
         Top             =   2580
         Width           =   2580
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4551;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "FDA Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   6
         Left            =   2820
         TabIndex        =   21
         Top             =   2385
         Width           =   1335
      End
      Begin MSForms.TextBox txtFDAListingNo 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Tag             =   "FDAListingNo"
         Top             =   2580
         Width           =   1335
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2355;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "FDA Listing No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   76
         Left            =   120
         TabIndex        =   20
         Top             =   2370
         Width           =   1230
      End
      Begin MSForms.TextBox txtFDACode 
         Height          =   315
         Left            =   1470
         TabIndex        =   6
         Tag             =   "FDAProductCode"
         Top             =   2580
         Width           =   1335
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2355;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "FDA Prod. Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   77
         Left            =   1470
         TabIndex        =   19
         Top             =   2385
         Width           =   1275
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Max. Lot Qty. :"
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
         Left            =   2895
         TabIndex        =   18
         Top             =   195
         Width           =   1200
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "HS Code :"
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
         Left            =   150
         TabIndex        =   17
         Top             =   195
         Width           =   780
      End
      Begin MSForms.ComboBox cmbComplete 
         Height          =   300
         Left            =   1995
         TabIndex        =   4
         Top             =   1905
         Width           =   3405
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "6006;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "SRO Complete No."
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
         Left            =   2025
         TabIndex        =   16
         Top             =   1695
         Width           =   1500
      End
      Begin MSForms.ComboBox cmbSRO 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   1905
         Width           =   1815
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3201;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "SRO No."
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
         Left            =   150
         TabIndex        =   15
         Top             =   1695
         Width           =   645
      End
      Begin MSForms.TextBox TBox 
         Height          =   900
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   750
         Width           =   5280
         VariousPropertyBits=   -1467987941
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "9313;1587"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
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
         Left            =   150
         TabIndex        =   13
         Top             =   525
         Width           =   960
      End
      Begin VB.Label label1 
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
         Left            =   165
         TabIndex        =   12
         Top             =   945
         Visible         =   0   'False
         Width           =   1065
      End
      Begin MSForms.TextBox TBox 
         Height          =   285
         Index           =   0
         Left            =   4230
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   1125
         VariousPropertyBits=   679495707
         MaxLength       =   2
         ScrollBars      =   2
         Size            =   "1984;503"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5685
      TabIndex        =   10
      Top             =   1065
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewItemCat.frx":0000
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
      TabIndex        =   8
      Top             =   135
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
      TabIndex        =   9
      Top             =   600
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewItemCat.frx":0112
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
Attribute VB_Name = "frmNewItemCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CatID As Integer

Private Sub cmbSRO_Change()
    Call cmbSRO_Click
End Sub

Private Sub cmbSRO_Click()
    If cmbSRO.MatchFound = False Then Exit Sub
    cmbComplete.ListIndex = cmbSRO.ListIndex
End Sub
Private Sub cmbcomplete_Change()
    Call cmbcomplete_Click
End Sub

Private Sub cmbcomplete_Click()
    If cmbComplete.MatchFound = False Then Exit Sub
    cmbSRO.ListIndex = cmbComplete.ListIndex
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        Unload Me
        Me.Show 1
    End If

End Sub

Public Sub EditCat(strCatID As Integer)
    
    CatID = strCatID
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VItemCatagories WHERE CatID=" & CatID, con, adOpenForwardOnly, adLockReadOnly
        'TBox(0) = ![CatID] & ""
        TBox(1) = ![Description] & ""
        cmbSRO = ![SRONo] & ""
        txtHSCode = !HSCode & ""
        txtMaxLotQty = Val(!MaxLotSize & "")
        txtFDAListingNo = !FDAListingNo & ""
        txtFDACode = !FDACode & ""
        txtFDADescription = !FDADescription & ""
        .Close
    End With
    Me.Show 1
    
    Set rs = Nothing
End Sub
Private Function Saved() As Boolean

On Error GoTo err
    
    If Trim(txtHSCode) = "" Then
        MsgBox "Please Provide HS Code.", vbInformation
        txtHSCode.SetFocus
        Exit Function
    ElseIf Trim(TBox(1)) = "" Then
        MsgBox "Invalid Category.", vbInformation
        SaveItem = False
        Exit Function
    ElseIf cmbSRO.MatchFound = False Then
        MsgBox "Invalid SRO No. Selected.", vbInformation
        SaveItem = False
        Exit Function
    ElseIf (strCompany = "Tecno" Or strCompany = "Dr-Frgz") And Val(txtMaxLotQty) <= 0 Then
        MsgBox "Invalid Max. Lot Qty.", vbInformation
        SaveItem = False
        Exit Function
    End If

    
    If CatID = 0 Then
        con.Execute "INSERT INTO ItemCatagories(Description,SRONo,HSCode,MaxLotSize,FDAListingNo,FDACode,FDADescription) " & _
         "VALUES('" & TBox(1) & "'," & cmbSRO.List(cmbSRO.ListIndex, 1) & ",'" & txtHSCode & "'," & _
         Val(txtMaxLotQty) & ",'" & txtFDAListingNo & "','" & txtFDACode & "','" & txtFDADescription & "')"
    Else
        con.Execute "UPDATE ItemCatagories SET Description='" & TBox(1) & "',SRONo=" & _
         cmbSRO.List(cmbSRO.ListIndex, 1) & ",HSCode='" & txtHSCode & "',MaxLotSize=" & _
         Val(txtMaxLotQty) & ",FDAListingNo='" & txtFDAListingNo & "',FDACode='" & txtFDACode & "',FDADescription='" & txtFDADescription & "' WHERE CatID=" & CatID
    End If
    
    CatID = 0
    Saved = True
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select EntryID,SRONo,CRebate,CompleteNo,SRODT From SROs", con, adOpenForwardOnly, adLockReadOnly
        cmbSRO.Clear
        cmbComplete.Clear
        Do Until .EOF
            cmbSRO.AddItem ![SRONo] & ""
            cmbSRO.List(.AbsolutePosition - 1, 1) = ![EntryID] & ""
            cmbComplete.AddItem ![CompleteNo] & " ( " & Val(!CRebate & "") & " )" & " (" & Format(!SRODT, "dd-MMM-yyyy") & ")"
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    CatID = 0
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

