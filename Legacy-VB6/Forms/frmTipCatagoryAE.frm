VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTipCatagoryAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip Catagory"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   135
      TabIndex        =   4
      Top             =   45
      Width           =   5475
      Begin MSForms.TextBox TBox 
         Height          =   1770
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   5280
         VariousPropertyBits=   -1467987941
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "9313;3122"
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
         TabIndex        =   6
         Top             =   165
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
         Left            =   165
         TabIndex        =   5
         Top             =   945
         Visible         =   0   'False
         Width           =   1065
      End
      Begin MSForms.TextBox TBox 
         Height          =   285
         Index           =   0
         Left            =   4230
         TabIndex        =   7
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
      TabIndex        =   3
      Top             =   1065
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmTipCatagoryAE.frx":0000
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
      TabIndex        =   1
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
      TabIndex        =   2
      Top             =   600
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmTipCatagoryAE.frx":0112
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
Attribute VB_Name = "frmTipCatagoryAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CatID As Integer

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
        .Open "SELECT * FROM Tip_Master WHERE EntryID=" & CatID, con, adOpenForwardOnly, adLockReadOnly
        TBox(1) = ![Description] & ""
        .Close
    End With
    Me.Show 1
    
    Set rs = Nothing
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    
    If Trim(TBox(1)) = "" Then
        MsgBox "Invalid Category.", vbInformation
        SaveItem = False
        Exit Function
    End If

    
    If CatID = 0 Then
        con.Execute "INSERT INTO Tip_Master(Description) " & _
         "VALUES('" & TBox(1) & "')"
    Else
        con.Execute "UPDATE Tip_Master SET Description='" & TBox(1) & "' WHERE EntryID=" & CatID
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
    
    CatID = 0
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub
