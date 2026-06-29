VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmNewLab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Lab..."
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   135
      TabIndex        =   5
      Top             =   45
      Width           =   5475
      Begin MSForms.TextBox TBox 
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   435
         Width           =   3090
         VariousPropertyBits=   679495707
         ScrollBars      =   2
         Size            =   "5450;503"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   810
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   1050
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
         Caption         =   "Address"
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
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   825
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lab Name"
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
         TabIndex        =   6
         Top             =   225
         Width           =   825
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5685
      TabIndex        =   4
      Top             =   1065
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewLab.frx":0000
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
      TabIndex        =   2
      Top             =   135
      Width           =   1665
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewLab.frx":0112
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
      Top             =   600
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewLab.frx":0224
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
Attribute VB_Name = "frmNewLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iID As Integer
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        Unload Me
        Me.Show 1
    End If

End Sub

Public Sub EditCat(TempID As Integer)
    
    iID = TempID
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From TestLabs Where ID=" & iID, con, adOpenForwardOnly, adLockReadOnly
        TBox(0) = ![Name] & ""
        TBox(1) = ![Address] & ""
        .Close
    End With
    Me.Show 1
    
    Set rs = Nothing
End Sub
Private Function Saved() As Boolean

On Error GoTo err

    If Trim(TBox(0)) = "" Then
        MsgBox "Invalid Name.", vbInformation
        SaveItem = False
        Exit Function
    End If


    If CatID = "" Then
        con.Execute "Insert into TestLabs(Name,Address) " & _
         "Values('" & TBox(0).Text & "','" & TBox(1) & "')"
    Else
        con.Execute "Update TestLabs Set Name='" & TBox(0) & "',Address='" & TBox(1) & "' Where ID=" & iID
    End If
    
    ID = ""
    Saved = True
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub
