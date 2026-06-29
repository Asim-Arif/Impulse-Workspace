VERSION 5.00
Begin VB.Form frmAEDesignation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Designation"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
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
   ScaleHeight     =   4155
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   465
      Left            =   3960
      TabIndex        =   9
      Top             =   3525
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   5655
      TabIndex        =   8
      Top             =   3525
      Width           =   1635
   End
   Begin VB.TextBox txtTrainings 
      Height          =   390
      Left            =   270
      TabIndex        =   7
      Top             =   2775
      Width           =   7020
   End
   Begin VB.TextBox txtSkills 
      Height          =   390
      Left            =   270
      TabIndex        =   5
      Top             =   1980
      Width           =   7020
   End
   Begin VB.TextBox txtDuties 
      Height          =   390
      Left            =   270
      TabIndex        =   3
      Top             =   1200
      Width           =   7020
   End
   Begin VB.TextBox txtDesignation 
      Height          =   390
      Left            =   270
      TabIndex        =   1
      Top             =   525
      Width           =   7020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Required Trainings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   6
      Top             =   2550
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Required Skills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   315
      TabIndex        =   4
      Top             =   1755
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Duties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   315
      TabIndex        =   2
      Top             =   975
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   300
      Width           =   1005
   End
End
Attribute VB_Name = "frmAEDesignation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lEntryID As Long
Dim bReturnValue As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    If Trim(txtDesignation.Text) = "" Then
        MsgBox "Please Enter Designation.", vbInformation
        Exit Sub
    End If
    
    If lEntryID = 0 Then
        con.Execute "INSERT INTO Designations(Designation,Duties,RequiredSkills,RequiredTrainings) VALUES('" & _
         Trim(txtDesignation) & "','" & txtDuties & "','" & txtSkills & "','" & txtTrainings & "')"
    Else
        con.Execute "UPDATE Designations SET Duties='" & txtDuties & "',RequiredSkills='" & txtSkills & _
         "',RequiredTrainings='" & txtTrainings & "' WHERE EntryID=" & lEntryID
    End If
    
    bReturnValue = True
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    lEntryID = 0
    bReturnValue = False
End Sub

Public Function ShowMe(Optional p_lEntryID As Long) As Boolean

    lEntryID = p_lEntryID
    If lEntryID > 0 Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT * FROM Designations WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            txtDesignation = !Designation & ""
            txtDuties = !Duties & ""
            txtSkills = !RequiredSkills & ""
            txtTrainings = !RequiredTrainings & ""
            .Close
        End With
    End If
    Me.Show 1
    ShowMe = bReturnValue
    
End Function
