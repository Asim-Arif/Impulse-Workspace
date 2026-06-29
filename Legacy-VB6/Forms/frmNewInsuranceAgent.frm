VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmNewInsuranceAgent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insurance Agent"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
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
   ScaleHeight     =   5820
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   45
      TabIndex        =   15
      Top             =   0
      Width           =   5805
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Agent"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Index           =   0
         Left            =   1650
         TabIndex        =   16
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Agent"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   1
         Left            =   1665
         TabIndex        =   17
         Top             =   135
         Width           =   2535
      End
   End
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   45
      TabIndex        =   14
      Top             =   570
      Width           =   5805
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Email1"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   105
         MaxLength       =   100
         TabIndex        =   0
         Tag             =   "Name"
         Top             =   435
         Width           =   5625
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Address1"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   8
         Left            =   105
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "Address"
         Top             =   2730
         Width           =   5595
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Email1"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   105
         MaxLength       =   100
         TabIndex        =   7
         Tag             =   "EMail1"
         Top             =   2175
         Width           =   2805
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Email2"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   2940
         MaxLength       =   100
         TabIndex        =   8
         Tag             =   "EMail2"
         Top             =   2175
         Width           =   2790
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Fax1"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   2010
         MaxLength       =   100
         TabIndex        =   5
         Tag             =   "Fax2"
         Top             =   1650
         Width           =   1860
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Phone1"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   105
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "Phone1"
         Top             =   1020
         Width           =   1860
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Phone2"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2010
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "Phone2"
         Top             =   1020
         Width           =   1860
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Phone3"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3900
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "Phone3"
         Top             =   1020
         Width           =   1830
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Fax2"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3900
         MaxLength       =   100
         TabIndex        =   6
         Tag             =   "Fax3"
         Top             =   1650
         Width           =   1830
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Fax1"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   105
         MaxLength       =   100
         TabIndex        =   4
         Tag             =   "Fax1"
         Top             =   1650
         Width           =   1860
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   630
         Left            =   420
         TabIndex        =   18
         Top             =   4470
         Width           =   5280
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   75
            TabIndex        =   11
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmNewInsuranceAgent.frx":0000
            Accelerator     =   110
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   13
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmNewInsuranceAgent.frx":0112
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSaveClose 
            Height          =   360
            Left            =   1815
            TabIndex        =   12
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmNewInsuranceAgent.frx":0224
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
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   225
         Width           =   405
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   27
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax 3"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   8
         Left            =   3930
         TabIndex        =   26
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone3"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   6
         Left            =   4035
         TabIndex        =   25
         Top             =   810
         Width           =   540
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone2"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   2100
         TabIndex        =   24
         Top             =   810
         Width           =   540
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone1"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   810
         Width           =   540
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax 1"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   22
         Top             =   1425
         Width           =   405
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email 2"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   15
         Left            =   3495
         TabIndex        =   21
         Top             =   1965
         Width           =   495
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email 1"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   20
         Top             =   1965
         Width           =   495
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax 2"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   30
         Left            =   2010
         TabIndex        =   19
         Top             =   1440
         Width           =   405
      End
      Begin MSForms.CheckBox chkinactive 
         Height          =   315
         Left            =   135
         TabIndex        =   10
         Tag             =   "Active"
         Top             =   4140
         Width           =   1785
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3149;556"
         Value           =   "0"
         Caption         =   "Agent is inactive"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Visible         =   0   'False
      Begin VB.Menu mnunewbank 
         Caption         =   "< Add New >"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBanks 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmNewInsuranceAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AgentID As Long
Dim Acno As String

Private Sub chkInactive_Click()
   If chkinactive = vbChecked And Val(lblBalance) <> 0 Then
      MsgBox "This Customer Cannot Be Made Inactive," & vbNewLine & _
             "To Make A Customer Inactive Balance Must Be Zero", vbInformation
      chkinactive = vbUnchecked
   End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Customer Name", vbInformation
        Exit Sub
    End If

    If Saved Then
        Unload Me
        CustID = 0
        Me.Show 1
    End If
Exit Sub

err:
 MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub


Private Sub cmdSaveClose_Click()

    On Error GoTo err
    If Trim(txtfields(0)) = "" Then
        MsgBox "Invalid Customer Name", vbInformation
        Exit Sub
    End If

    If Saved Then Unload Me

Exit Sub
err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    Me.KeyPreview = True
    AgentID = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub
Private Sub ClearFields()
   For i = 0 To txtfields.Count - 1
      txtfields(i) = ""
   Next
   txtfields(1).SetFocus
End Sub



Private Function CreateInsertSQL(strTable As String) As String
    
    Dim C As Control
    Dim RetString As String
    
    
    'Create The Field Name Portion of Qry...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    For Each C In Controls
        
        If (TypeOf C Is VB.TextBox) Or (TypeOf C Is VB.ComboBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is MSForms.CheckBox) Or (TypeOf C Is ComboList.Usercontrol1) Then
            If C.Tag <> "" Then
                RetString = RetString & C.Tag & ","
            End If
        End If
      
    Next
    
 
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    ' Create The Values Portion of Qry....
    '---------------------------------------------------------------------------
    For Each C In Controls
        
        
        If C.Tag <> "" Then
            If (TypeOf C Is VB.TextBox) Or (TypeOf C Is VB.ComboBox) Or (TypeOf C Is MSForms.ComboBox) Then
            
                If C.Name = "NBox" Then
                    RetString = RetString & Val(C.Text & "") & ","
                Else
                    RetString = RetString & "'" & C.Text & "',"
                End If
            ElseIf (TypeOf C Is MSForms.CheckBox) Then
                RetString = RetString & Abs(Not C.value) & ","
            ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & C.Text & "',"
            End If
        End If
       
    Next
    
      
    'Finalaize the Qry
    '---------------------------------------------------------------------------
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ")"
    CreateInsertSQL = RetString
    
End Function
Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim C As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each C In Controls
    
        If C.Tag <> "" Then
            If (TypeOf C Is VB.TextBox) Or (TypeOf C Is VB.ComboBox) Or (TypeOf C Is MSForms.ComboBox) Then '
                If C.Name = "NBox" Then
                    RetString = RetString & C.Tag & "=" & C.Text & ","
                Else
                    RetString = RetString & C.Tag & "='" & C.Text & "',"
                End If
            ElseIf (TypeOf C Is MSForms.CheckBox) Then
                RetString = RetString & C.Tag & "=" & Abs(Not C.value) & ","
            ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                RetString = RetString & C.Tag & "='" & C.Text & "',"
            End If
        End If
    Next
    
  
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Public Sub EditAgent(ID As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    AgentID = ID
    
    With rs
        .Open "Select * From InsuranceAgents Where EntryID=" & AgentID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each C In Controls
            If C.Tag <> "" Then
                If TypeOf C Is MSForms.CheckBox Then
                    C.value = Not .Fields(C.Tag)
                ElseIf TypeOf C Is ComboList.Usercontrol1 Then
                    C.Text = .Fields(C.Tag) & ""
                Else
                    C.Text = .Fields(C.Tag) & ""
                End If
            End If
        Next
        
        .Close
       
    End With
    
    
   Set rs = Nothing
   Me.Show 1
   
   Exit Sub
err:
    MsgBox err.Description
End Sub
Private Function Saved() As Boolean
    
    On Error GoTo err
    
    Dim Sql As String
    
    
    If AgentID = 0 Then
    
        Sql = CreateInsertSQL("InsuranceAgents")
    Else
       Sql = CreateUpdateSQL("InsuranceAgents")
       Sql = Sql & " Where EntryID=" & AgentID
    End If

    With con
        Call StartTrans(con)
   
        .Execute Sql, a
        .CommitTrans
   
    End With
    Saved = True
    Exit Function
err:
    MsgBox err.Description
End Function
