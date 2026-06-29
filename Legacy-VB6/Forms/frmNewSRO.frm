VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewSRO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SRO"
   ClientHeight    =   4485
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
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
      TabIndex        =   10
      Top             =   0
      Width           =   5805
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SRO"
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
         Left            =   2580
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SRO"
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
         Left            =   2595
         TabIndex        =   12
         Top             =   135
         Width           =   735
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
      Height          =   3855
      Left            =   45
      TabIndex        =   9
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
         Tag             =   "SRONo"
         Top             =   435
         Width           =   1515
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Address1"
         DataSource      =   "datPrimaryRS"
         Height          =   1335
         Index           =   5
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "Description"
         Top             =   1650
         Width           =   5595
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Phone1"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "CompleteNo"
         Top             =   435
         Width           =   4080
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Phone2"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "STax"
         Top             =   1020
         Width           =   1350
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Phone3"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   4380
         MaxLength       =   20
         TabIndex        =   4
         Tag             =   "EDuty"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Fax1"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   1620
         MaxLength       =   100
         TabIndex        =   2
         Tag             =   "CRebate"
         Top             =   1020
         Width           =   1350
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   630
         Left            =   420
         TabIndex        =   13
         Top             =   3090
         Width           =   5280
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   75
            TabIndex        =   6
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmNewSRO.frx":0000
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
            TabIndex        =   8
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmNewSRO.frx":0112
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
            TabIndex        =   7
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmNewSRO.frx":0224
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
      Begin MSComCtl2.DTPicker SRODT 
         Height          =   285
         Left            =   105
         TabIndex        =   20
         Tag             =   "SRODT"
         Top             =   1020
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   51052547
         CurrentDate     =   37250
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   810
         Width           =   345
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SRO No."
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   225
         Width           =   615
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripton"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   18
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excise Duty"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   6
         Left            =   4365
         TabIndex        =   17
         Top             =   810
         Width           =   840
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Tax"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   3000
         TabIndex        =   16
         Top             =   810
         Width           =   690
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complete SRO No."
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   1665
         TabIndex        =   15
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custom Rebate"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   1650
         TabIndex        =   14
         Top             =   810
         Width           =   1110
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
Attribute VB_Name = "frmNewSRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SROID As Long
Dim Acno As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    If Trim(txtfields(0)) = "" Then
        MsgBox "Invalid SRO No.", vbInformation
        Exit Sub
    ElseIf Val(txtfields(2)) = 0 Then
        MsgBox "Invalid Custom Rebate Value.", vbInformation
        Exit Sub
    'ElseIf Val(txtfields(3)) = 0 Then
    '    MsgBox "Invalid Sales Tax Value.", vbInformation
    '    Exit Sub
    'ElseIf Val(txtfields(4)) = 0 Then
    '    MsgBox "Invalid Excise Duty Value.", vbInformation
    '    Exit Sub
    End If
    
    If Saved Then
        Unload Me
        SROID = 0
        Me.Show 1
    End If
    Exit Sub

err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub


Private Sub cmdSaveClose_Click()

    On Error GoTo err
    If Trim(txtfields(0)) = "" Then
        MsgBox "Invalid SRO No.", vbInformation
        Exit Sub
    ElseIf Val(txtfields(2)) = 0 Then
        MsgBox "Invalid Custom Rebate Value.", vbInformation
        Exit Sub
    'ElseIf Val(txtfields(3)) = 0 Then
    '    MsgBox "Invalid Sales Tax Value.", vbInformation
    '    Exit Sub
    'ElseIf Val(txtfields(4)) = 0 Then
    '    MsgBox "Invalid Excise Duty Value.", vbInformation
    '    Exit Sub
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
        
        If (TypeOf C Is VB.TextBox) Or (TypeOf C Is VB.ComboBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is MSForms.CheckBox) Or (TypeOf C Is ComboList.Usercontrol1) Or (TypeOf C Is DTPicker) Then
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
            ElseIf (TypeOf C Is DTPicker) Then
                RetString = RetString & "'" & C.value & "',"
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
            ElseIf (TypeOf C Is DTPicker) Then
                RetString = RetString & C.Tag & "='" & C.value & "',"
            End If
        End If
    Next
    
  
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Public Sub EditAgent(ID As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    SROID = ID
    
    With rs
        .Open "Select * From SROs Where EntryID=" & SROID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each C In Controls
            If C.Tag <> "" Then
                If TypeOf C Is MSForms.CheckBox Then
                    C.value = Not .Fields(C.Tag)
                ElseIf TypeOf C Is ComboList.Usercontrol1 Then
                    C.Text = .Fields(C.Tag) & ""
                ElseIf TypeOf C Is DTPicker Then
                    C.value = .Fields(C.Tag) & ""
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
    
    Dim SQL As String
    
    
    If SROID = 0 Then
    
        SQL = CreateInsertSQL("SROs")
    Else
       SQL = CreateUpdateSQL("SROs")
       SQL = SQL & " Where EntryID=" & SROID
    End If

    With con
        Call StartTrans(con)
   
        .Execute SQL, a
        .CommitTrans
   
    End With
    Saved = True
    Exit Function
err:
    MsgBox err.Description
End Function
