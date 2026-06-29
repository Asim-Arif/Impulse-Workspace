VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAuditAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audit"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9165
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9165
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBrowseBarcode 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8385
      TabIndex        =   22
      Top             =   7335
      Width           =   405
   End
   Begin VB.TextBox txtFilePath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      TabIndex        =   21
      Top             =   7350
      Width           =   8040
   End
   Begin VB.OptionButton OptAudit 
      Caption         =   "External"
      Height          =   270
      Index           =   1
      Left            =   5895
      TabIndex        =   16
      Top             =   375
      Width           =   1020
   End
   Begin VB.OptionButton OptAudit 
      Caption         =   "Internal"
      Height          =   270
      Index           =   0
      Left            =   4695
      TabIndex        =   15
      Tag             =   "InternalAudit"
      Top             =   375
      Width           =   1020
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   2865
      TabIndex        =   1
      Tag             =   "DT"
      Top             =   360
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   108068867
      CurrentDate     =   40698
   End
   Begin MSComCtl2.DTPicker DTNextAudit 
      Height          =   330
      Left            =   7125
      TabIndex        =   17
      Tag             =   "DTNextAudit"
      Top             =   360
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   108068867
      CurrentDate     =   40698
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PDF Path to Attach"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   315
      TabIndex        =   23
      Top             =   7065
      Width           =   2040
   End
   Begin MSForms.TextBox TBox 
      Height          =   1065
      Index           =   4
      Left            =   315
      TabIndex        =   20
      Tag             =   "ActionRequested"
      Top             =   5835
      Width           =   8475
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14949;1879"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actions / Recommendations"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   19
      Top             =   5595
      Width           =   1980
   End
   Begin VB.Label lblNextAudit 
      AutoSize        =   -1  'True
      Caption         =   "Next. Audit Date"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   7140
      TabIndex        =   18
      Top             =   150
      Width           =   1215
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   2
      Left            =   315
      TabIndex        =   2
      Tag             =   "ComplaintBy"
      Top             =   1035
      Width           =   8475
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "14949;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   345
      TabIndex        =   14
      Top             =   810
      Width           =   540
   End
   Begin MSForms.TextBox TBox 
      Height          =   1065
      Index           =   1
      Left            =   315
      TabIndex        =   3
      Tag             =   "Description"
      Top             =   1665
      Width           =   8475
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14949;1879"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Team"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   345
      TabIndex        =   13
      Top             =   1425
      Width           =   390
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Findings"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   345
      TabIndex        =   12
      Top             =   4125
      Width           =   585
   End
   Begin MSForms.TextBox TBox 
      Height          =   1065
      Index           =   0
      Left            =   315
      TabIndex        =   5
      Tag             =   "Findings"
      Top             =   4365
      Width           =   8475
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14949;1879"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scope"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   19
      Left            =   345
      TabIndex        =   11
      Top             =   2745
      Width           =   435
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Audit No. :"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   73
      Left            =   315
      TabIndex        =   10
      Top             =   150
      Width           =   780
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   14
      Left            =   315
      TabIndex        =   0
      Tag             =   "ComplaintNo"
      Top             =   360
      Width           =   2520
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   1065
      Index           =   15
      Left            =   315
      TabIndex        =   4
      Tag             =   "Reason"
      Top             =   2985
      Width           =   8475
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14949;1879"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   2895
      TabIndex        =   9
      Top             =   150
      Width           =   345
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   7215
      TabIndex        =   8
      Top             =   7980
      Width           =   1575
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2778;741"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   420
      Left            =   5580
      TabIndex        =   7
      Top             =   7980
      Width           =   1575
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2778;741"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   420
      Left            =   3945
      TabIndex        =   6
      Top             =   7980
      Visible         =   0   'False
      Width           =   1575
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2778;741"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmAuditAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long

Public Sub EditItem(p_lEntryID As Long)

    On Error GoTo err
    lEntryID = p_lEntryID
    
    Call ShowData(lEntryID)
    
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillCmbs()

    
    
    
End Sub


Private Sub cmdBrowseBarcode_Click()
    
    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        txtFilePath = Picfile
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdNext_Click()
    If SaveItem Then
        Unload Me
        
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub

Private Sub DTPicker2_Change()
    If lEntryID = 0 Then
        Tbox(14) = GetNewAuditNo(DTPicker2.Value)
    End If
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    
    
    Call FillCmbs
    OptAudit(0).Value = True
     
    DTPicker2.Value = Date
    Call DTPicker2_Change
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    
    
    Dim Sql As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        Sql = CreateInsertSQL("CustomerComplaints")
    Else
        Sql = CreateUpdateSQL("CustomerComplaints")
        Sql = Sql & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute Sql
    
    
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "CustomerComplaints")
    Else
        lmyEntryID = lEntryID
    End If
    
    con.Execute "UPDATE CustomerComplaints SET Type=2 WHERE EntryID=" & lmyEntryID
     
    con.CommitTrans
       
    Call AttachPDF(lmyEntryID)
    SaveItem = True
    Exit Function
err:
    MsgBox err.Description
    
End Function



Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "INSERT INTO " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Or (TypeOf c Is VB.OptionButton) Then
            If c.Tag <> "" Then
                RetString = RetString & c.Tag & ","
            End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.ComboBox) Then
                If c.Name = "cmbCountry" Then
                    RetString = RetString & "'" & c.Text & "" & "',"
                ElseIf c.Name = "cmbPenalty" Then
                    RetString = RetString & c.ListIndex & ","
                Else
                    RetString = RetString & "'" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Then
                If c.Name = "NBox" Then
                    RetString = RetString & Val(c.Text & "") & ","
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & c.ID & "',"
            ElseIf (TypeOf c Is DTPicker) Then
                RetString = RetString & "'" & c.Value & "',"
            ElseIf (TypeOf c Is VB.OptionButton) Then
                RetString = RetString & "" & Abs(c.Value) & ","
            End If
        End If
    Next
    
    '---------------------------------------------------------------------------
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ")"
    CreateInsertSQL = RetString
    
End Function

Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "UPDATE " & strTable & " SET "
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.ComboBox) Then
                If c.Name = "cmbCountry" Then
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                ElseIf c.Name = "cmbPenalty" Then
                    RetString = RetString & c.Tag & "=" & c.ListIndex & ","
                Else
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Then   '
                If c.Name = "cmbCatID" Or c.Name = "cmbItemType" Then
                    RetString = RetString & c.Tag & "=" & Val(c.List(c.ListIndex, 1)) & ","
                ElseIf c.Name = "NBox" Then
                    RetString = RetString & c.Tag & "=" & Val(c.Text) & ","
                ElseIf c.Name = "cmbSteel" Or c.Name = "cmbItemGroup" Then
                    RetString = RetString & c.Tag & "=" & c.List(c.ListIndex, 1) & ","
                ElseIf c.Name = "cmbGuage" Then
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.ID & "',"
            ElseIf (TypeOf c Is DTPicker) Then
                RetString = RetString & c.Tag & "='" & c.Value & "',"
            ElseIf (TypeOf c Is VB.OptionButton) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function

Private Sub ShowData(p_lEntryID As Long, Optional bCopy As Boolean = False)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
        .Open "SELECT * FROM CustomerComplaints WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" And c.Name <> "OptAudit" Then
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    c.ID = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is DTPicker Then
                    If Not IsNull(.Fields(c.Tag)) Then
                        c.Value = .Fields(c.Tag)
                    End If
                ElseIf c.Name = "cmbPenalty" Then
                    c.ListIndex = Val(.Fields(c.Tag) & "")
                Else
                    c.Text = .Fields(c.Tag) & ""
                End If
            End If
        Next
        OptAudit(Abs(Not !InternalAudit)).Value = True
'        If Not IsNull(![AssetPic]) Then
'            Set Image1.DataSource = rs
'            Image1.DataField = "AssetPic"
'            lblPic.Visible = False
'        Else
'            lblPic.Visible = True
'        End If
        
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub AddNew()
    
    Tbox(14) = GetNewAuditNo(DTPicker2.Value)
    Me.Show 1
    
End Sub

Private Function GetNewAuditNo(myDT As Date) As String
    'CC-1/2011
    Dim lComplaintNo As Long
    lComplaintNo = GetSingleLongValue("MAX(CAST(SUBSTRING(ComplaintNo,4,CHARINDEX('/',ComplaintNo)-4) AS INT))", "CustomerComplaints", " WHERE Type=2")
    lComplaintNo = lComplaintNo + 1
    
    GetNewAuditNo = "AU-" & lComplaintNo & "/" & year(myDT)
    
End Function


Private Sub OptAudit_Click(Index As Integer)
    
    If Index = 0 Then
        lblNextAudit.Caption = "Target Date"
        lblReason.Caption = "Reason for Audit"
    Else
        lblNextAudit.Caption = "Next Date"
        lblReason.Caption = "Auditor"
    End If
    
End Sub

Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,AttachmentPDF FROM CustomerComplaints WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub

