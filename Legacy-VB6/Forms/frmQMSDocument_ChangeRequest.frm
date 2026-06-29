VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQMSDocument_ChangeRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Control (Change Request)"
   ClientHeight    =   5115
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
   ScaleHeight     =   5115
   ScaleWidth      =   9165
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdViewPDF 
      Caption         =   "View PDF"
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
      Left            =   6705
      TabIndex        =   21
      Top             =   3690
      Width           =   1635
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8730
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
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
      TabIndex        =   9
      Top             =   3990
      Width           =   8040
   End
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
      TabIndex        =   3
      Top             =   3975
      Width           =   405
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   5565
      TabIndex        =   0
      Tag             =   "RevisionDT"
      Top             =   360
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   118292483
      CurrentDate     =   40698
   End
   Begin MSForms.CommandButton cmdApprove 
      Height          =   420
      Left            =   315
      TabIndex        =   20
      Top             =   4530
      Visible         =   0   'False
      Width           =   2370
      Caption         =   "Approve"
      PicturePosition =   327683
      Size            =   "4180;741"
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   3
      Left            =   1965
      TabIndex        =   19
      Top             =   360
      Width           =   1605
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2831;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbDocumentNo 
      Height          =   315
      Left            =   315
      TabIndex        =   18
      Tag             =   "QMS_RefID"
      Top             =   360
      Width           =   1635
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2884;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Required"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   17
      Top             =   1845
      Width           =   1245
   End
   Begin MSForms.TextBox TBox 
      Height          =   1620
      Index           =   1
      Left            =   315
      TabIndex        =   16
      Tag             =   "DetailofChange"
      Top             =   2055
      Width           =   8430
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14870;2857"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section Effected"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   15
      Top             =   870
      Width           =   1185
   End
   Begin MSForms.TextBox TBox 
      Height          =   705
      Index           =   0
      Left            =   315
      TabIndex        =   14
      Tag             =   "SectionEffected"
      Top             =   1080
      Width           =   8430
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14870;1244"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Type"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   1965
      TabIndex        =   13
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revision #"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   7230
      TabIndex        =   12
      Top             =   150
      Width           =   765
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   4
      Left            =   7230
      TabIndex        =   2
      Tag             =   "RevisionNo"
      Top             =   360
      Width           =   1530
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2699;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   2
      Left            =   3585
      TabIndex        =   1
      Top             =   360
      Width           =   1965
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "3466;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Name"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   2
      Left            =   3585
      TabIndex        =   11
      Top             =   150
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Updated PDF Path to Attach"
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
      Left            =   330
      TabIndex        =   10
      Top             =   3705
      Width           =   3015
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Document No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   73
      Left            =   315
      TabIndex        =   8
      Top             =   150
      Width           =   1020
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Revision Date"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   5550
      TabIndex        =   7
      Top             =   150
      Width           =   990
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   7215
      TabIndex        =   6
      Top             =   4530
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
      TabIndex        =   5
      Top             =   4530
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
      TabIndex        =   4
      Top             =   4530
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
Attribute VB_Name = "frmQMSDocument_ChangeRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim strFileName As String
Dim strDelUsers As String

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

    On Error GoTo 0
    
'    cmbEmp.ListHeight = 2500
'    Call cmbEmp.AddVals(con, "'{'+EmpID+'} '+Name", "Employees", "EmpID")
    
    'Call AddToCombo(cmbDocumentType, "DocumentType", "QMS_DocumentTypes")
    Call AddToCombo(cmbDocumentNo, "DocumentNumber", "QMS_Documents", , , "EntryID")
End Sub



Private Sub cmbDocumentNo_Change()
    Call cmbDocumentNo_Click
End Sub

Private Sub cmbDocumentNo_Click()
    
    If cmbDocumentNo.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT DocumentName,DocumentType FROM QMS_Documents WHERE EntryID=" & Val(cmbDocumentNo.List(cmbDocumentNo.ListIndex, 1)), con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            TBox(3) = !DocumentType & ""
            TBox(2) = !DocumentName & ""
        End If
        .Close
        
        .Open "SELECT * FROM QMS_Documents_ChangeRequests WHERE QMS_RefID=" & Val(cmbDocumentNo.List(cmbDocumentNo.ListIndex, 1)) & " AND Approved IS NULL", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then 'Already Change Request Pending....
            DTPicker2.Value = !RevisionDT
            TBox(4) = !RevisionNo & ""
            TBox(0) = !SectionEffected & ""
            TBox(1) = !DetailOfChange & ""
            lEntryID = Val(!EntryID & "")
            If IsNull(!ChangedPDF) = False Then
                txtFilePath.Text = "Attached"
                txtFilePath.Tag = "Attached"
            End If
            
            cmdApprove.Visible = UserHasAccess("CloseCAPA")
        Else
            lEntryID = 0
        End If
        .Close
   
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdApprove_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
    
        Dim lQMSEntryID As Long
        lQMSEntryID = Val(cmbDocumentNo.List(cmbDocumentNo.ListIndex, 1))
        Call StartTrans(con)
        con.Execute "UPDATE QMS_Documents_ChangeRequests SET Approved=1,ApprovedUserName='" & CurrentUserName & "',ApprovedMachineName='" & strComputerName & "',ApprovedEntryDT=getDate() WHERE EntryID=" & lEntryID
        con.Execute "UPDATE QMS_Documents SET Revision='" & TBox(4) & "' WHERE EntryID=" & lQMSEntryID
        
        'Update Original PDF Details in Change Request for History Purpose...
        con.Execute "UPDATE QMS_Documents_ChangeRequests SET OriginalPDF=T1.AttachmentPDF,OriginalPDFFileName=T1.AttachedPDFFileName FROM QMS_Documents_ChangeRequests " & _
         "INNER JOIN QMS_Documents T1 ON T1.EntryID=QMS_Documents_ChangeRequests.QMS_RefID WHERE QMS_Documents_ChangeRequests.EntryID=" & lEntryID
         
        
        'Update PDF Details in Document...
        con.Execute "UPDATE QMS_Documents SET AttachmentPDF=ChangedPDF,AttachedPDFFileName=ChangedPDFFileName FROM QMS_Documents " & _
         "INNER JOIN QMS_Documents_ChangeRequests ON QMS_Documents.EntryID=QMS_Documents_ChangeRequests.QMS_RefID WHERE QMS_Documents_ChangeRequests.EntryID=" & lEntryID
                  
        con.CommitTrans
        Unload Me
    End If
    
End Sub

Private Sub cmdBrowseBarcode_Click()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        txtFilePath = Picfile
        txtFilePath.Tag = ""
        strFileName = CD1.FileTitle
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


Private Sub cmdViewPDF_Click()

    If lEntryID = 0 Then Exit Sub
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "QMS_Documents_ChangeRequests", " WHERE EntryID=" & lEntryID & " AND ChangedPDF IS NOT NULL")
    
    If lCount = 0 Then
        MsgBox "No PDF attached.", vbInformation
        Exit Sub
    Else
        Call ShowTechDrawing("QMS_Documents_ChangeRequests", "ChangedPDF", " WHERE EntryID=" & lEntryID, txtFilePath.Text, CD1)
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker2.Value = Date
    strDelUsers = ""
    Call FillCmbs
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    
    
    Dim SQL As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        SQL = CreateInsertSQL("QMS_Documents_ChangeRequests")
    Else
        SQL = CreateUpdateSQL("QMS_Documents_ChangeRequests")
        SQL = SQL & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute SQL
     
    con.CommitTrans
     
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "QMS_Documents_ChangeRequests")
    Else
        lmyEntryID = lEntryID
    End If
    
    Call AttachPDF(lmyEntryID)
   
    
    lEntryID = 0
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
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Then
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
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.ComboBox) Then
                'RetString = RetString & c.Tag & "='" & c.Text & "',"
                RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
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
        .Open "SELECT * FROM QMS_Documents WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    c.ID = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is DTPicker Then
                    If Not IsNull(.Fields(c.Tag)) Then
                        c.Value = .Fields(c.Tag)
                    End If
                Else
                    c.Text = .Fields(c.Tag) & ""
                End If
            End If
        Next
        txtFilePath = !AttachedPDFFileName & ""
        txtFilePath.Tag = !AttachedPDFFileName & ""
        .Close
       
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub ChangeRequest(p_lQMSEntryID As Long)
    
    Dim i As Integer
    For i = 0 To cmbDocumentNo.ListCount - 1
        If p_lQMSEntryID = Val(cmbDocumentNo.List(i, 1)) Then
            cmbDocumentNo.ListIndex = i
            Exit For
        End If
    Next
    
    
    Me.Show 1
    
End Sub

Private Function GetNewDocumentNo() As String
    
    Dim lDocNo As Long
    lDocNo = GetSingleLongValue("MAX(CAST(RIGHT(DocumentNumber,LEN(DocumentNumber)-8) AS INT))", "QMS_Documents")
    lDocNo = lDocNo + 1
    
    GetNewDocumentNo = "QMS-DOC/" & lDocNo
    
End Function

Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" And txtFilePath.Tag = "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,ChangedPDF,ChangedPDFFileName FROM QMS_Documents_ChangeRequests WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            Rec.Fields(2) = strFileName
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub
