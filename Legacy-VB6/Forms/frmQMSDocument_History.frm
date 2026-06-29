VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmQMSDocument_History 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Control"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11805
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
   ScaleHeight     =   6135
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView LV 
      Height          =   4605
      Left            =   315
      TabIndex        =   13
      Top             =   990
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   8123
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detail of Change"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "App. Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Older PDF"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Changed PDF"
         Object.Width           =   2822
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1665
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   1545
      TabIndex        =   1
      Tag             =   "IssueDT"
      Top             =   360
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   130875395
      CurrentDate     =   40698
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Change History :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   18
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Type"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   7
      Left            =   9555
      TabIndex        =   17
      Top             =   150
      Width           =   1125
   End
   Begin MSForms.ComboBox cmbInternal 
      Height          =   300
      Left            =   9555
      TabIndex        =   16
      Tag             =   "Internal"
      Top             =   360
      Width           =   1770
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3122;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   300
      Index           =   1
      Left            =   7560
      TabIndex        =   15
      Tag             =   "InitiaterEmpID"
      Top             =   360
      Width           =   1980
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "3492;529"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initiator"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   6
      Left            =   7590
      TabIndex        =   14
      Top             =   165
      Width           =   570
   End
   Begin MSForms.ComboBox cmbDocumentType 
      Height          =   300
      Left            =   2910
      TabIndex        =   12
      Tag             =   "DocumentType"
      Top             =   360
      Width           =   1770
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3122;529"
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
      Caption         =   "Document Type"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   2940
      TabIndex        =   11
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   6255
      TabIndex        =   10
      Top             =   150
      Width           =   600
   End
   Begin MSForms.TextBox TBox 
      Height          =   300
      Index           =   4
      Left            =   6225
      TabIndex        =   3
      Tag             =   "Revision"
      Top             =   360
      Width           =   1320
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2328;529"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   300
      Index           =   2
      Left            =   4695
      TabIndex        =   2
      Tag             =   "DocumentName"
      Top             =   360
      Width           =   1515
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2672;529"
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
      Left            =   4710
      TabIndex        =   9
      Top             =   150
      Width           =   1170
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
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   14
      Left            =   315
      TabIndex        =   0
      Tag             =   "DocumentNumber"
      Top             =   360
      Width           =   1215
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2143;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Issue Date"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   1530
      TabIndex        =   7
      Top             =   150
      Width           =   780
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   9750
      TabIndex        =   6
      Top             =   5640
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
      Left            =   8100
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
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
      Left            =   6465
      TabIndex        =   4
      Top             =   5640
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
Attribute VB_Name = "frmQMSDocument_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim iColNo As Integer
Public Sub ShowHistory(p_lEntryID As Long)

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
    
    Call AddToCombo(cmbDocumentType, "DocumentType", "QMS_DocumentTypes")
    
    With cmbInternal
        .AddItem "External"
        .AddItem "Internal"
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker2.Value = Date
    iColNo = 0
    Call FillCmbs
    
End Sub


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
                If c.Name = "cmbDocumentType" Then
                    RetString = RetString & "'" & c.Text & "" & "',"
                ElseIf c.Name = "cmbInternal" Then
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
                If c.Name = "cmbInternal" Then
                    RetString = RetString & c.Tag & "=" & c.ListIndex & ","
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
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
                    If c.Name = "cmbInternal" Then
                        c.ListIndex = Abs(.Fields(c.Tag))
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        .Close
        
        Dim ITM As ListItem
        .Open "SELECT * FROM QMS_Documents_ChangeRequests WHERE QMS_RefID=" & p_lEntryID & " AND Approved=1", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            
            ITM.ListSubItems.add(, , Format(!DTEntry, "dd-MMM-yyyy")).Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !DetailOfChange & ""
            ITM.ListSubItems.add , , Format(!ApprovedEntryDT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !OriginalPDFFileName & ""
            ITM.ListSubItems.add , , !ChangedPDFFileName & ""
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub AddNew()
    
    TBox(14) = "" 'GetNewDocumentNo()
    Me.Show 1
    
End Sub

Private Function GetNewDocumentNo() As String
    
    Dim lDocNo As Long
    lDocNo = GetSingleLongValue("MAX(CAST(RIGHT(DocumentNumber,LEN(DocumentNumber)-8) AS INT))", "QMS_Documents")
    lDocNo = lDocNo + 1
    
    GetNewDocumentNo = "QMS-DOC/" & lDocNo
    
End Function

Private Sub LV_DblClick()
    
    
    Dim strFieldName As String
    
    If LV.ListItems.count = 0 Then Exit Sub
    If iColNo = 5 Then  'Older PDF
        strFieldName = "OriginalPDF"
    ElseIf iColNo = 6 Then  'Changed PDF
        strFieldName = "ChangedPDF"
    End If
     
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "QMS_Documents_ChangeRequests", " WHERE EntryID=" & Val(LV.SelectedItem.ListSubItems(1).Tag) & " AND " & strFieldName & " IS NOT NULL")
    
    If lCount = 0 Then
        MsgBox "No PDF attached.", vbInformation
        Exit Sub
    Else
        Call ShowTechDrawing("QMS_Documents_ChangeRequests", strFieldName, " WHERE EntryID=" & Val(LV.SelectedItem.ListSubItems(1).Tag), LV.SelectedItem.SubItems(iColNo - 1), CD1)
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        With LV.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit For
            End If
        End With
    Next
End Sub
