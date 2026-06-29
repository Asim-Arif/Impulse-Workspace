VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmQMSDocument_AE 
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   315
      TabIndex        =   4
      Top             =   1005
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10860
      TabIndex        =   23
      Top             =   990
      Width           =   465
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3480
      Left            =   315
      TabIndex        =   22
      Top             =   1320
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   6138
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee"
         Object.Width           =   11377
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Department"
         Object.Width           =   3651
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Designation"
         Object.Width           =   3651
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
      TabIndex        =   13
      Top             =   5100
      Width           =   10560
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
      Left            =   10920
      TabIndex        =   7
      Top             =   5085
      Width           =   405
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
      Format          =   122290179
      CurrentDate     =   40698
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   7
      Left            =   9555
      TabIndex        =   27
      Top             =   150
      Width           =   495
   End
   Begin MSForms.ComboBox cmbInternal 
      Height          =   300
      Left            =   9555
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   165
      Width           =   570
   End
   Begin MSForms.ComboBox cmbDocumentType 
      Height          =   300
      Left            =   2910
      TabIndex        =   21
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
   Begin VB.Label lblDocTypes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2940
      MouseIcon       =   "frmQMSDocument_AE.frx":0000
      TabIndex        =   20
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   6765
      TabIndex        =   19
      Top             =   795
      Width           =   855
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   0
      Left            =   6765
      TabIndex        =   5
      Top             =   990
      Width           =   2055
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "3625;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   8850
      TabIndex        =   18
      Top             =   795
      Width           =   840
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   6255
      TabIndex        =   17
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
      Height          =   315
      Index           =   3
      Left            =   8850
      TabIndex        =   6
      Top             =   990
      Width           =   1980
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "3492;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distributor"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   345
      TabIndex        =   16
      Top             =   795
      Width           =   750
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
      TabIndex        =   15
      Top             =   150
      Width           =   1170
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
      Left            =   330
      TabIndex        =   14
      Top             =   4815
      Width           =   2040
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Document No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   73
      Left            =   315
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   150
      Width           =   780
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   9750
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   5640
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
      TabIndex        =   8
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
Attribute VB_Name = "frmQMSDocument_AE"
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
    
    cmbEmp.ListHeight = 2500
    Call cmbEmp.AddVals(con, "'{'+EmpID+'} '+Name", "Employees", "EmpID")
    
    Call AddToCombo(cmbDocumentType, "DocumentType", "QMS_DocumentTypes")
    With cmbInternal
        .AddItem "External"
        .AddItem "Internal"
    End With
    
End Sub

Private Sub cmbEmp_matched()
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT DeptName,Designation FROM VEmp WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            TBox(0) = ""
            TBox(3) = ""
        Else
            TBox(0) = !DeptName & ""
            TBox(3) = !Designation & ""
        End If
        .Close
    End With
    Set rs = Nothing
End Sub

Private Sub cmdAdd_Click()
    
    If cmbEmp.MatchFound = False Then
        MsgBox "Please Select Employee.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Tag = cmbEmp.ID Then
            MsgBox "Already Added.", vbInformation
            Exit Sub
        End If
    Next
        
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, , cmbEmp.Text)
    ITM.Tag = cmbEmp.ID
    ITM.ListSubItems.add , , TBox(0)
    ITM.ListSubItems.add , , TBox(3)
    
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


Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker2.Value = Date
    strDelUsers = ""
    Call FillCmbs
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    
    
    Dim Sql As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        Sql = CreateInsertSQL("QMS_Documents")
    Else
        Sql = CreateUpdateSQL("QMS_Documents")
        Sql = Sql & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute Sql
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "QMS_Documents")
    Else
        lmyEntryID = lEntryID
    End If
    
    'Distribution List..
    If strDelUsers <> "" Then
        strDelUsers = Left(strDelUsers, Len(strDelUsers) - 1)
        con.Execute "DELETE FROM QMS_Documents_DistributionList WHERE EntryID IN(" & strDelUsers & ")"
    End If
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.ListSubItems(1).Tag) = 0 Then
                con.Execute "INSERT INTO QMS_Documents_DistributionList(QMS_RefID,EmpID) VALUES(" & lmyEntryID & ",'" & .Tag & "')"
            End If
        End With
    Next
    con.CommitTrans
     
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "QMS_Documents")
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
            If c.Tag <> "" And c.Name <> "txtFilePath" Then
                RetString = RetString & c.Tag & ","
            End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        
        If c.Tag <> "" And c.Name <> "txtFilePath" Then
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
        
        If c.Tag <> "" And c.Name <> "txtFilePath" Then
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
        txtFilePath = !AttachedPDFFileName & ""
        txtFilePath.Tag = !AttachedPDFFileName & ""
        .Close
        
        Dim ITM As ListItem
        .Open "SELECT * FROM VQMS_Documents_DistributionList WHERE QMS_RefID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !EmpId & " " & !Name)
            ITM.Tag = !EmpId & ""
            ITM.ListSubItems.add(, , !DeptName & "").Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !Designation & ""
            
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

Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" And txtFilePath.Tag = "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,AttachmentPDF,AttachedPDFFileName FROM QMS_Documents WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
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

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LV.SelectedItem.ListSubItems(1).Tag) > 0 Then
                strDelUsers = strDelUsers & Val(LV.SelectedItem.ListSubItems(1).Tag) & ","
            End If
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub OptInternal_Click(Index As Integer)
    
End Sub
