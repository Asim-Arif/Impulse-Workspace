VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNC_AE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Non Compliance"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8880
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
   ScaleHeight     =   3720
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraReadOnly 
      Caption         =   "Customer Complaint Details :"
      Enabled         =   0   'False
      Height          =   1500
      Left            =   120
      TabIndex        =   3
      Top             =   75
      Width           =   8655
      Begin ComboList.Usercontrol1 cmbCustRO 
         Height          =   285
         Left            =   75
         TabIndex        =   4
         Tag             =   "CustCode"
         Top             =   450
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker DTPicker2RO 
         Height          =   330
         Left            =   6840
         TabIndex        =   5
         Tag             =   "DT"
         Top             =   450
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108068867
         CurrentDate     =   40698
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Complaint Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   5
         Left            =   6885
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   14
         Left            =   4290
         TabIndex        =   22
         Tag             =   "ComplaintNo"
         Top             =   450
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
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Complaint No."
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   73
         Left            =   4275
         TabIndex        =   21
         Top             =   240
         Width           =   1005
      End
      Begin MSForms.ComboBox cmbCountryRO 
         Height          =   315
         Left            =   2415
         TabIndex        =   20
         Tag             =   "Country"
         Top             =   450
         Width           =   1845
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3254;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   19
         Top             =   240
         Width           =   690
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Country"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   7
         Left            =   2430
         TabIndex        =   18
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complaint By"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   17
         Top             =   810
         Width           =   930
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   16
         Tag             =   "ComplaintBy"
         Top             =   1035
         Width           =   2310
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4075;556"
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
         Index           =   3
         Left            =   2430
         TabIndex        =   15
         Top             =   810
         Width           =   840
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   3
         Left            =   2415
         TabIndex        =   14
         Tag             =   "Designation"
         Top             =   1035
         Width           =   1845
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3254;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbPenaltyRO 
         Height          =   315
         Left            =   4290
         TabIndex        =   13
         Tag             =   "Penalty"
         Top             =   1035
         Width           =   1695
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2990;556"
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
         Caption         =   "Penalty"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   4275
         TabIndex        =   12
         Top             =   810
         Width           =   540
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   4
         Left            =   6015
         TabIndex        =   11
         Tag             =   "PenaltyQty"
         Top             =   1035
         Width           =   810
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "1429;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   5
         Left            =   6840
         TabIndex        =   10
         Tag             =   "PenaltyAmount"
         Top             =   1035
         Width           =   810
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "1429;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox TBoxRO 
         Height          =   315
         Index           =   6
         Left            =   7665
         TabIndex        =   9
         Tag             =   "PenaltyCurrency"
         Top             =   1035
         Width           =   810
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "1429;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   5
         Left            =   6030
         TabIndex        =   8
         Top             =   810
         Width           =   270
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   6
         Left            =   6915
         TabIndex        =   7
         Top             =   810
         Width           =   555
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   7
         Left            =   7695
         TabIndex        =   6
         Top             =   810
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2535
      Top             =   -195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
   End
   Begin MSComCtl2.DTPicker DTNC 
      Height          =   330
      Left            =   2505
      TabIndex        =   28
      Tag             =   "DT"
      Top             =   1815
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   108068867
      CurrentDate     =   40698
   End
   Begin MSForms.TextBox txtNC 
      Height          =   315
      Index           =   3
      Left            =   1965
      TabIndex        =   35
      Tag             =   "Description"
      Top             =   2445
      Width           =   6825
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "12039;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   1995
      TabIndex        =   34
      Top             =   2235
      Width           =   795
   End
   Begin MSForms.TextBox txtNC 
      Height          =   315
      Index           =   2
      Left            =   90
      TabIndex        =   33
      Tag             =   "Department"
      Top             =   2445
      Width           =   1845
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "3254;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   2235
      Width           =   855
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Source"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   6405
      TabIndex        =   31
      Top             =   1620
      Width           =   495
   End
   Begin MSForms.TextBox txtNC 
      Height          =   315
      Index           =   1
      Left            =   6435
      TabIndex        =   30
      Tag             =   "Source"
      Top             =   1815
      Width           =   2340
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "4128;556"
      Value           =   "Customer Complaint"
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
      Index           =   2
      Left            =   2535
      TabIndex        =   29
      Top             =   1605
      Width           =   345
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "NC No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   27
      Top             =   1605
      Width           =   510
   End
   Begin MSForms.TextBox txtNC 
      Height          =   315
      Index           =   0
      Left            =   135
      TabIndex        =   26
      Tag             =   "NCNo"
      Top             =   1815
      Width           =   2340
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "4128;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbNCType 
      Height          =   315
      Left            =   3990
      TabIndex        =   25
      Tag             =   "NCType"
      Top             =   1815
      Width           =   2430
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4286;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "NC Type"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   3990
      TabIndex        =   24
      Top             =   1605
      Width           =   615
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   7215
      TabIndex        =   2
      Top             =   3120
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
      TabIndex        =   1
      Top             =   3120
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
      TabIndex        =   0
      Top             =   3120
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
Attribute VB_Name = "frmNC_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim lCCEntryID As Long

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
    
    cmbCustRO.ListHeight = 2500
    Call cmbCustRO.AddVals(con, "DISTINCT CustCode", "ForeignCustomers", "CustCode")
    
    
    
    
    With cmbPenaltyRO
        .AddItem "None"
        .AddItem "Free Placement"
        .AddItem "Credit Note"
    End With
    
    With cmbNCType
        .AddItem "Major"
        .AddItem "Minor"
        .AddItem "Observation"
        .AddItem "Critical"
    End With
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    
    Sql = "SELECT DISTINCT Country FROM ForeignCustomers "
    
    If cmbCustRO.ID <> "0" Then
        Sql = Sql & " WHERE CustCode='" & cmbCustRO.ID & "'"
    End If
    
    With rs
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        cmbCountryRO.Clear
        Do Until .EOF
            cmbCountryRO.AddItem ![Country] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountryRO.ListCount > 0 Then cmbCountryRO.ListIndex = 0
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbPenalty_Change()
    Call cmbPenalty_Click
End Sub

Private Sub cmbPenalty_Click()

    TBoxRO(4).Enabled = False
    TBoxRO(5).Enabled = False
    TBoxRO(6).Enabled = False
    If cmbPenaltyRO.ListIndex = 1 Then
        TBoxRO(4).Enabled = True
    ElseIf cmbPenaltyRO.ListIndex = 2 Then
        TBoxRO(5).Enabled = True
        TBoxRO(6).Enabled = True
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

Private Sub DTNC_Change()
    txtNC(0) = GetNewNCNo(DTNC.Value)
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker2RO.Value = Date
    
    DTNC.Value = Date
    Call DTNC_Change
    Call FillCmbs
    
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    
    
    Dim Sql As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        Sql = CreateInsertSQL("NonCompliance")
    Else
        Sql = CreateUpdateSQL("NonCompliance")
        Sql = Sql & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute Sql
    
    
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "CustomerComplaints")
        con.Execute "UPDATE CustomerComplaints SET NCNo='" & txtNC(0).Text & "' WHERE EntryID=" & lCCEntryID
    Else
        lmyEntryID = lEntryID
    End If
    
    
    con.CommitTrans
   
   
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
        If Right(c.Name, 2) <> "RO" Then
            If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Then
                If c.Tag <> "" Then
                    RetString = RetString & c.Tag & ","
                End If
            End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each c In Controls
        If Right(c.Name, 2) <> "RO" Then
            If c.Tag <> "" Then
                If (TypeOf c Is DTPicker) Then
                    RetString = RetString & "'" & c.Value & "',"
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
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
        If Right(c.Name, 2) <> "RO" Then
            If c.Tag <> "" Then
                If (TypeOf c Is DTPicker) Then
                    RetString = RetString & c.Tag & "='" & c.Value & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            End If
        End If
    Next
    
    
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function



Private Sub ShowCCData(p_lEntryID As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    Dim iType As Integer, bInternalAudit As Boolean
    With rs
        .Open "SELECT * FROM CustomerComplaints WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If Right(c.Name, 2) = "RO" Then
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
                    ElseIf c.Name = "cmbPenaltyRO" Then
                        c.ListIndex = Val(.Fields(c.Tag) & "")
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        iType = Val(!Type & "")
        If iType = 2 Then
            bInternalAudit = !InternalAudit
        End If
        .Close
        
    End With
    If iType = 0 Then
        txtNC(1) = "Customer Complaint"
    ElseIf iType = 1 Then
        txtNC(1) = "Feedback/Observation"
    ElseIf iType = 2 Then
        If bInternalAudit Then
            txtNC(1) = "Internal Audit"
        Else
            txtNC(1) = "External Audit"
        End If
    End If
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub AddNew(p_ComplaintNo As Long)

    lCCEntryID = p_ComplaintNo
    Call ShowCCData(lCCEntryID)
    
    txtNC(0) = GetNewNCNo(DTNC.Value)
    txtNC(3) = GetSingleStringValue("Description", "CustomerComplaints", " WHERE EntryID=" & lCCEntryID)
    
    Me.Show 1
    
End Sub

Private Function GetNewNCNo(myDT As Date) As String
    'CC-1/2011
    Dim lNCNo As Long
    lNCNo = GetSingleLongValue("MAX(CAST(SUBSTRING(NCNo,4,CHARINDEX('/',NCNo)-4) AS INT))", "NonCompliance")
    lNCNo = lNCNo + 1
    
    GetNewNCNo = "NC-" & lNCNo & "/" & year(myDT)
    
End Function

Private Sub ShowData(p_lEntryID As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
        .Open "SELECT * FROM NonCompliance WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If Right(c.Name, 2) <> "RO" Then
                If c.Tag <> "" Then
                    If TypeOf c Is DTPicker Then
                        If Not IsNull(.Fields(c.Tag)) Then
                            c.Value = .Fields(c.Tag)
                        End If
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

