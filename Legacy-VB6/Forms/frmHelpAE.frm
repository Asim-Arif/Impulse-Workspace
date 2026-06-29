VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmHelpAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11760
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
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Help"
      ClipControls    =   0   'False
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
      Height          =   3945
      Left            =   90
      TabIndex        =   9
      Top             =   15
      Width           =   11475
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         Left            =   195
         TabIndex        =   10
         Top             =   420
         Width           =   11130
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   4965
            TabIndex        =   22
            Tag             =   "DTUpdate"
            Top             =   1770
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   117374979
            CurrentDate     =   40774
         End
         Begin VB.ListBox lstDesignations 
            Appearance      =   0  'Flat
            Height          =   2175
            Left            =   6525
            TabIndex        =   18
            Top             =   825
            Width           =   4500
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   3675
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "PDF Files Only|*.pdf"
         End
         Begin VB.CommandButton cmdBrowseBarcode 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   6060
            TabIndex        =   5
            Top             =   2550
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
            Left            =   105
            TabIndex        =   15
            Top             =   2565
            Width           =   5940
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Update Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   5010
            TabIndex        =   21
            Top             =   1545
            Width           =   915
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Designations"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   6525
            TabIndex        =   20
            Top             =   255
            Width           =   915
         End
         Begin MSForms.ComboBox cmbDesignation 
            Height          =   315
            Left            =   6525
            TabIndex        =   19
            Top             =   465
            Width           =   4485
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "7911;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Author"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   2190
            TabIndex        =   17
            Top             =   885
            Width           =   495
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   3
            Tag             =   "Author"
            Top             =   1095
            Width           =   4275
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "7541;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
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
            Left            =   120
            TabIndex        =   16
            Top             =   2280
            Width           =   2040
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   105
            TabIndex        =   2
            Tag             =   "Title"
            Top             =   1095
            Width           =   2040
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "3598;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Title"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   14
            Top             =   870
            Width           =   300
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Catagory"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   105
            TabIndex        =   13
            Top             =   240
            Width           =   675
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   315
            Left            =   105
            TabIndex        =   0
            Tag             =   "RefID"
            Top             =   465
            Width           =   4275
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "7541;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   105
            TabIndex        =   4
            Tag             =   "Description"
            Top             =   1770
            Width           =   4830
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "8520;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Keywords"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   12
            Top             =   1560
            Width           =   705
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Help ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   4395
            TabIndex        =   11
            Top             =   255
            Width           =   525
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   2
            Left            =   4395
            TabIndex        =   1
            Tag             =   "HelpID"
            Top             =   465
            Width           =   2040
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "3598;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   465
      Left            =   9960
      TabIndex        =   8
      Top             =   4065
      Width           =   1575
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2778;820"
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
      Height          =   465
      Left            =   8355
      TabIndex        =   7
      Top             =   4065
      Width           =   1575
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2778;820"
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
      Height          =   465
      Left            =   6750
      TabIndex        =   6
      Top             =   4065
      Width           =   1575
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2778;820"
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
Attribute VB_Name = "frmHelpAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim lEntryIDCopyFrom As Long

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
    
    Call AddToCombo(cmbDesignation, "Designation", "Designations")
    
    Call AddToCombo(cmbCatID, "Description", "Help_Master", "", False, "EntryID")
    
End Sub

Private Sub cmbDesignation_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode = 13 Then
        If cmbDesignation.MatchFound = False Then
            MsgBox "Invalid Designation."
            Exit Sub
        End If
        Dim i As Integer, iIndex As Integer
        For i = 0 To lstDesignations.ListCount - 1
            If cmbDesignation.Text = lstDesignations.List(i) Then
                MsgBox "Already Added.", vbInformation
                Exit Sub
            End If
        Next
        iIndex = lstDesignations.ListCount
        lstDesignations.AddItem cmbDesignation.Text
        'lstDesignations.ItemData(iIndex) = "0"
    End If
    
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



Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker1.Value = Date
    
    Call FillCmbs
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    Dim i As Integer
    If cmbCatID.MatchFound = False Then
        MsgBox "Invalid Catagory.", vbInformation
        Exit Function
    ElseIf TBox(0) = "" Then
        MsgBox "Invalid Topic.", vbInformation
        Exit Function
    ElseIf TBox(1) = "" Then
        MsgBox "Invalid Author.", vbInformation
        Exit Function
'    ElseIf TBox(3) = "" Then
'        MsgBox "Invalid Description.", vbInformation
'        Exit Function
    End If
    
    Dim Sql As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        Sql = CreateInsertSQL("Help")
    Else
        Sql = CreateUpdateSQL("Help")
        Sql = Sql & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute Sql
    
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "Help")
    Else
        lmyEntryID = lEntryID
    End If
    
    con.Execute "DELETE FROM Help_Designations WHERE Help_RefID=" & lmyEntryID
    
    
    For i = 0 To lstDesignations.ListCount - 1
        con.Execute "INSERT INTO Help_Designations(Help_RefID,Designation) VALUES(" & lmyEntryID & ",'" & lstDesignations.List(i) & "')"
    Next
    
    con.CommitTrans
    
    If CurrentUserName = "Administrator" Then
        Call AttachPDF(lmyEntryID)
    End If
    
    SaveItem = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,PDFAttachment FROM Help WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
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
                If c.Name = "cmbConvention" Then
                    RetString = RetString & "'" & c.Text & "" & "',"
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
                If c.Name = "cmbConvention" Then
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
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
            End If
        End If
    Next
    
    
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Public Sub CopyItem(p_lEntryID As Long)
       
    lEntryID = 0
    
    lEntryIDCopyFrom = p_lEntryID
    'CopyItemCode = strItemCode
    Call ShowData(lEntryIDCopyFrom, True)
    Call GetNewHelpID
    
    Me.Show 1
    
End Sub

Private Sub ShowData(p_lEntryID As Long, Optional bCopy As Boolean = False)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
        .Open "SELECT * FROM VHelp WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
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
                    If c.Name = "cmbCatID" Then
                        c.Text = .Fields("HelpCatagory")
                    Else
                        c.Text = .Fields(c.Tag) & ""
                    End If
                End If
            End If
        Next
        
        'TBox(8) = ![CatID] & "-"
        'TBox(3) = ![HelpID] & ""
        
        .Close
        
        .Open "SELECT * FROM Help_Designations WHERE Help_RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        lstDesignations.Clear
        Do Until .EOF
            lstDesignations.AddItem !Designation & ""
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
    
    Call GetNewHelpID
    Me.Show 1
    
End Sub

Private Sub GetNewHelpID()
    'Get Help ID...
    TBox(2) = GetSingleLongValue("MAX(CAST(HelpID AS INT))", "Help") + 1
    TBox(2) = Format(TBox(2), "0000")
End Sub

Private Sub lblPic_DblClick()

    
    
End Sub

Private Sub lstDesignations_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstDesignations.ListCount = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to Remove?", vbQuestion + vbYesNo) = vbYes Then
            lstDesignations.RemoveItem lstDesignations.ListIndex
        End If
    End If
    
End Sub
