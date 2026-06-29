VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmQMSMeeting_AE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meeting"
   ClientHeight    =   7605
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
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBrowseBarcode 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6060
      TabIndex        =   25
      Top             =   7020
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
      TabIndex        =   24
      Top             =   7035
      Width           =   5940
   End
   Begin VB.Frame Frame1 
      Caption         =   "Meeting"
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
      Height          =   6510
      Left            =   90
      TabIndex        =   6
      Top             =   195
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
         Height          =   6090
         Left            =   165
         TabIndex        =   7
         Top             =   315
         Width           =   11130
         Begin ComboList.Usercontrol1 cmbAttendees 
            Height          =   285
            Left            =   6525
            TabIndex        =   23
            Top             =   465
            Width           =   4500
            _ExtentX        =   7938
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
         Begin VB.TextBox txtDiscussion 
            Appearance      =   0  'Flat
            Height          =   2715
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   20
            Tag             =   "Discussion"
            Top             =   3315
            Width           =   10950
         End
         Begin VB.TextBox txtAgenda 
            Appearance      =   0  'Flat
            Height          =   1305
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   19
            Tag             =   "Agenda"
            Top             =   1770
            Width           =   6315
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   105
            TabIndex        =   15
            Tag             =   "DT"
            Top             =   1080
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy HH:mm"
            Format          =   118292483
            CurrentDate     =   40774
         End
         Begin VB.ListBox lstAttendees 
            Appearance      =   0  'Flat
            Height          =   1590
            Left            =   6525
            TabIndex        =   12
            Top             =   825
            Width           =   4500
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   3255
            Top             =   -135
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "PDF Files Only|*.pdf"
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Discussion"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   22
            Top             =   3105
            Width           =   735
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Duration"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   4395
            TabIndex        =   21
            Top             =   840
            Width           =   615
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   4395
            TabIndex        =   18
            Tag             =   "Duration"
            Top             =   1050
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
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   2175
            MouseIcon       =   "frmQMSMeeting_AE.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   255
            Width           =   360
         End
         Begin MSForms.ComboBox cmbType 
            Height          =   315
            Left            =   2160
            TabIndex        =   16
            Tag             =   "Type"
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
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Meeting Date && time"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   14
            Top             =   870
            Width           =   1455
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Attendees"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   6525
            TabIndex        =   13
            Top             =   255
            Width           =   750
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Agenda / Purpose"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   11
            Top             =   1530
            Width           =   1290
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   1
            Tag             =   "Place"
            Top             =   1050
            Width           =   2205
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "3889;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Place"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   2160
            TabIndex        =   10
            Top             =   825
            Width           =   375
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   7380
            TabIndex        =   2
            Tag             =   "Keywords"
            Top             =   2460
            Width           =   3660
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "6456;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Keywords :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   6570
            TabIndex        =   9
            Top             =   2505
            Width           =   810
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Code"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   8
            Top             =   255
            Width           =   375
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   2
            Left            =   105
            TabIndex        =   0
            Tag             =   "Code"
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
      TabIndex        =   26
      Top             =   6750
      Width           =   2040
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   465
      Left            =   10005
      TabIndex        =   5
      Top             =   7005
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
      Left            =   8400
      TabIndex        =   4
      Top             =   7005
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
      Left            =   6795
      TabIndex        =   3
      Top             =   7005
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
Attribute VB_Name = "frmQMSMeeting_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim lEntryIDCopyFrom As Long
Dim strFileName As String

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
     
    cmbAttendees.ListHeight = 3000
    cmbAttendees.AddVals con, "'{' + EmpID + '} '+Name", "Employees", "EmpID", " WHERE Active=1"
     
    Call AddToCombo(cmbType, "MeetingType", "QMS_Meetings_Types")
    
End Sub


Private Sub cmbAttendees_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        lstAttendees.AddItem cmbAttendees.Text
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



Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker1.Value = Date
    
    Call FillCmbs
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    Dim i As Integer
     
    Dim SQL As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        SQL = CreateInsertSQL("QMS_Meetings")
    Else
        SQL = CreateUpdateSQL("QMS_Meetings")
        SQL = SQL & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute SQL
    
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "QMS_Meetings")
    Else
        lmyEntryID = lEntryID
    End If
    
    con.Execute "DELETE FROM QMS_Meetings_Attendees WHERE RefID=" & lmyEntryID
    
    For i = 0 To lstAttendees.ListCount - 1
        con.Execute "INSERT INTO QMS_Meetings_Attendees(RefID,AttendeeName) VALUES(" & lmyEntryID & ",'" & lstAttendees.List(i) & "')"
    Next
     
     
    Call AttachPDF(lmyEntryID)
    
    con.CommitTrans
    
'    If CurrentUserName = "Administrator" Then
'
'    End If
    
    SaveItem = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" And txtFilePath.Tag = "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,PDF,PDFFileName FROM QMS_Meetings WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
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

Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "INSERT INTO " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Or (TypeOf c Is VB.TextBox) Then
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
                If c.Name = "cmbType" Then
                    RetString = RetString & "'" & c.Text & "" & "',"
                Else
                    RetString = RetString & "'" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Then
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
                If c.Name = "cmbType" Then
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Then   '
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
    'Call getnextmeetingid
    
    Me.Show 1
    
End Sub

Private Sub ShowData(p_lEntryID As Long, Optional bCopy As Boolean = False)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
        .Open "SELECT * FROM QMS_Meetings WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
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
        txtFilePath = !PDFFileName & ""
        txtFilePath.Tag = !PDFFileName & ""
        
        .Close
        
        .Open "SELECT * FROM QMS_Meetings_Attendees WHERE RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        lstAttendees.Clear
        Do Until .EOF
            lstAttendees.AddItem !AttendeeName & ""
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
    
    Call GetNewMeetingID
    Me.Show 1
    
End Sub

Private Sub GetNewMeetingID()
    'Get Help ID...
    TBox(2) = GetSingleLongValue("MAX(CAST(Code AS INT))", "QMS_Meetings") + 1
    TBox(2) = Format(TBox(2), "0000")
End Sub

Private Sub lblPic_DblClick()

    
    
End Sub

Private Sub label1_Click(Index As Integer)
    If Index = 6 Then 'Type
        Dim f As New frmMiscS_New
        Load f
        f.ShowForm "QMS_Meetings_Types", "MeetingType", "Type", "EntryID", "Meeting Types"
         
        Call AddToCombo(cmbType, "MeetingType", "QMS_Meetings_Types")
    End If
End Sub

Private Sub lstAttendees_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstAttendees.ListCount = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to Remove?", vbQuestion + vbYesNo) = vbYes Then
            lstAttendees.RemoveItem lstAttendees.ListIndex
        End If
    End If
    
End Sub
