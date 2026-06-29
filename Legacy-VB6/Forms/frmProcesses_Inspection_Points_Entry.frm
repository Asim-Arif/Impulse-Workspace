VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProcesses_Inspection_Points_Entry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Inspection Points"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
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
   ScaleHeight     =   7110
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      ForeColor       =   &H00800000&
      Height          =   6255
      Left            =   105
      TabIndex        =   3
      Top             =   195
      Width           =   6915
      Begin VB.TextBox txtLotNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   375
         Width           =   1785
      End
      Begin VB.TextBox txtProcessName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1950
         TabIndex        =   6
         Top             =   375
         Width           =   4815
      End
      Begin VB.TextBox txtSortNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   2775
         TabIndex        =   5
         Top             =   3540
         Visible         =   0   'False
         Width           =   1000
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5370
         Left            =   150
         TabIndex        =   0
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9472
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inspection Point"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Result"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Lot No."
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   150
         Width           =   525
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Process"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   9
         Left            =   1950
         TabIndex        =   4
         Top             =   150
         Width           =   555
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5910
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   6585
      Width           =   1695
      Caption         =   " Save            "
      PicturePosition =   327683
      Size            =   "2990;661"
      Picture         =   "frmProcesses_Inspection_Points_Entry.frx":0000
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5340
      TabIndex        =   2
      Top             =   6585
      Width           =   1695
      Caption         =   "Close"
      PicturePosition =   327683
      Size            =   "2990;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmProcesses_Inspection_Points_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lVRD_EntryID  As Long

Public Sub RefreshList(iProcessID As Integer)

    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim iResult As Integer, strResult As String
    With rs
        .Open "SELECT * FROM Process_Inspection_Points WHERE ProcessID=" & iProcessID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !Point_Description & "")
            ITM.Tag = Val(!EntryID & "")
            iResult = GetSingleLongValue("InspectionData", "VRD_PIP_Details", " WHERE VRD_RefID=" & lVRD_EntryID & " AND P_I_P_RefID=" & Val(!EntryID & ""))
            If iResult = 1 Then
                strResult = "No"
            ElseIf iResult = 2 Then
                strResult = "Yes"
            ElseIf iResult = 3 Then
                strResult = "N.A"
            Else
                iResult = 2
                strResult = "Yes"
'                iResult = -1
'                strResult = ""
            End If
            ITM.ListSubItems.add(, , strResult).Tag = iResult
            .MoveNEXT
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo err
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).ListSubItems(1).Tag = "-1" Then
            MsgBox "Some Data Entry is missing.", vbInformation
            Exit Sub
        End If
    Next
    Call StartTrans(con)
    con.Execute "DELETE FROM VRD_PIP_Details WHERE VRD_RefID=" & lVRD_EntryID
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO VRD_PIP_Details(VRD_RefID,P_I_P_RefID,InspectionData) VALUES(" & lVRD_EntryID & "," & _
             Val(.Tag) & "," & .ListSubItems(1).Tag & ")"
        End With
    Next
    con.CommitTrans
    MsgBox "Done.", vbInformation
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim iResult As Integer, strResult As String
    iResult = LV.SelectedItem.ListSubItems(1).Tag
    
    If iResult = -1 Then
        iResult = 1
    ElseIf iResult = 1 Then
        iResult = 2
    ElseIf iResult = 2 Then
        iResult = 3
    ElseIf iResult = 3 Then
        iResult = 1
    End If
    
    If iResult = 1 Then
        strResult = "No"
    ElseIf iResult = 2 Then
        strResult = "Yes"
    ElseIf iResult = 3 Then
        strResult = "N.A"
    Else
        iResult = -1
        strResult = ""
    End If
    LV.SelectedItem.SubItems(1) = strResult
    LV.SelectedItem.ListSubItems(1).Tag = iResult
                
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
'    If KeyCode = 46 Then
'        If MsgBox("Do you want to Delete?", vbQuestion + vbYesNo) = vbYes Then
'            Dim iRowAffect As Integer
'            con.Execute "DELETE FROM Process_Inspection_Points WHERE EntryID=" & Val(LV.SelectedItem.Tag), iRowAffect
'            If iRowAffect = 1 Then
'                Call RefreshList
'            End If
'        End If
'    End If
    
End Sub


Public Sub ShowMe(p_lVRD_EntryID As Long)

    lVRD_EntryID = p_lVRD_EntryID
    Dim iProcessID As Integer, strLotNo As String
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ProcessID,LotNo FROM VendRcvdDetail WHERE EntryID=" & lVRD_EntryID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            iProcessID = Val(!ProcessID & "")
            strLotNo = !LotNo & ""
        End If
        .Close
    End With
    txtLotNo = strLotNo
    txtProcessName = GetSingleStringValue("Description", "Processes", " WHERE ProcessID=" & iProcessID)
    Call RefreshList(iProcessID)
    
    Me.Show 1
    
End Sub
