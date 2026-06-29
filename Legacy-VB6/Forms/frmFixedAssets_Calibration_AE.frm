VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmFixedAssets_Calibration_AE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calibration"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Certificate :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2370
      Left            =   135
      TabIndex        =   13
      Top             =   3330
      Width           =   6360
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
         Left            =   5790
         TabIndex        =   21
         Top             =   1740
         Width           =   405
      End
      Begin VB.TextBox txtFilePath 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   135
         TabIndex        =   20
         Top             =   1755
         Width           =   5640
      End
      Begin VB.TextBox txtMasterReference 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2385
         TabIndex        =   17
         Top             =   480
         Width           =   3825
      End
      Begin VB.TextBox txtCertNo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         TabIndex        =   15
         Top             =   480
         Width           =   2235
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
         Left            =   150
         TabIndex        =   22
         Top             =   1500
         Width           =   2040
      End
      Begin MSForms.ComboBox cmbVender 
         Height          =   330
         Left            =   135
         TabIndex        =   19
         Top             =   1095
         Width           =   6075
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "10716;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vender"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   18
         Top             =   855
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Master Calibrator Refernce"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2385
         TabIndex        =   16
         Top             =   270
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Certificate No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdSaveNext 
      Caption         =   "Sa&ve && New"
      Height          =   495
      Left            =   2655
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin ComboList.Usercontrol1 cmbAsset 
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   6360
      _ExtentX        =   11218
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save && Close"
      Height          =   495
      Left            =   3975
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   1485
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   6360
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   1020
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
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
      Format          =   118685699
      CurrentDate     =   37055
   End
   Begin MSComCtl2.DTPicker DTSchedule 
      Height          =   315
      Left            =   2115
      TabIndex        =   2
      Top             =   1020
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
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
      Format          =   118685699
      CurrentDate     =   37055
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1425
      Top             =   5715
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
   End
   Begin MSForms.ComboBox cmbCalibrationFrequency 
      Height          =   315
      Left            =   4230
      TabIndex        =   12
      Top             =   1020
      Width           =   2265
      VariousPropertyBits=   746604575
      DisplayStyle    =   3
      Size            =   "3995;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Calibration Frequency"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4245
      TabIndex        =   11
      Top             =   795
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Next Due Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2130
      TabIndex        =   10
      Top             =   795
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   9
      Top             =   1575
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Calibration Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   8
      Top             =   795
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Asset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   7
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frmFixedAssets_Calibration_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lEntryID As Long
Dim bCancelled As Boolean
Dim strFileName As String

Private Sub cmbAsset_matched()
    Dim lFrequency As Long
    lFrequency = GetSingleLongValue("CalibrationFrequency", "FixedAssets", " WHERE AssetID='" & cmbAsset.ID & "'")
    cmbCalibrationFrequency.ListIndex = lFrequency
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

Private Sub cmdSave_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdSaveNext_Click()
    If Saved Then
        txtRemarks = ""
        cmbAsset.ID = ""
        cmbAsset.SetFocus
    End If
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
     
    If cmbAsset.MatchFound = False Then
        MsgBox "Please Select Asset.", vbInformation
        Exit Function
    End If
    
    If DTSchedule.Value < DT.Value Then
        MsgBox "Invalid Next Schedule Date.", vbInformation
        Exit Function
    End If
    Dim strAccNo As String
    
    If cmbVender.MatchFound Then
        strAccNo = cmbVender.List(cmbVender.ListIndex, 1)
    Else
        strAccNo = ""
    End If
     
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        con.Execute "INSERT INTO FixedAssets_Calibration(AssetID,DT,Remarks,NextDueDT,UserName" & _
         ",MachineName,CertificateNo,MasterReference,AccNo) VALUES('" & cmbAsset.ID & "','" & DT.Value & "','" & txtRemarks & "','" & _
         DTSchedule.Value & "','" & CurrentUserName & "','" & strComputerName & "','" & txtCertNo & "','" & txtMasterReference & _
         "','" & strAccNo & "')"
         
         lmyEntryID = GetSingleLongValue("MAX(EntryID)", "FixedAssets_Calibration")
    Else
        con.Execute "UPDATE FixedAssets_Calibration SET AssetID='" & cmbAsset.ID & "',DT='" & DT.Value & "',NextDueDT='" & _
         DTSchedule.Value & "',CertificateNo='" & txtCertNo & "',MasterReference='" & txtMasterReference & "',AccNo='" & _
         strAccNo & "',Remarks='" & txtRemarks & "' WHERE EntryID=" & lEntryID
         
        lmyEntryID = lEntryID
    End If
    
    Call AttachPDF(lmyEntryID)
    
    Saved = True
    bCancelled = False
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub Form_Load()

    DT.Value = Date
    DTSchedule.Value = DateAdd("d", 7, Date)
    
    cmbAsset.ClearVals
    cmbAsset.ListHeight = 4000
    cmbAsset.AddVals con, "'{' + AssetID + '} ' + Description", "FixedAssets", "AssetID", " WHERE RequiresCalibration=1 ORDER BY AssetID"
   
    bCancelled = True
        
    cmbCalibrationFrequency.AddItem "Monthly"
    cmbCalibrationFrequency.AddItem "Yearly"
    
    
    Call AddToCombo(cmbVender, "AccTitle", "VVenders", , , "AccNo")
        
    lEntryID = 0
    
End Sub

Public Function AddNew() As Boolean
    Me.Show 1
    AddNew = Not bCancelled
End Function

Public Function EditEntry(p_lEntryID As Long) As Boolean

    lEntryID = p_lEntryID
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM FixedAssets_Calibration WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        cmbAsset.ID = !AssetID & ""
        DT.Value = !DT
        DTSchedule = !NextDueDT
        txtRemarks = !Remarks & ""
        
        txtCertNo = !CertificateNo & ""
        txtMasterReference = !MasterReference & ""
        For i = 0 To cmbVender.ListCount - 1
            If cmbVender.List(i, 1) = !AccNo & "" Then
                cmbVender.ListIndex = i
                Exit For
            End If
        Next
        
        txtFilePath = !PDFFileName & ""
        txtFilePath.Tag = !PDFFileName & ""
        
        .Close
    End With
    Me.Show 1
    EditEntry = Not bCancelled
    
End Function


Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" And txtFilePath.Tag = "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,PDF,PDFFileName FROM FixedAssets_Calibration WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
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

