VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMonthLedger 
   ClientHeight    =   6585
   ClientLeft      =   165
   ClientTop       =   60
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11865
      Begin VB.CheckBox chkGraph 
         Caption         =   "&View Graph"
         Height          =   255
         Left            =   9360
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   7755
         TabIndex        =   4
         Top             =   345
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   93585411
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6210
         TabIndex        =   3
         Top             =   345
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   93585411
         CurrentDate     =   37055
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   10605
         TabIndex        =   10
         Top             =   300
         Width           =   1215
         Caption         =   "Refresh"
         PicturePosition =   327683
         Size            =   "2143;635"
         Picture         =   "frmMonthLedger.frx":0000
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMonthLedger.frx":015E
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   0
         Left            =   15
         TabIndex        =   11
         Top             =   120
         Width           =   11805
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   30
         TabIndex        =   0
         Top             =   345
         Width           =   1890
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   1905
         TabIndex        =   1
         Top             =   345
         Width           =   4305
         VariousPropertyBits=   545277979
         DisplayStyle    =   3
         Size            =   "7594;556"
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   1920
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   4320
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   705
      Top             =   2220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Trial Balance"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1695
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonthLedger.frx":01F9
            Key             =   "ACC"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonthLedger.frx":064D
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonthLedger.frx":0761
            Key             =   "PAR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonthLedger.frx":0875
            Key             =   "ACCIN"
         EndProperty
      EndProperty
   End
   Begin CRVIEWERLibCtl.CRViewer CRV 
      Height          =   5280
      Left            =   15
      TabIndex        =   6
      Top             =   735
      Width           =   10425
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   7920
      TabIndex        =   7
      Top             =   6120
      Width           =   1425
      ForeColor       =   64
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "frmMonthLedger.frx":0CC9
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   9420
      TabIndex        =   9
      Top             =   6120
      Width           =   1425
      Caption         =   "Close       "
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "frmMonthLedger.frx":0DDB
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
Attribute VB_Name = "frmMonthLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rpt As CRAXDDRT.Report
Dim dontShowList As Boolean
Dim PermissionDenied As Boolean

Private Sub CmbAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDown Then
   If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
   KeyCode = 0
   dontShowList = True
ElseIf KeyCode = vbKeyUp Then
   If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
   KeyCode = 0
   dontShowList = True
ElseIf KeyCode = 13 And List1.Visible Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If

End Sub

Private Sub CmbAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If dontShowList Then dontShowList = False Else ShowList List1, CmbAccTitle, CmbAccTitle
End Sub
Private Sub chkGraph_Click()
On Error GoTo err
   rpt.Section8.Suppress = (chkGraph <> vbChecked)
   CRV.Refresh
   Exit Sub
err:
If err.Number = -2147417848 Then Resume Next Else MsgBox "Error : " & err.Number & vbNewLine & err.Description
End Sub

Private Sub cmbAccNo_Change()
    If cmbAccNo.MatchFound Then Call cmbAccNo_Click
End Sub

Private Sub cmbAccNo_Click()

    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then
      CmbAccTitle.ListIndex = cmbAccNo.ListIndex
    End If
     
End Sub

Private Sub CmbAccTitle_Change()
    If CmbAccTitle.MatchFound Then
      Call cmbAccTitle_Click
    End If
      'List1.Visible = True
      
End Sub

Private Sub cmbAccTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
      
End Sub



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo err
 CRV.PrintReport
 
 Exit Sub
err:
   MsgBox "Error :" & err.Number & vbNewLine & err.Description
End Sub

Private Sub cmdRefresh_Click()
If Not cmbAccNo.MatchFound Then
   MsgBox "Select An Account First", vbInformation
   Exit Sub
End If
   Call ShowLedger
End Sub

Private Sub DtFrom_Change()
    Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
       Call setDateRange(DtTo)
End Sub

Private Sub Form_Activate()
'Check For User Access
If PermissionDenied Then
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   Exit Sub
End If
End Sub

Private Sub Form_Load()

    'PermissionDenied = Not UserHasAccess("MonthlyLedger")
   
   
    DtTo = Date
    DtFrom.Day = 1

    DtFrom.year = year(DateAdd("M", -3, Date))
    DtFrom.Month = Month(DateAdd("M", -3, Date))
    Call AddToCombo(cmbAccNo, "Code", "Heads", "   Order By Code")
    Call AddToCombo(CmbAccTitle, "Head", "Heads", "   Order By Code")
    
    Call AddToCombo(cmbAccNo, "AccNo", "Accounts", "   Order By AccNo", True)
    Call AddToCombo(CmbAccTitle, "AccTitle", "Accounts", "   Order By AccNo", True)
    
    Set rpt = rptApp.OpenReport(RptPath & "\monthLedger.rpt")
    
End Sub

Private Sub Form_Resize()
    FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
    CRV.Move 50, FAcc.Top + FAcc.Height + 50, Me.ScaleWidth - 100, Me.ScaleHeight - FAcc.Height - cmdClose.Height - 200
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
    cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
End Sub

Private Sub List1_Click()
If UCase(Screen.ActiveControl.Name) <> "CMBACCTITLE" Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If
End Sub
Public Sub ShowLedger()
   Screen.MousePointer = vbHourglass
   
   
   
   'Rpt.Database.Tables(1).Location = DBName
  ' Rpt.Database.Verify
   rpt.Database.Verify
   
   
 'For i = 1 To Rpt.FormulaFields.Count - 1
 '  Debug.Print i, Rpt.FormulaFields(i).Name
 'Next
 
 rpt.RecordSelectionFormula = "{VLedger.AccNo} like '" & cmbAccNo & "*' AND ({VLedger.VDate}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
   
   OpenBal = GetOpenBal(cmbAccNo, DtFrom)

 'Company Name
    rpt.FormulaFields(3).Text = "'" & strCompany & "'"
    rpt.FormulaFields(5).Text = "'Ledger of " & CmbAccTitle & "'"
    rpt.FormulaFields(7).Text = OpenBal
 'Date Range
    rpt.FormulaFields(4).Text = "'From " & Format(DtFrom, "MMMM yyyy") & " To " & Format(DtTo, "MMMM yyyy") & "'"
     
     

CRV.ReportSource = rpt
On Error Resume Next
CRV.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
On Error GoTo 0
CRV.ViewReport
CRV.Zoom 90
Screen.MousePointer = vbDefault

End Sub


Private Function GetOpenBal(AccNo As String, Dated As Date) As Double
   Dim rs As New ADODB.Recordset
   With rs
      .Open "Select accno From Accounts Where (AccNo='" & AccNo & "' OR Ltrim(Accno) Like '" & AccNo & "%') and Parent=0", con, adOpenForwardOnly, adLockReadOnly
      For i = 0 To .RecordCount - 1
         GetOpenBal = GetOpenBal + getBalance(![AccNo], Dated - 1)
         .MoveNEXT
      Next
      .Close
   End With
   Set rs = Nothing
End Function
