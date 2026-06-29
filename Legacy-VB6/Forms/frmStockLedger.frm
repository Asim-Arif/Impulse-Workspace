VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmStockLedger 
   Caption         =   "RM Ledger"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13395
   Icon            =   "frmStockLedger.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   360
      Left            =   11535
      TabIndex        =   10
      Top             =   2055
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   635
      ButtonWidth     =   1455
      ButtonHeight    =   635
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Change Printer Settings"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbCat 
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   435
      Width           =   2190
      _ExtentX        =   3863
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
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   2295
      TabIndex        =   8
      Top             =   435
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   18
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockLedger.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   5835
      Left            =   90
      TabIndex        =   6
      Top             =   840
      Width           =   10530
      lastProp        =   500
      _cx             =   18574
      _cy             =   10292
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   11805
      Begin VB.ComboBox cmbRange 
         Height          =   315
         ItemData        =   "frmStockLedger.frx":0528
         Left            =   5775
         List            =   "frmStockLedger.frx":053E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1740
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   9030
         TabIndex        =   2
         Top             =   420
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   89194499
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   7515
         TabIndex        =   3
         Top             =   420
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   56360963
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmStockLedger.frx":058E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Index           =   7
         Left            =   60
         TabIndex        =   5
         Top             =   135
         Width           =   10635
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   10590
         TabIndex        =   4
         Top             =   405
         Width           =   1170
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2064;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   9105
      TabIndex        =   7
      Top             =   5370
      Width           =   1335
      ForeColor       =   0
      Caption         =   "Close     "
      PicturePosition =   327683
      Size            =   "2355;661"
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
Attribute VB_Name = "frmStockLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As CRAXDDRT.Report

Private Sub cmbCat_matched()
    cmbMaterial.ClearVals
    If cmbCat.ID = "0" Then
        Call cmbMaterial.AddVals(con, "'{' + RMID1 + '} ' + RMName", "VRM", "RMID1", " ORDER BY RMID1")
    Else
        cmbMaterial.AddItem "<All Materials>", "0"
        Call cmbMaterial.AddVals(con, "'{' + RMID1 + '} ' + RMName", "VRM", "RMID1", " WHERE GroupID=" & cmbCat.ID & " ORDER BY RMID1")
        cmbMaterial.ID = "0"
    End If
End Sub

Private Sub cmbRange_Click()
   
   DtTo = Date
   DtFrom.Enabled = False
   DtTo.Enabled = False
   
 Select Case cmbRange.ListIndex
   Case 0
      DtFrom = Date
   Case 1
      DtFrom = DateAdd("d", -7, Date)
   Case 2
      DtFrom = DateAdd("d", -15, Date)
   Case 3
      DtFrom = DateAdd("d", -30, Date)
    Case 4
      DtFrom = DateAdd("d", -60, Date)
   Case 5
      DtFrom.Enabled = True
      DtTo.Enabled = True
      DtFrom.SetFocus
      Exit Sub
 End Select
 
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdRefresh_Click()

    'set Flag for Refresh
        
    Me.Caption = "Stock Ledger"
    
    
    Set Report = rptApp.OpenReport(RptPath & "\StockLedger.rpt")
    
    On Error Resume Next
    Report.Database.Tables(1).ConnectBufferString = strConnectionStringrpt
    
    Screen.MousePointer = vbHourglass
    
    With Report
        If cmbMaterial.ID = "0" Then
            .RecordSelectionFormula = "({VStockLedger.TransactionDate} = Date(" & Format(DtFrom, "yyyy,MM,dd") & ") to  Date(" & Format(DtTo, "yyyy,MM,dd") & "))"
        Else
            .RecordSelectionFormula = "{VSTockLedger.MaterialID}='" & cmbMaterial.ID & "' and ({VStockLedger.TransactionDate} = Date(" & Format(DtFrom, "yyyy,MM,dd") & ") to  Date(" & Format(DtTo, "yyyy,MM,dd") & "))"
        End If
        .FormulaFields.GetItemByName("FromTo").Text = """From   " & Format(DtFrom, "dd-MMM-yyyy") & "    to     " & Format(DtTo, "dd-MMM-yyyy") & """"
        .FormulaFields.GetItemByName("OpeningBalance").Text = GetOpeningBalance()
        .DiscardSavedData
    End With
    
    CRViewer1.ReportSource = Report
    
    Me.Show
    Toolbar1.Visible = True
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault
    'Set Report = Nothing

End Sub


Private Sub Form_Load()
   
    'cmbRange.ListIndex = 0
    Sign = 1
    
    cmbMaterial.ListHeight = 5600
    
    cmbCat.ListHeight = 5600
    cmbCat.AddItem "<All Groups>", "0"
    Call cmbCat.AddVals(con, "Description", "RMGroups", "ID", "Order By Description")
    cmbCat.ID = "0"
    
    cmbRange.ListIndex = 4
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Dim TitleBottom As Long
  
    TitleBottom = FAcc.Top + FAcc.Height
    FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
    
    cmbCat.Move FAcc.Left + 60, FAcc.Top + 420
    cmbMaterial.Move cmbCat.Left + cmbCat.Width + 50, cmbCat.Top
    CRViewer1.Move 50, TitleBottom, ScaleWidth - 100, ScaleHeight - TitleBottom - cmdClose.Height - 100
    cmdClose.Move ScaleWidth - cmdClose.Width - 50, ScaleHeight - cmdClose.Height - 50
    Toolbar1.Move 8350, 830
    
End Sub

Private Function GetOpeningBalance() As Double
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim OpeningBalance As Double
    
    With rs
        .Open "Select Sum(Quantity) from VSTockLedger Where MAterialID='" & cmbMaterial.ID & "' and TransactionDAte < '" & Format(DtFrom, "dd-MMM-yyyy") & "'", con, adOpenForwardOnly, adLockReadOnly
        OpeningBalance = Val(.Fields(0) & "")
        .Close
        .Open "Select OpeningStock From RM Where RMID1='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
         If Not .EOF Then OpeningBalance = OpeningBalance + Val(.Fields(0) & "")
         .Close
    End With
    
    Set rs = Nothing
    GetOpeningBalance = Round(OpeningBalance, 2)
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Report.PrinterSetup Me.hwnd
End Sub


