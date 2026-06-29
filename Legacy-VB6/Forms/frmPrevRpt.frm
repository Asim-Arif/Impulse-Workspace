VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmPrevRpt 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   7095
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   360
      Left            =   1980
      TabIndex        =   0
      Top             =   1065
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1125
      Top             =   4560
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
            Picture         =   "frmPrevRpt.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   7005
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   7065
      lastProp        =   500
      _cx             =   12462
      _cy             =   12356
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
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   5520
      TabIndex        =   1
      Top             =   7200
      Width           =   1425
      ForeColor       =   0
      Caption         =   " Close      "
      PicturePosition =   327683
      Size            =   "2514;635"
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
Attribute VB_Name = "frmPrevRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SrcReport As CRAXDDRT.Report
Dim frmToShowAtUnload As Form
Dim bTrackPrinting As Boolean, strTableName As String, strFieldName As String, strValue As String, bAllowDuplicate As Boolean
Dim bAlreadyPrinted As Boolean


Public Sub ShowReport(strSel As String, ArgSourceRpt As CRAXDDRT.Report, Optional strQueryString As String, Optional bDiscardSaved As Boolean = True, Optional strPageName As String = "", Optional p_frmToShowAtUnload As Form, Optional p_bTrackPrinting As Boolean = False, Optional p_strTableName As String = "", Optional p_strFieldName As String = "", Optional p_strValue As String = "", Optional p_bAllowDuplicate As Boolean = False)
     
    Set SrcReport = ArgSourceRpt
    
    If p_bTrackPrinting Then
        bTrackPrinting = p_bTrackPrinting
        strTableName = p_strTableName
        strFieldName = p_strFieldName
        strValue = p_strValue
        bAllowDuplicate = p_bAllowDuplicate
        
        Dim lCount As Long
        lCount = GetSingleLongValue("COUNT(" & strFieldName & ")", strTableName, " WHERE " & strFieldName & "='" & strValue & "'")
        If lCount > 0 Then
            bAlreadyPrinted = True
            If bAllowDuplicate = False Then
'                CRViewer1.EnablePrintButton = False
'                CRViewer1.EnableExportButton = False
            End If
            SrcReport.FormulaFields.GetItemByName("DuplicateMark").Text = "'Duplicate'"
        End If
'        CRViewer1.EnableExportButton = False
    End If
    
    strSel = strSel & " "
    
    On Error Resume Next
    SrcReport.Database.Tables(1).ConnectBufferString = strConnectionStringrpt
    
    Screen.MousePointer = vbHourglass
     
    If Not (p_frmToShowAtUnload Is Nothing) Then
        Set frmToShowAtUnload = p_frmToShowAtUnload
    End If
    
    If bDiscardSaved Then
        SrcReport.DiscardSavedData
    End If
    If strSel <> "" Then SrcReport.RecordSelectionFormula = Left(strSel, Len(strSel) - 1)
    SrcReport.RecordSelectionFormula = strSel '& "'"
    If strQueryString <> "" Then
        SrcReport.SQLQueryString = strQueryString
    End If
    
    If strPageName <> "" Then
        Dim myPageID As Long
        myPageID = getPageID(strPageName)
        If myPageID <> 0 Then
            SrcReport.PaperSize = myPageID
        End If
    End If
    
    CRViewer1.ReportSource = SrcReport
    'CRViewer1.ReportSource = rptFProforma
   'ÚÇÕã ÚÇÑÝ
'    SrcReport.Database.Verify
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault
    Me.Show
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    'SrcReport.Export
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    On Error Resume Next
    If bTrackPrinting Then
        
        If bAlreadyPrinted = False Then
            con.Execute "INSERT INTO " & strTableName & "(" & strFieldName & ",UserName,MachineName) VALUES('" & _
             strValue & "','" & CurrentUserName & "','" & strComputerName & "')"
        End If
    End If
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    Set SrcReport = Nothing
    Set frmPrevRpt = Nothing
End Sub

Private Sub Form_Resize()

    With CRViewer1
        .Top = 0
        .Left = 0
        .Height = ScaleHeight - 500
        .Width = ScaleWidth
        cmdClose.Move .Width - cmdClose.Width, .Height + 100
    End With
    
    Toolbar1.Move 5850, 50
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    If Not (frmToShowAtUnload Is Nothing) Then
        frmToShowAtUnload.Visible = True
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    SrcReport.PrinterSetup Me.hwnd
End Sub
