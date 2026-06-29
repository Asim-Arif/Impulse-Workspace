VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonthlyProductionDashBoard 
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14385
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleMode       =   0  'User
   ScaleWidth      =   12884.84
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   14325
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Monthwise Production"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Index           =   8
         Left            =   75
         TabIndex        =   3
         Top             =   165
         Width           =   14175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Monthwise Production"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   9
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   14160
      End
   End
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9675
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   14325
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11700
         TabIndex        =   9
         Top             =   9165
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker DTMonth 
         Height          =   315
         Left            =   45
         TabIndex        =   7
         Top             =   360
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
         CustomFormat    =   "yyyy"
         Format          =   100073475
         CurrentDate     =   37055
      End
      Begin MSComctlLib.ListView LV 
         Height          =   8430
         Left            =   45
         TabIndex        =   8
         Top             =   705
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   14870
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16711680
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Month"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Receiving"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "On Time"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "QC"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Packing"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Demand"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Supply"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   1620
         TabIndex        =   6
         Top             =   300
         Width           =   1050
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1852;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   435
         Left            =   13080
         TabIndex        =   5
         Top             =   9165
         Width           =   1140
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2011;767"
         Accelerator     =   67
         FontName        =   "Calibri"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Year"
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
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   1
         Top             =   120
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmMonthlyProductionDashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ProductionTimeLineReport_Monthly.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DTMonth.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DTMonth.Value
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, "", False
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Form_Load()
    
    DTMonth.Value = Date
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Private Sub RefreshLV()

    On Error GoTo err
    
    Dim cmd As New ADODB.Command, rs As New ADODB.Recordset
    Dim myDTFrom As Date, myDTTo As Date
    myDTFrom = DateSerial(DTMonth.year, 1, 1)
    myDTTo = DateSerial(DTMonth.year, 12, 31)
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "MonthWiseProductionReport_SP"
        .CommandTimeout = 120
        .Parameters("@DTFrom").Value = myDTFrom
        .Parameters("@DTTo").Value = myDTTo
        .Parameters("@Summary").Value = True
        Set rs = .Execute
    End With
    
    Dim lTotalRcvdQty As Long, lTotalRcvdItems As Long
    Dim lTotalQCQty As Long, lTotalQCItems As Long
    
    Dim lTotalPackQty As Long, lTotalPackItems As Long
    Dim lTotalOrderQty As Long, lTotalOrderItems As Long
    Dim lTotalInvoiceQty As Long, lTotalInvoiceItems As Long
    
    Dim ITM As ListItem
    With rs
        LV.ListItems.Clear
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , Format(!DateFull, "MMM"))
            ITM.Tag = Format(!DateFull, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Val(!TotalRcvdItems & "") & "/" & Val(!TotalRcvdQty & "")
            ITM.ListSubItems.add , , Round(Val(!RcvdOnTime & ""), 2)
            
            ITM.ListSubItems.add , , Val(!TotalQCItems & "") & "/" & Val(!TotalQCQty & "")
            ITM.ListSubItems.add , , Val(!TotalPackItems & "") & "/" & Val(!TotalPackQty & "")
            ITM.ListSubItems.add , , Val(!TotalOrderItems & "") & "/" & Val(!TotalOrderQty & "")
            ITM.ListSubItems.add , , Val(!TotalInvoiceItems & "") & "/" & Val(!TotalInvoiceQty & "")
            
            lTotalRcvdQty = lTotalRcvdQty + Val(!TotalRcvdQty & "")
            lTotalRcvdItems = lTotalRcvdItems + Val(!TotalRcvdItems & "")
            
            lTotalQCQty = lTotalQCQty + Val(!TotalQCQty & "")
            lTotalQCItems = lTotalQCItems + Val(!TotalQCItems & "")
            
            
            lTotalPackQty = lTotalPackQty + Val(!TotalPackQty & "")
            lTotalPackItems = lTotalPackItems + Val(!TotalPackItems & "")
            
            lTotalOrderQty = lTotalOrderQty + Val(!TotalOrderQty & "")
            lTotalOrderItems = lTotalOrderItems + Val(!TotalOrderItems & "")
            
            lTotalInvoiceQty = lTotalInvoiceQty + Val(!TotalInvoiceQty & "")
            lTotalInvoiceItems = lTotalInvoiceItems + Val(!TotalInvoiceItems & "")
            
            .MoveNext
        Loop
    End With
    Set rs = Nothing
    
    Set ITM = LV.ListItems.add(, , "Total")
    ITM.ListSubItems.add , , lTotalRcvdItems & "/" & lTotalRcvdQty
    ITM.ListSubItems.add , , ""
    ITM.ListSubItems.add , , lTotalQCItems & "/" & lTotalQCQty
    ITM.ListSubItems.add , , lTotalPackItems & "/" & lTotalPackQty
    ITM.ListSubItems.add , , lTotalOrderItems & "/" & lTotalOrderQty
    ITM.ListSubItems.add , , lTotalInvoiceItems & "/" & lTotalInvoiceQty
    
    ITM.Bold = True
    
    Dim i As Integer
    For i = 1 To ITM.ListSubItems.count
        ITM.ListSubItems(i).Bold = True
    Next
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.Index = LV.ListItems.count Then Exit Sub
    
    Dim myDTFrom As Date, myDTTo As Date
    myDTFrom = DateSerial(DTMonth.year, LV.SelectedItem.Index, 1)
    myDTTo = DateSerial(DTMonth.year, LV.SelectedItem.Index + 1, 0)
    
    Call PrintMonthlyDateWiseProductionReport(False, myDTFrom, myDTTo)
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > Val(LV.ColumnHeaders(i).Left) And X < (Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width)) Then
            iColNo = i
            Exit For
        End If
    Next
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    Dim strRet As String
    
    strRet = ShowPopUpMenu1("Demand And Supply Graph")
    
    If strRet = "Demand And Supply Graph" Then
        Dim myDT As Date
        myDT = LV.SelectedItem.Tag
        Call PrintDemandAndSupplyGraph(myDT)
    End If
    
End Sub
