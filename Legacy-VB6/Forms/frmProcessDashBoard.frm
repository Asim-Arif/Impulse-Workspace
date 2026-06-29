VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmProcessDashBoard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   10560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcessDashBoard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtWeek4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   10875
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   10110
      Width           =   1200
   End
   Begin VB.TextBox txtWeek3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   10110
      Width           =   1200
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   6015
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   10110
      Width           =   1200
   End
   Begin VB.TextBox txtWeek1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   7230
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   10110
      Width           =   1200
   End
   Begin VB.TextBox txtWeek2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   8445
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   10110
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10110
      Width           =   1215
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   9495
      Index           =   0
      Left            =   60
      ScaleHeight     =   9495
      ScaleWidth      =   12120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   12120
      Begin VB.CheckBox chkRemember 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Remember"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9165
         TabIndex        =   11
         Top             =   120
         Width           =   1605
      End
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   330
         Left            =   30
         TabIndex        =   6
         Top             =   420
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   582
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView LV 
         Height          =   8640
         Left            =   30
         TabIndex        =   3
         Top             =   810
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   15240
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16761024
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Maker"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Balance"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Week 1"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Week 2"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Week 3"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Week 4"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   420
         Left            =   10860
         TabIndex        =   5
         Top             =   375
         Width           =   1230
         ForeColor       =   65280
         BackColor       =   4210752
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2170;741"
         Accelerator     =   82
         FontName        =   "Calibri"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   4
         Top             =   210
         Width           =   615
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   5295
      TabIndex        =   7
      Top             =   10170
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESS DASHBOARD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   12120
   End
End
Attribute VB_Name = "frmProcessDashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(220, 219, 244) 'RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    Dim lBorderColor As Long, lForeColor As Long
    For i = 0 To PicWC.count - 1
        If i = 1 Then
            lBorderColor = RGB(153, 204, 255)
            lForeColor = 0
        ElseIf i = 0 Then
            lBorderColor = RGB(220, 219, 244) 'RGB(255, 204, 153)
            lForeColor = 0
        ElseIf i = 3 Then
            lBorderColor = RGB(255, 153, 204)
            lForeColor = 0
        ElseIf i = 2 Then
            lBorderColor = RGB(0, 255, 0)
            lForeColor = 0
        ElseIf i = 4 Then
            lBorderColor = RGB(200, 200, 192)
            lForeColor = 0
        End If
        Call DrawCaptions(PicWC(i), lBorderColor, lForeColor)
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox, lBorderColor As Long, lForeColor As Long)

    Dim lFillColor
    'lBorderColor = 0 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    lFillColor = lBorderColor 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 150)-(Pic.Width, 150), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 300), lFillColor, BF
        Pic.ForeColor = lForeColor 'vbGreen
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 0
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 50)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()

    On Error GoTo err
   
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim dTotal As Double
    
    LV.ListItems.Clear
    dTotal = 0
    
    Dim bFactoryMaker As Boolean
    
    Dim lForeColor As Long, i As Integer, dRate As Double
    Dim lTotalBalanceCount As Long, lTotalBalanceQty As Long
    Dim lWeek1Qty As Long, lWeek1Count As Long
    Dim lWeek2Qty As Long, lWeek2Count As Long
    Dim lWeek3Qty As Long, lWeek3Count As Long
    Dim lWeek4Qty As Long, lWeek4Count As Long
    
    Dim DTFrom1 As Date, DTTo1 As Date
    Dim DTFrom2 As Date, DTTo2 As Date
    Dim DTFrom3 As Date, DTTo3 As Date
    Dim DTFrom4 As Date, DTTo4 As Date
     
    DTFrom1 = DateAdd("d", -30, Date)
    DTTo1 = DateAdd("d", 7, DTFrom1)
    Dim DTTemp As Date
    
    DTTemp = DateAdd("d", -Weekday(Date) + 1, Date) 'Last Sunday
    DTTo4 = DTTemp
    DTFrom4 = DateAdd("d", -6, DTTo4)
    
    DTTo3 = DateAdd("d", -1, DTFrom4)
    DTFrom3 = DateAdd("d", -6, DTTo3)
    
    DTTo2 = DateAdd("d", -1, DTFrom3)
    DTFrom2 = DateAdd("d", -6, DTTo2)
    
    DTTo1 = DateAdd("d", -1, DTFrom2)
    DTFrom1 = DateAdd("d", -6, DTTo1)
    With rs
        .Open "SELECT Makers.VendID,Makers.VenderName,BalanceQty,ItemCount,Week1Qty,Week1Count,Week2Qty,Week2Count,Week3Qty," & _
            "Week3Count,Week4Qty,Week4Count FROM Makers INNER JOIN (SELECT VendID,SUM(IssQty-RcvdQty) AS BalanceQty" & _
            ",COUNT(DISTINCT ItemCode) AS ItemCount FROM VVendIssdDetail WHERE ProcessID=" & cmbProcess.ID & _
            " AND IssQty>RcvdQty GROUP BY VendID,VenderName) T1 ON Makers.VendID=T1.VendID LEFT OUTER JOIN " & _
            "(SELECT VendID,SUM(RcvdQty) AS Week1Qty,COUNT(DISTINCT ItemCode) AS Week1Count FROM VVendRcvdDetail_Simple " & _
            "WHERE ProcessID=" & cmbProcess.ID & " AND DT BETWEEN '" & DTFrom1 & "' AND '" & DTTo1 & "' GROUP BY VendID) " & _
            "TWeek1 ON Makers.VendID=TWeek1.VendID LEFT OUTER JOIN " & _
            "(SELECT VendID,SUM(RcvdQty) AS Week2Qty,COUNT(DISTINCT ItemCode) AS Week2Count FROM VVendRcvdDetail_Simple " & _
            "WHERE ProcessID=" & cmbProcess.ID & " AND DT BETWEEN '" & DTFrom2 & "' AND '" & DTTo2 & "' GROUP BY VendID) " & _
            "TWeek2 ON Makers.VendID=TWeek2.VendID LEFT OUTER JOIN " & _
            "(SELECT VendID,SUM(RcvdQty) AS Week3Qty,COUNT(DISTINCT ItemCode) AS Week3Count FROM VVendRcvdDetail_Simple " & _
            "WHERE ProcessID=" & cmbProcess.ID & " AND DT BETWEEN '" & DTFrom3 & "' AND '" & DTTo3 & "' GROUP BY VendID) " & _
            "TWeek3 ON Makers.VendID=TWeek3.VendID LEFT OUTER JOIN " & _
            "(SELECT VendID,SUM(RcvdQty) AS Week4Qty,COUNT(DISTINCT ItemCode) AS Week4Count FROM VVendRcvdDetail_Simple " & _
            "WHERE ProcessID=" & cmbProcess.ID & " AND DT BETWEEN '" & DTFrom4 & "' AND '" & DTTo4 & "' GROUP BY VendID) " & _
            "TWeek4 ON Makers.VendID=TWeek4.VendID", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
             
            Set ITM = LV.ListItems.add(, , !VenderName)
            ITM.Tag = Val(!VendID & "")
           
            ITM.ListSubItems.add(, , Val(!ItemCount & "") & "/" & Val(!BalanceQty & "")).Tag = Val(!BalanceQty & "")
            
            lTotalBalanceCount = lTotalBalanceCount + Val(!ItemCount & "")
            lTotalBalanceQty = lTotalBalanceQty + Val(!BalanceQty & "")
            
            ITM.ListSubItems.add(, , Val(!Week1Count & "") & "/" & Val(!Week1Qty & "")).Tag = Val(!Week1Qty & "")
            lWeek1Qty = lWeek1Qty + Val(!Week1Qty & "")
            lWeek1Count = lWeek1Count + Val(!Week1Count & "")
            
            ITM.ListSubItems.add(, , Val(!Week2Count & "") & "/" & Val(!Week2Qty & "")).Tag = Val(!Week2Qty & "")
            lWeek2Qty = lWeek2Qty + Val(!Week2Qty & "")
            lWeek2Count = lWeek2Count + Val(!Week2Count & "")
            
            ITM.ListSubItems.add(, , Val(!Week3Count & "") & "/" & Val(!Week3Qty & "")).Tag = Val(!Week3Qty & "")
            lWeek3Qty = lWeek3Qty + Val(!Week3Qty & "")
            lWeek3Count = lWeek3Count + Val(!Week3Count & "")
            
            ITM.ListSubItems.add(, , Val(!Week4Count & "") & "/" & Val(!Week4Qty & "")).Tag = Val(!Week4Qty & "")
            lWeek4Qty = lWeek4Qty + Val(!Week4Qty & "")
            lWeek4Count = lWeek4Count + Val(!Week4Count & "")
            
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    txtBalance = lTotalBalanceCount & "/" & lTotalBalanceQty
    txtWeek1 = lWeek1Count & "/" & lWeek1Qty
    txtWeek2 = lWeek2Count & "/" & lWeek2Qty
    txtWeek3 = lWeek3Count & "/" & lWeek3Qty
    txtWeek4 = lWeek4Count & "/" & lWeek4Qty
    'txtTotal = Format(dTotal, "#,##0.00")
    'Call GetTotal
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Function GetTotal() As Double

    Dim i As Integer
    Dim lIssd As Long, lRcvd As Long, lBal As Long
    For i = 1 To LV.ListItems.count
        lIssd = lIssd + Val(LV.ListItems(i).SubItems(1))
        lRcvd = lRcvd + Val(LV.ListItems(i).SubItems(2))
        lBal = lBal + Val(LV.ListItems(i).SubItems(3))
        
    Next
'    txtIssd = lIssd
'    txtRcvd = lRcvd
'    txtBalance = lBal
    
End Function
    
Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    'Show Detail
    Dim f As New frmProcessDashBoardDetails
    Load f
    Call f.ShowIssuanceList(LV.SelectedItem.Tag, cmbProcess.ID)
    
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = x
                lDownY = y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
End Sub

Public Function ShowMe() As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
'    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
'    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show
    
    cmbProcess.ListHeight = 5800
    Call cmbProcess.AddVals(con, "'{ ' + Code + ' } ' + Description ", "Processes", "ProcessID")
      
    
    Me.Show
    ShowMe = bSaved
    
End Function
