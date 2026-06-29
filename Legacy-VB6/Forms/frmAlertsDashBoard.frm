VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAlertsDashBoard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   10560
   ClientLeft      =   3705
   ClientTop       =   465
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlertsDashBoard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
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
      ScaleWidth      =   11310
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   11310
      Begin MSComctlLib.ListView LV 
         Height          =   4110
         Left            =   30
         TabIndex        =   4
         Top             =   5355
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   7250
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Alert"
            Object.Width           =   17639
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hidden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Hidden"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.MonthView MV1 
         Height          =   5310
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   9366
         _Version        =   393216
         ForeColor       =   8388608
         BackColor       =   4210752
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthColumns    =   2
         ShowToday       =   0   'False
         StartOfWeek     =   114819073
         TitleBackColor  =   4210752
         TitleForeColor  =   16777215
         CurrentDate     =   40887
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALERTS"
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
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   11235
   End
End
Attribute VB_Name = "frmAlertsDashBoard"
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
    
Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
   
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.ListSubItems(1).Tag = "" Then Exit Sub
    'If Val(LV.SelectedItem.ListSubItems(2).Tag) > 0 Then Exit Sub
    
    If Left(LV.SelectedItem.SubItems(1), 3) = "(D)" Then
        Exit Sub
    End If
    Dim f As New frmAddFollowupData
    Load f
    Call f.AddFollowup(Val(LV.SelectedItem.Tag), LV.SelectedItem.ListSubItems(1).Tag, Val(LV.SelectedItem.ListSubItems(2).Tag))
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print Master PO Status")
    If strRet = "Print Master PO Status" Then
        Call PrintMasterPOStatus(LV.SelectedItem.ListSubItems(3).Tag)
    End If
    
End Sub

Private Sub MV1_DateClick(ByVal DateClicked As Date)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strAlert As String, dBalance As Double, lForeColor As Long, i As Integer
    With rs
        .Open "SELECT EntryIDReturnDTs,ItemName,Makers.VenderName,ItemCode,IssQty,RcvQty,Makers.VendID1,Prefix,FollowUpData.EntryID AS FollowUpEntryID,MasterPONO FROM " & _
        "VPurchaseCalendar INNER JOIN Items ON VPurchaseCalendar.ItemCode=Items.ItemID INNER JOIN Makers ON " & _
        "VPurchaseCalendar.VendID=Makers.VendID LEFT OUTER JOIN FollowUpData ON EntryIDReturnDTs=VIDRDTS_RefID " & _
        "AND Prefix=FollowUpData.FollowUp WHERE ReturnDT='" & DateClicked & "' AND IssQty>RcvQty " & _
        "AND CustCode<>'Stock' AND MasterPONo IS NOT NULL AND Makers.VendID<>207 AND CustCode IN(SELECT CustCode FROM Users_Customers " & _
        "WHERE UserID=" & UserID & ")", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryIDReturnDTs & "")
            
            strAlert = ""
            If !Prefix <> "" Then
                strAlert = !Prefix & ""
            End If
            dBalance = Val(!IssQty & "") - Val(!RcvQty & "")
            'strAlert = strAlert & "(" & !ItemCode & ") " & !ItemName & " / " & dBalance & " / " & "(" & !VendID1 & ") " & !VenderName
            strAlert = strAlert & " " & dBalance & " Pcs of " & !ItemCode & " (" & !ItemName & ") from " & !VendID1 & " (" & !VenderName & ")"
            ITM.ListSubItems.add(, , strAlert).Tag = !Prefix & ""
            ITM.ListSubItems.add(, , "").Tag = Val(!FollowUpEntryID & "")
            ITM.ListSubItems.add(, , "").Tag = !MasterPONo & ""
            
            If Val(!FollowUpEntryID & "") = 0 Then
                If !Prefix & "" = "(D)" Then
                    lForeColor = RGB(50, 50, 255)
                Else
                    lForeColor = RGB(255, 50, 50)
                End If
            Else
                lForeColor = RGB(100, 200, 100)
            End If
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub MV1_GetDayBold(ByVal StartDate As Date, ByVal count As Integer, State() As Boolean)

    Dim LastDate As Date
    LastDate = DateAdd("d", (count - 1), StartDate)
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ReturnDT FROM VPurchaseCalendar WHERE IssQty>RcvQty AND CustCode<>'Stock' AND MasterPONo IS NOT NULL " & _
            "AND (ReturnDT BETWEEN '" & StartDate & "' AND '" & LastDate & "') AND ReturnDT IS NOT NULL AND VendID<>207 AND CustCode IN(SELECT CustCode FROM Users_Customers WHERE UserID=" & UserID & ") GROUP BY ReturnDT", con, adOpenForwardOnly, adLockReadOnly

        Do Until .EOF
            State(DateDiff("d", StartDate, !ReturnDT)) = True
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
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
    
    MV1.Value = Date
    Call MV1_DateClick(Date)
    
    Me.Show
    ShowMe = bSaved
    
End Function
