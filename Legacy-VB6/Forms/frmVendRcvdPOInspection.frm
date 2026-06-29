VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmVendRcvdPOInspection 
   BorderStyle     =   0  'None
   Caption         =   "Receiving Lots"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14580
   DrawWidth       =   2
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
   ScaleHeight     =   8895
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00ECE0E5&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8895
      Index           =   0
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   14565
      TabIndex        =   1
      Tag             =   "LOTS RECEIVING"
      Top             =   0
      Width           =   14565
      Begin ComboList.Usercontrol1 cmbMaterial 
         Height          =   390
         Left            =   75
         TabIndex        =   26
         Top             =   1860
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   688
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBackColor   =   16777215
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   12210
         TabIndex        =   24
         Top             =   1860
         Width           =   1920
      End
      Begin VB.OptionButton OptTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         Caption         =   "Inspection"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3150
         Width           =   2355
      End
      Begin VB.TextBox txtOrderNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   5730
         TabIndex        =   0
         Top             =   1035
         Width           =   2010
      End
      Begin VB.TextBox txtRcvdQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   75
         TabIndex        =   4
         Top             =   2700
         Width           =   1950
      End
      Begin VB.TextBox txtVender 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   2040
         TabIndex        =   3
         Top             =   2700
         Width           =   9285
      End
      Begin VB.TextBox txtDT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   11355
         TabIndex        =   2
         Top             =   2700
         Width           =   2775
      End
      Begin VB.PictureBox PicTabs 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE0E5&
         ForeColor       =   &H80000008&
         Height          =   5085
         Index           =   0
         Left            =   75
         ScaleHeight     =   5055
         ScaleWidth      =   14190
         TabIndex        =   10
         Top             =   3720
         Width           =   14220
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00ECE0E5&
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   10155
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   4245
            Width           =   1290
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            Left            =   2475
            TabIndex        =   16
            Top             =   4455
            Width           =   4095
         End
         Begin VB.TextBox txtDisposation 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            Left            =   30
            TabIndex        =   15
            Top             =   4455
            Width           =   2415
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7110
            TabIndex        =   14
            Top             =   885
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.CommandButton cmdSaveInspection 
            BackColor       =   &H00ECE0E5&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   11505
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   4245
            Width           =   1290
         End
         Begin VB.OptionButton OptInspection 
            BackColor       =   &H00ECE0E5&
            Caption         =   "Pass"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   8220
            TabIndex        =   12
            Tag             =   "-1"
            Top             =   4485
            Width           =   1035
         End
         Begin VB.OptionButton OptInspection 
            BackColor       =   &H00ECE0E5&
            Caption         =   "Fail"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   9225
            TabIndex        =   11
            Top             =   4485
            Width           =   990
         End
         Begin MSComCtl2.DTPicker DTInspection 
            Height          =   390
            Left            =   6600
            TabIndex        =   18
            Top             =   4455
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   116916227
            CurrentDate     =   40357
         End
         Begin MSComctlLib.ListView LV1 
            Height          =   4110
            Left            =   30
            TabIndex        =   19
            Top             =   30
            Width           =   14100
            _ExtentX        =   24871
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
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Attributes Description"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "AQL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Sample Size"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Reject On"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Test Spec. No."
               Object.Width           =   4339
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Act. Rej."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Status"
               Object.Width           =   4586
            EndProperty
         End
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H00ECE0E5&
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   4245
            Width           =   1290
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Insp."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   7
            Left            =   6600
            TabIndex        =   22
            Top             =   4125
            Width           =   1110
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "General Remarks"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   6
            Left            =   2490
            TabIndex        =   21
            Top             =   4125
            Width           =   1920
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Material Disposation"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   8
            Left            =   15
            TabIndex        =   20
            Top             =   4125
            Width           =   2310
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   18
         Left            =   12195
         TabIndex        =   25
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   5985
         TabIndex        =   9
         Top             =   630
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   1515
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rcvd. Qty"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   75
         TabIndex        =   7
         Top             =   2355
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   2025
         TabIndex        =   6
         Top             =   2355
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   11340
         TabIndex        =   5
         Top             =   2355
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmVendRcvdPOInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const lCaptionHeight = 600
Dim bMouseDown As Boolean
Dim lDownX As Long, lDownY As Long

Dim iColNo As Integer
Dim strRcvID As String, lGroupID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    For i = 0 To PicWC.count - 1
        Call DrawCaptions(PicWC(i))
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox)

    Dim lBorderColor As Long, lFillColor
    lBorderColor = RGB(128, 100, 162)
    lFillColor = RGB(128, 100, 162)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 350)-(Pic.Width, 350), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 600), lFillColor, BF
        Pic.ForeColor = vbWhite
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 180
        Pic.Print Pic.Tag
    End If
        
    Pic.Line (0, Pic.Height - 20)-(Pic.Width, Pic.Height - 20), lBorderColor 'Bottom

    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 20, Pic.Height), lBorderColor  'Right

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmbMaterial_matched()

    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim strMaterialID As String
    With rs
        
        .Open "SELECT RMSize,GroupID,QtyRcvd FROM VendRcvdDetailPO INNER JOIN VRMAndItemsUnioned ON VendRcvdDetailPO.MaterialID=VRMAndItemsUnioned.RMID1 WHERE EntryID=" & cmbMaterial.ID, con, adOpenForwardOnly, adLockReadOnly
        txtRcvdQty = Val(!QtyRcvd & "")
        txtSize = !RMSize & ""
        lGroupID = Val(!GroupID & "")
        .Close
        
        .Open "SELECT * FROM RMGroupParameters WHERE GroupID=" & lGroupID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        Do Until .EOF
            Set ITM = LV1.ListItems.add(, !EntryID & "'", !parameterName & "")
            ITM.ListSubItems.add , , !AQL & ""
            ITM.ListSubItems.add , , !SampleSize & ""
            ITM.ListSubItems.add , , !RejectOn & ""
            ITM.ListSubItems.add , , !TechSpecNo & "" 'cmbProcessNo.List(cmbProcessNo.ListIndex, 2)
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "OK"
            .MoveNEXT
        Loop
        .Close
                
        .Open "SELECT * FROM VendRcvdDetailPO_Inspection WHERE VRDPO_RefID=" & cmbMaterial.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            LV1.Tag = Val(!EntryID & "")
            txtDisposation = !Disposation & ""
            txtRemarks = !Comments & ""
            DTInspection = !DT
            OptInspection(Abs(Not !Status)).Value = True
            '------------------------------------
        Else
            LV1.Tag = "0"
            txtDisposation = ""
            txtRemarks = ""
            DTInspection = Date
        End If
        .Close
        
        If Val(LV1.Tag) > 0 Then
            .Open "SELECT * FROM VendRcvdDetailPO_InspectionDetail WHERE VRDPO_Insp_RefID=" & Val(LV1.Tag), con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                Set ITM = LV1.ListItems(!RMGP_RefID & "'")
                ITM.SubItems(1) = !AQL & ""
                ITM.SubItems(2) = !SampleSize & ""
                ITM.SubItems(3) = !RejectOn & ""
                ITM.SubItems(4) = !TestSpecificationNo & ""
                ITM.SubItems(5) = !ActualRejection & ""
                ITM.SubItems(6) = !Status & ""
                .MoveNEXT
            Loop
            .Close
        End If
             
'        .Open "SELECT * FROM IPOTemperValues WHERE ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & " AND IPO_RefID=" & txtLotBarcode & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
'        LVTemper.ListItems.Clear
'        Do Until .EOF
'            Set ITM = LVTemper.ListItems.add(, !EntryID & "'", Val(!SNo & ""))
'            ITM.ListSubItems.add , , !TempValue & ""
'            .MoveNEXT
'        Loop
'        .Close
        
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdPrint_Click()
    
    On Error GoTo err
    'If Val(LV1.Tag) <= 0 Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\VendRcvdDetailPOInspection.rpt")
    Dim f As New frmPrevRpt
    Me.Visible = False
    f.ShowReport "{VVendRcvdDetailPO.EntryID}=" & cmbMaterial.ID, rpt, , , , Me
    
    Exit Sub
err:
    MsgBox err.Description

End Sub


Private Sub cmdSaveInspection_Click()

    On Error GoTo err
    If OptInspection(0).Tag = "-1" Then
        MsgBox "Please Select Pass / Fail", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    
    Dim myDT As Date
    myDT = Date
    Call StartTrans(con)
     
    Dim lEntryID As Long
    
    If Val(LV1.Tag) > 0 Then
        con.Execute "DELETE FROM VendRcvdDetailPO_Inspection WHERE EntryID=" & Val(LV1.Tag)
    End If
     
    con.Execute "INSERT INTO VendRcvdDetailPO_Inspection(VRDPO_RefID,Disposation,Comments,DT,Status,UserName,MachineName) " & _
     "VALUES(" & cmbMaterial.ID & ",'" & txtDisposation & "','" & txtRemarks & "','" & DTInspection & "'," & _
     IIf(OptInspection(0).Tag = "1", 0, 1) & ",'" & CurrentUserName & "','" & strComputerName & "')"
      
    lEntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetailPO_Inspection", " WHERE MachineName='" & strComputerName & "'")
    
    For i = 1 To LV1.ListItems.count
        With LV1.ListItems(i)
            con.Execute "INSERT INTO VendRcvdDetailPO_InspectionDetail(VRDPO_Insp_RefID,RMGP_RefID,AQL,SampleSize,RejectOn," & _
             "TestSpecificationNo,ActualRejection,Status) VALUES(" & lEntryID & "," & Val(.key) & ",'" & _
             .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" & _
             .SubItems(5) & "','" & .SubItems(6) & "')"
        End With
    Next
    
    'Now Save Temper Values....
'    con.Execute "DELETE FROM IPOTemperValues WHERE IPO_RefID=" & txtLotBarcode.Text & " AND ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1)
'    For i = 1 To LVTemper.ListItems.count
'        con.Execute "INSERT INTO IPOTemperValues(IPO_RefID,ProcessID,SNo,TempValue) VALUES(" & txtLotBarcode.Text & _
'         "," & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & "," & i & ",'" & LVTemper.ListItems(i).SubItems(1) & "')"
'    Next
    
    con.CommitTrans
    LV1.Tag = lEntryID
    
    'Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    Call DrawPicCaptions
    bMouseDown = False
    
    cmbMaterial.ListHeight = 4000
    
End Sub


Private Sub LV1_DblClick()
    If LV1.ListItems.count = 0 Then Exit Sub
    If iColNo = 1 Then Exit Sub
    With txtEdit
        .Move LV1.Left + LV1.ColumnHeaders(iColNo).Left, LV1.Top + LV1.SelectedItem.Top, LV1.ColumnHeaders(iColNo).Width
        .Text = LV1.SelectedItem.SubItems(iColNo - 1)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV1.ColumnHeaders.count
        If x > Val(LV1.ColumnHeaders(i).Left) And x < Val(LV1.ColumnHeaders(i).Left) + Val(LV1.ColumnHeaders(i).Width) Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub


Private Sub OptInspection_Click(Index As Integer)
    OptInspection(0).Tag = Index
End Sub

Private Sub OptTabs_Click(Index As Integer)

    Dim i As Integer
    For i = 0 To PicTabs.count - 1
        PicTabs(i).Visible = False
    Next
    PicTabs(Index).Visible = True
    
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        LV1.SelectedItem.SubItems(iColNo - 1) = txtEdit
        If LV1.SelectedItem.Index < LV1.ListItems.count Then
            LV1.ListItems(LV1.SelectedItem.Index + 1).Selected = True
            LV1_DblClick
            Exit Sub
        ElseIf iColNo < LV1.ColumnHeaders.count Then
            iColNo = iColNo + 1
            LV1.ListItems(1).Selected = True
            Call LV1_DblClick
            Exit Sub
        End If
        
        txtEdit.Visible = False
        'Call txtLotBarcode_KeyPress(13)
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
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


Public Function ShowMe(p_strRcvID As String) As Boolean

    On Error GoTo err
    strRcvID = p_strRcvID
        
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VVendRcvdDetailPO WHERE RcvID='" & strRcvID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtOrderNo = !PORefNo & ""
            
            txtVender = !AccTitle & ""
            txtDT = Format(!DT, "dd-MMM-yyyy")
            lGroupID = Val(!GroupID & "")
        End If
        .Close
        
        cmbMaterial.AddVals con, "'{' + RMID1 + '} ' + RMName", "VendRcvdDetailPO INNER JOIN VRMAndItemsUnioned ON VendRcvdDetailPO.MaterialID=VRMAndItemsUnioned.RMID1", "EntryID", " WHERE RcvID='" & strRcvID & "'"
        
        
'        .Open "SELECT * FROM ProcessInspection WHERE IPO_RefID=" & Val(txtLotBarcode) & " AND ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
'        If .EOF = False Then
'            LV1.Tag = Val(!EntryID & "")
'            txtDisposation = !Disposation & ""
'            txtRemarks = !Comments & ""
'            DTInspection = !DT
'            OptInspection(Abs(Not !LotStatus)).Value = True
'            '------------------------------------
'        Else
'            LV1.Tag = "0"
'            txtDisposation = ""
'            txtRemarks = ""
'            DTInspection = Date
'        End If
'        .Close
'
'        If Val(LV1.Tag) > 0 Then
'            .Open "SELECT * FROM ProcessInspectionDetail WHERE PI_RefID=" & Val(LV1.Tag), con, adOpenForwardOnly, adLockReadOnly
'            Do Until .EOF
'                Set ITM = LV1.ListItems(!PIP_RefID & "'")
'                ITM.SubItems(1) = !AQL & ""
'                ITM.SubItems(2) = !SampleSize & ""
'                ITM.SubItems(3) = !RejectOn & ""
'                ITM.SubItems(4) = !TestSpecificationNo & ""
'                ITM.SubItems(5) = !ActualRejection & ""
'                ITM.SubItems(6) = !Status & ""
'                .MoveNEXT
'            Loop
'            .Close
'        End If
    End With
    Set rs = Nothing
        
    Me.Show 1
    Exit Function
err:
    MsgBox err.Description
End Function
