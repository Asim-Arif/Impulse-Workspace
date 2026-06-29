VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmAddFollowupData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Followup Data"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
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
   ScaleHeight     =   9315
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddNewFollowup 
      Caption         =   "New Follow up"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4650
      TabIndex        =   38
      Top             =   2220
      Width           =   2370
   End
   Begin VB.ComboBox cmbFollowup 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   105
      TabIndex        =   37
      Text            =   "Combo1"
      Top             =   2190
      Width           =   4455
   End
   Begin VB.CheckBox chks 
      Caption         =   "F III"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   2
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "(F3)"
      Top             =   2190
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.CheckBox chks 
      Caption         =   "F II"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   29
      Tag             =   "(F2)"
      Top             =   2190
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.CheckBox chks 
      Caption         =   "F I"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   0
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "(F1)"
      Top             =   2190
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Frame Frame2 
      Caption         =   "Followup :"
      Height          =   6315
      Left            =   105
      TabIndex        =   9
      Top             =   2970
      Width           =   11445
      Begin VB.TextBox txtManagerRemarks 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   165
         TabIndex        =   35
         Top             =   5115
         Width           =   11025
      End
      Begin ComboList.Usercontrol1 cmbEmp 
         Height          =   285
         Left            =   4170
         TabIndex        =   14
         Top             =   1170
         Width           =   6990
         _ExtentX        =   12330
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
      Begin VB.TextBox txtFollowUpOther 
         BackColor       =   &H00F1FEFE&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1170
         Width           =   2130
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3300
         TabIndex        =   22
         Top             =   1650
         Visible         =   0   'False
         Width           =   2000
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   390
         Left            =   165
         TabIndex        =   20
         Top             =   435
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   688
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
         Format          =   146210819
         CurrentDate     =   40894
      End
      Begin MSComctlLib.ListView LV 
         Height          =   3345
         Left            =   165
         TabIndex        =   19
         Top             =   1485
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Process"
            Object.Width           =   14111
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Qty"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8190
         TabIndex        =   16
         Top             =   5700
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   9720
         TabIndex        =   15
         Top             =   5700
         Width           =   1455
      End
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   1815
         TabIndex        =   11
         Top             =   435
         Width           =   9345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Director / Manager Remarks"
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
         Index           =   13
         Left            =   180
         TabIndex        =   36
         Top             =   4905
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Follow up by employee (select emp id)"
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
         Index           =   10
         Left            =   4185
         TabIndex        =   27
         Top             =   960
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Follow up by (if other)"
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
         Index           =   9
         Left            =   1995
         TabIndex        =   26
         Top             =   960
         Width           =   1845
      End
      Begin MSForms.ComboBox cmbFollowupBy 
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   1170
         Width           =   1815
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3201;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Follow up method"
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
         Left            =   180
         TabIndex        =   25
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   180
         TabIndex        =   21
         Top             =   225
         Width           =   405
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
         Index           =   8
         Left            =   1815
         TabIndex        =   10
         Top             =   225
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order Details :"
      Height          =   1605
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   11445
      Begin VB.TextBox txtMakerName 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
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
         Height          =   390
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   435
         Width           =   8130
      End
      Begin VB.TextBox txtMakerCode 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
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
         Height          =   390
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   435
         Width           =   2835
      End
      Begin VB.TextBox txtDeliveryDT 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
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
         Height          =   390
         Left            =   9465
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1110
         Width           =   1695
      End
      Begin VB.TextBox txtDT 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
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
         Height          =   390
         Left            =   7755
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1110
         Width           =   1695
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
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
         Height          =   390
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1110
         Width           =   870
      End
      Begin VB.TextBox txtCustomer 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
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
         Height          =   390
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1110
         Width           =   1365
      End
      Begin VB.TextBox txtItemName 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
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
         Height          =   390
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1110
         Width           =   4095
      End
      Begin VB.TextBox txtItemCode 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
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
         Height          =   390
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1110
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Maker Name"
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
         Index           =   12
         Left            =   3060
         TabIndex        =   34
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Maker Code"
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
         Index           =   11
         Left            =   210
         TabIndex        =   32
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Date"
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
         Left            =   9510
         TabIndex        =   24
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   7785
         TabIndex        =   18
         Top             =   900
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
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
         Left            =   6885
         TabIndex        =   7
         Top             =   900
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
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
         Left            =   5760
         TabIndex        =   5
         Top             =   885
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item Name"
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
         Left            =   1380
         TabIndex        =   3
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item Code"
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
         Left            =   210
         TabIndex        =   1
         Top             =   885
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmAddFollowupData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lVIDRDTS As Long, strFollowup As String, lEntryID As Long
Dim bCancelled As Boolean

Public Function AddFollowup(p_lVIDRDTS As Long, p_strFollowUP As String, Optional p_lEntryID As Long = 0) As Boolean
    
'    Dim lCount As Long
'    lCount = GetSingleLongValue("COUNT(DetailEntryID)", "VVendIssued", " WHERE DetailEntryID=" & p_lVID_RefID & " AND TotalIssQty>ISNULL(RcvdQty,0)")
'
'    If lCount = 0 Then  'Already Received..
'        MsgBox "Already Received.", vbInformation
'        Unload Me
'        Exit Function
'    End If
    lEntryID = p_lEntryID
    lVIDRDTS = p_lVIDRDTS
    
    strFollowup = p_strFollowUP
    
    cmbFollowupBy.AddItem "Phone"
    cmbFollowupBy.AddItem "Visit"
    cmbFollowupBy.AddItem "Other"
    
    cmbEmp.ClearVals
    cmbEmp.ListHeight = 2300
    cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name + ' {' + Designation + '} ' + '{' + DeptName + '}'", "VEmp", "EmpID"
        
    Call ViewDetails(lVIDRDTS)
    
    With cmbFollowup
        .AddItem "(F1)"
        .AddItem "(F2)"
        .AddItem "(F3)"
        .ListIndex = 0
    End With
    
    Call AddToCombo(cmbFollowup, "DISTINCT Followup", "FollowUpData", " WHERE VIDRDTS_RefID=" & lVIDRDTS & " AND Followup NOT IN('(F1)','(F2)','(F3)')", True)
    
'    If strFollowup = "(F1)" Then
'        chks(0).Value = vbChecked
'    ElseIf strFollowup = "(F2)" Then
'        chks(1).Value = vbChecked
'    ElseIf strFollowup = "(F3)" Then
'        chks(2).Value = vbChecked
'    End If
    
    Me.Show 1
    AddFollowup = Not bCancelled
    
End Function

Private Sub ViewDetails(p_lEntryID As Long)
    
    Dim ITM As ListItem
    Dim rs As New ADODB.Recordset, i As Integer
    With rs
        .Open "SELECT * FROM VPurchaseCalendar INNER JOIN Items ON VPurchaseCalendar.ItemCode=Items.ItemID WHERE EntryIDReturnDTs=" & p_lEntryID & " AND Prefix='" & strFollowup & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtMakerCode = !VendID1 & ""
            txtMakerName = !VenderName & ""
            txtItemCode = !ItemCode & ""
            txtItemName = !ItemName & ""
            txtQty = Val(!IssQty & "") - Val(!RcvQty & "")
            txtCustomer = !CustCode & ""
            txtDT = Format(!ReturnDT, "dd-MMM-yyyy")
            txtDeliveryDT = Format(!ActualReturnDT, "dd-MMM-yyyy")
        End If
        .Close
        
        .Open "SELECT * FROM VPurchaseCalendar WHERE EntryIDReturnDTs=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
'            If !Prefix & "" = "(F1)" Then
'                If IsNull(!ReturnDT) Then
'                    Chks(0).Enabled = False
'                ElseIf !ReturnDT > Date Then
'                    Chks(0).Enabled = False
'                End If
'            ElseIf !Prefix & "" = "(F2)" Then
'                If IsNull(!ReturnDT) Then
'                    Chks(1).Enabled = False
'                ElseIf !ReturnDT > Date Then
'                    Chks(1).Enabled = False
'                End If
'            ElseIf !Prefix & "" = "(F3)" Then
'                If IsNull(!ReturnDT) Then
'                    Chks(2).Enabled = False
'                ElseIf !ReturnDT > Date Then
'                    Chks(2).Enabled = False
'                End If
'            End If
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM SubProcesses WHERE Active=1 ORDER BY SubProcessID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !SubProcessName & "")
            ITM.Tag = Val(!SubProcessID & "")
            ITM.ListSubItems.add , , ""
            .MoveNEXT
        Loop
        .Close
         
    End With
    Set rs = Nothing
    
End Sub


Private Sub Chks_Click(Index As Integer)
    
    If chks(Index).Value = vbUnchecked Then Exit Sub
    strFollowup = chks(Index).Tag
    
    Dim i As Integer
    For i = 0 To 2
        If i <> Index Then
            chks(i).Value = vbUnchecked
        End If
    Next
    
    Call ViewFollowUpData
    
End Sub

Private Sub cmbFollowup_Change()
    cmbFollowup_Click
End Sub

Private Sub cmbFollowup_Click()

    strFollowup = cmbFollowup.Text
    
    Call ViewFollowUpData
    
End Sub

Private Sub cmbFollowupBy_Change()
    Call cmbFollowupBy_Click
End Sub

Private Sub cmbFollowupBy_Click()

    If cmbFollowupBy.ListIndex <> 2 Then
        txtFollowUpOther.Locked = True
        txtFollowUpOther.Text = ""
        txtFollowUpOther.BackColor = &HF1FEFE
    Else
        txtFollowUpOther.Locked = False
        txtFollowUpOther.Text = ""
        txtFollowUpOther.BackColor = vbWhite
    End If
    
End Sub

Private Sub cmdAddNewFollowup_Click()

    Dim strFollowup_New As String
    strFollowup_New = cmbFollowup.List(cmbFollowup.ListCount - 1)
    strFollowup_New = Mid$(strFollowup_New, 3, Len(strFollowup_New) - 3)
        
    strFollowup_New = "(F" & (Val(strFollowup_New) + 1) & ")"
    
    cmbFollowup.AddItem strFollowup_New
    
End Sub

Private Sub cmdCancel_Click()

    bCancelled = True
    Unload Me
     
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
    
    If txtRemarks = "" Then
        MsgBox "Please Add Remarks.", vbInformation
        Exit Sub
    End If
    
    If cmbFollowupBy.MatchFound = False Then
        MsgBox "Please Select Followup By.", vbInformation
        Exit Sub
    Else
        If cmbFollowupBy.ListIndex = 2 Then
            If txtFollowUpOther = "" Then
                MsgBox "Please Type Followup in Other.", vbInformation
                Exit Sub
            End If
        End If
    End If
    
    If cmbEmp.MatchFound = False Then
        MsgBox "Please Select Employee who did the followup.", vbInformation
        Exit Sub
    ElseIf cmbFollowup.ListIndex = -1 Then
        MsgBox "Please Select Followup.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer ', lEntryID As Long
    Dim strFollowUpOther As String
    If cmbFollowupBy.ListIndex = 2 Then
        strFollowUpOther = "'" & txtFollowUpOther & "'"
    Else
        strFollowUpOther = "NULL"
    End If
    Call StartTrans(con)
    'Save...
    
    If lEntryID = 0 Then
        con.Execute "INSERT INTO FollowUpData(VIDRDTS_RefID,FollowUp,DT,UserName,MachineName,Remarks,FollowUpBy,FollowUpOther,FollowUpPerson,ManagerRemarks)" & _
         " VALUES(" & lVIDRDTS & ",'" & strFollowup & "','" & DT & "','" & CurrentUserName & "','" & _
         strComputerName & "','" & txtRemarks & "'," & cmbFollowupBy.ListIndex & "," & strFollowUpOther & ",'" & cmbEmp.ID & "','" & txtManagerRemarks & "')"
        
        lEntryID = GetSingleLongValue("MAX(EntryID)", "FollowUpData")
    Else
        con.Execute "UPDATE FollowUpData SET Remarks='" & txtRemarks & "',FollowUpBy=" & cmbFollowupBy.ListIndex & _
         ",FollowUpOther=" & strFollowUpOther & ",FollowUpPerson='" & cmbEmp.ID & "',ManagerRemarks='" & txtManagerRemarks & "' WHERE EntryID=" & lEntryID
    End If
   
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.ListSubItems(1).Tag) > 0 Then
                con.Execute "UPDATE FollowUpData_Detail SET Qty=" & Val(.SubItems(1)) & " WHERE EntryID=" & Val(.ListSubItems(1).Tag)
            Else
                con.Execute "INSERT INTO FollowUpData_Detail(FollowUpData_RefID,SubProcessID,Qty) VALUES(" & _
                 lEntryID & "," & Val(.Tag) & "," & Val(.SubItems(1)) & ")"
            End If
        End With
    Next
    
    con.CommitTrans
    bCancelled = False
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    bCancelled = True
    DT.Value = Date
        
    
    
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(2).Left, LV.Top + LV.SelectedItem.Top
        .Text = Val(LV.SelectedItem.SubItems(1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(1) = Val(txtEdit)
        If LV.SelectedItem.Index < LV.ListItems.count Then
            LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
            Call LV_DblClick
        Else
            txtEdit.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub ViewFollowUpData()
    
    'Clear the Previous Data...
    txtRemarks = ""
    cmbEmp.ID = ""
    cmbFollowupBy.ListIndex = -1
    txtFollowUpOther = ""
    Dim ITM As ListItem
    Dim rs As New ADODB.Recordset, i As Integer, lRefID As Long
    With rs
        
        '.Open "SELECT * FROM FollowUpData WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM FollowUpData WHERE VIDRDTS_RefID=" & lVIDRDTS & " AND FollowUp='" & strFollowup & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            DT.Value = !DT
            txtRemarks = !Remarks & ""
            cmbEmp.ID = !FollowUpPerson & ""
            cmbFollowupBy.ListIndex = Val(!FollowUpBy & "")
            txtFollowUpOther = !FollowUpOther & ""
            txtManagerRemarks = !ManagerRemarks & ""
            lRefID = Val(!EntryID & "")
        End If
        .Close
        
        
        .Open "SELECT * FROM FollowUpData_Detail WHERE FollowUpData_RefID=" & lRefID
        Do Until .EOF
            For i = 1 To LV.ListItems.count
                If Val(LV.ListItems(i).Tag) = Val(!SubProcessID & "") Then
                    LV.ListItems(i).SubItems(1) = Val(!Qty & "")
                    LV.ListItems(i).ListSubItems(1).Tag = Val(!EntryID & "")
                    Exit For
                End If
            Next
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rs = Nothing
    
End Sub

