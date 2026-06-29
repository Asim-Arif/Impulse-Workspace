VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAddBookmarkforReceiving 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bookmarking"
   ClientHeight    =   7500
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
   ScaleHeight     =   7500
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Bookmarking :"
      Height          =   4890
      Left            =   105
      TabIndex        =   15
      Top             =   1980
      Width           =   11445
      Begin VB.TextBox txtOrderQty 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   10395
         TabIndex        =   35
         Top             =   3255
         Width           =   810
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
         Height          =   555
         Left            =   8190
         TabIndex        =   29
         Top             =   3945
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
         Height          =   555
         Left            =   9720
         TabIndex        =   28
         Top             =   3945
         Width           =   1455
      End
      Begin VB.TextBox txtEntryDT 
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
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4110
         Width           =   2610
      End
      Begin VB.TextBox txtMachineName 
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
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4110
         Width           =   2610
      End
      Begin VB.TextBox txtUserName 
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
         TabIndex        =   23
         Top             =   4110
         Width           =   2610
      End
      Begin VB.TextBox txtRefTo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   5715
         TabIndex        =   21
         Top             =   2550
         Width           =   5475
      End
      Begin VB.TextBox txtInstructions 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   180
         TabIndex        =   19
         Top             =   2550
         Width           =   5505
      End
      Begin VB.TextBox txtActions 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   1725
         Left            =   180
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   495
         Width           =   11010
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
         Index           =   15
         Left            =   10395
         TabIndex        =   34
         Top             =   3030
         Width           =   300
      End
      Begin MSForms.ComboBox cmbItemCode 
         Height          =   390
         Left            =   2040
         TabIndex        =   33
         Top             =   3255
         Width           =   8355
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "14737;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbOrderNo 
         Height          =   390
         Left            =   180
         TabIndex        =   32
         Top             =   3255
         Width           =   1860
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3281;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
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
         Index           =   14
         Left            =   2040
         TabIndex        =   31
         Top             =   3030
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Order No."
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
         TabIndex        =   30
         Top             =   3030
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entry Date"
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
         Left            =   5475
         TabIndex        =   25
         Top             =   3915
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Computer Name"
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
         Left            =   2820
         TabIndex        =   24
         Top             =   3915
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
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
         Left            =   180
         TabIndex        =   22
         Top             =   3900
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ref. To"
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
         Left            =   5745
         TabIndex        =   20
         Top             =   2325
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Instructions By"
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
         Left            =   180
         TabIndex        =   18
         Top             =   2325
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Perform Following Action on above lot"
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
         TabIndex        =   16
         Top             =   270
         Width           =   3195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Issuance Details :"
      Height          =   1785
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   11445
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
         Left            =   9210
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1185
         Width           =   2025
      End
      Begin VB.TextBox txtOrderNo 
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
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1185
         Width           =   2445
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
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1200
         Width           =   2445
      End
      Begin VB.TextBox txtProcess 
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
         TabIndex        =   8
         Top             =   1200
         Width           =   4065
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
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   6975
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
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2025
      End
      Begin VB.TextBox txtLotNo 
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
         Top             =   480
         Width           =   2025
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
         Left            =   9210
         TabIndex        =   13
         Top             =   975
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Order No."
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
         Left            =   6720
         TabIndex        =   11
         Top             =   975
         Width           =   780
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
         Left            =   4290
         TabIndex        =   9
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Process"
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
         Left            =   210
         TabIndex        =   7
         Top             =   975
         Width           =   660
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
         Left            =   4275
         TabIndex        =   5
         Top             =   255
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
         Left            =   2250
         TabIndex        =   3
         Top             =   255
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lot No."
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
         Left            =   210
         TabIndex        =   1
         Top             =   255
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmAddBookmarkforReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lVID_RefID As Long
Dim bCancelled As Boolean
Dim bViewOnly As Boolean
Dim DTDisplayStart As Date, DTDisplayFinish As Date
Dim lVID_Bookmark_EntryID As Long
Dim bUpdateBookmark As Boolean

Public Function AddBookmark(p_lVID_RefID As Long) As Boolean
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(DetailEntryID)", "VVendIssued", " WHERE DetailEntryID=" & p_lVID_RefID & " AND TotalIssQty>ISNULL(RcvdQty,0)")
    
    If lCount = 0 Then  'Already Received..
        MsgBox "Already Received.", vbInformation
        Unload Me
        Exit Function
    End If
    
    bViewOnly = False
    
    lVID_RefID = p_lVID_RefID
    
    Call ViewLotDetails(lVID_RefID)
    
    txtUserName = CurrentUserName
    txtMachineName = strComputerName
    txtEntryDT = GetServerDate(True)
   
    Me.Show 1
    AddBookmark = Not bCancelled
    
End Function

Private Sub ViewLotDetails(lEntryID As Long)
    
    bViewOnly = True
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ItemCode,ItemName,LotNo,RecieptID,IssQty,OrderNo,CustCode,Description FROM VVendIssdDetail WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            If !LotNo & "" = "0" Then
                txtLotNo = !RecieptID & ""
            Else
                txtLotNo = !LotNo & ""
            End If
            txtItemCode = !ItemCode & ""
            txtItemName = !ItemName & ""
            txtProcess = !Description & ""
            txtQty = Val(!IssQty & "")
            txtOrderNo = !OrderNo & ""
            txtCustomer = !CustCode & ""
        End If
        .Close
        
        If bViewOnly Then   'Show Bookmark's Details
            .Open "SELECT * FROM VID_Bookmarks WHERE VID_RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                txtActions = !Actions & ""
                txtInstructions = !InstructionsBy & ""
                txtRefTo = !RefTo & ""
                txtUserName = !UserName & ""
                txtMachineName = !MachineName & ""
                txtEntryDT = Format(!DTEntry, "dd-MMM-yyyy HH:nn")
                lVID_Bookmark_EntryID = Val(!EntryID & "")
                
                cmbOrderNo.Text = !OrderNo & ""
                Dim i As Integer
                For i = 0 To cmbItemCode.ListCount - 1
                    If cmbItemCode.List(i, 1) = !ITemID & "" Then
                        cmbItemCode.ListIndex = i
                        Exit For
                    End If
                Next
                txtOrderQty = Val(!Qty & "")
            End If
            .Close
        End If
        
    End With
    Set rs = Nothing
    
    
     
End Sub

Public Function ViewBookmark(p_lVID_RefID As Long, Optional p_bUpdateBookmark As Boolean = False) As Boolean

    cmdSave.Visible = False
    cmdCancel.Caption = "OK"
    '--------------------------------------------
    txtActions.Locked = True
    txtInstructions.Locked = True
    txtRefTo.Locked = True
    '--------------------------------------------
    bViewOnly = True
    '--------------------------------------------
    
    'Closing here If no Bookmarks Found.
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(VID_RefID)", "VID_Bookmarks", " WHERE VID_RefID=" & p_lVID_RefID)
    
    If lCount = 0 Then
        Unload Me
        Exit Function
    End If
     
    bUpdateBookmark = p_bUpdateBookmark
    
    ViewLotDetails (p_lVID_RefID)
    
    DTDisplayStart = GetServerDate(True)
    
    Me.Show 1
    
End Function

Private Sub cmbOrderNo_Change()
    Call cmbOrderNo_Click
End Sub

Private Sub cmbOrderNo_Click()

    If cmbOrderNo.MatchFound = False Then
        cmbItemCode.Clear
        Exit Sub
    Else
        Call AddToCombo(cmbItemCode, "'('+CompItemID+') ' + ItemName", "VrptOrders_GroupedItems", " WHERE OrderNo='" & cmbOrderNo.Text & "'", , "CompItemID")
    End If
    
End Sub

Private Sub cmdCancel_Click()

    If bViewOnly = False Then
    
        bCancelled = True
        Unload Me
        
    Else
        
        If bUpdateBookmark Then
        
            DTDisplayFinish = GetServerDate(True)
        
            con.Execute "UPDATE VID_Bookmarks SET ViewDone=1,DTViewStart='" & DTDisplayStart & "',DTViewFinish='" & _
             DTDisplayFinish & "' WHERE EntryID=" & lVID_Bookmark_EntryID
             
        End If
        
        bCancelled = True
        Unload Me
        
    End If
    
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
    
    If txtActions = "" Then
        MsgBox "Please Specify Some Actions.", vbInformation
        Exit Sub
    End If
    
    If cmbOrderNo.MatchFound = False Then
        MsgBox "Please Select Order No.", vbInformation
        Exit Sub
    ElseIf cmbItemCode.MatchFound = False Then
        MsgBox "Please Select Item Code.", vbInformation
        Exit Sub
    ElseIf Val(txtOrderQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    'Save...
    If lVID_Bookmark_EntryID = 0 Then
        con.Execute "INSERT INTO VID_Bookmarks(VID_RefID,Actions,InstructionsBy,RefTo,UserName,MachineName,OrderNo,ItemID,Qty)" & _
         " VALUES(" & lVID_RefID & ",'" & txtActions & "','" & txtInstructions & "','" & txtRefTo & "','" & _
         CurrentUserName & "','" & strComputerName & "','" & cmbOrderNo & "','" & cmbItemCode.List(cmbItemCode.ListIndex, 1) & "'," & Val(txtOrderQty) & ")"
    Else
        con.Execute "UPDATE VID_Bookmarks SET Actions='" & txtActions & "',InstructionsBy='" & txtInstructions & "',RefTo='" & txtRefTo & _
         "',OrderNo='" & cmbOrderNo & "',ItemID='" & cmbItemCode.List(cmbItemCode.ListIndex, 1) & "',Qty=" & Val(txtOrderQty) & " WHERE EntryID=" & lVID_Bookmark_EntryID
    End If
    bCancelled = False
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    bCancelled = True
    Call AddToCombo(cmbOrderNo, "OrderNo", "FCustomerOrders", " WHERE CustCode<>'Stock'")
    
End Sub
