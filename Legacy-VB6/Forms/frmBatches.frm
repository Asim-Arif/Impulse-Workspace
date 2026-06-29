VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBatches 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Making new batchs & assigning them "
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   225
      TabIndex        =   0
      Top             =   75
      Width           =   6825
      Begin VB.TextBox txtBatchNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   9
         Top             =   405
         Width           =   1995
      End
      Begin VB.CommandButton cmdAssign 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   7
         Top             =   2115
         Width           =   885
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   4320
         Left            =   3015
         TabIndex        =   6
         Top             =   735
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   7620
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Batch No"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox cmbCustomers 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   405
         Width           =   2025
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "Compute"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   2
         Top             =   720
         Width           =   2025
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   60
         TabIndex        =   1
         Top             =   1110
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   3015
         TabIndex        =   12
         Top             =   405
         Width           =   1485
         _ExtentX        =   2619
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   60358659
         CurrentDate     =   37250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Customers"
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
         Height          =   195
         Left            =   4500
         TabIndex        =   11
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Batch No(if not from date)"
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
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   195
         Width           =   2205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date :"
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
         Height          =   195
         Left            =   3045
         TabIndex        =   8
         Top             =   195
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   225
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   6675
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmBatches.frx":0000
         Height          =   4680
         Left            =   60
         TabIndex        =   5
         Top             =   150
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   8255
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "batchno"
            Caption         =   "Batch #"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "orderno"
            Caption         =   "Order #"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "customer"
            Caption         =   "Customer"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmBatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NewOrderNo As String

Private Sub cmbCustomers_Change()
    Call RefreshList
End Sub

Private Sub cmdAssign_Click()
    
    On Error GoTo err
    If List1.SelCount = 0 Then MsgBox "No batch is selected": Exit Sub
    
    With LV1
        If .SelectedItem.Selected = False Then
            MsgBox "No order is selected "
            Exit Sub
        End If
    
        If (Len(LV1.SelectedItem.SubItems(1)) > 1) Then
            MsgBox "Order has been assigned a batch No."
            Exit Sub
        End If
    End With
 
    con.Execute "Update FCustomerOrders Set BatchNo='" & List1.Text & "' where CustCode='" & cmbCustomers.Text & "' and orderno='" & LV1.SelectedItem.Text & "'"
    
 
    Call RefreshList
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCompute_Click()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    Dim BatchNo As String
    
    
    If txtBatchNo <> "" Then
        BatchNo = txtBatchNo
    Else
        BatchNo = Format(DTPicker1.value, "yymmdd") 'Mid(year(DTPicker1.value), 3, 2) & m & d
    End If
     
    With rs
        .Open "select BatchNo from Batches Where BatchNo='" & BatchNo & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            MsgBox "Batch No Is Present Against Selected Date."
            Exit Sub
        Else
            con.Execute "Insert Into Batches (BatchNo) values ('" & BatchNo & "')"
        End If
    End With
     
    Call RefreshList
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_Load()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    
    With rs
        .Open "Select Distinct(CustCode) from ForeignCustomers", con, adOpenForwardOnly, adLockReadOnly
        cmbCustomers.Clear
        Do Until .EOF
            cmbCustomers.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
    
   If cmbCustomers.ListCount > 0 Then cmbCustomers.ListIndex = 0
   
   Call RefreshList
   
   
    Exit Sub
err:
    MsgBox err.Description
   

End Sub
Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub
Private Sub RefreshList()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select BatchNo From Batches Where BatchNo Not In(Select BatchNo From FCustomerOrders Where BatchNo Is Not Null)", con, adOpenForwardOnly, adLockReadOnly
        List1.Clear
        Do Until .EOF
            List1.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
        
        .Open "Select OrderNo,BatchNo From FCustomerOrders Where CustCode='" & cmbCustomers.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        Do Until .EOF
            Set ITM = LV1.ListItems.Add(, , ![OrderNo] & "")
            ITM.SubItems(1) = ![BatchNo] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

    If NewOrderNo <> "" Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select BatchNo From CustOrders Where OrderNo='" & NewOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
            If IsNull(.Fields(0)) Then
                MsgBox "Please Assign Batch No. To The Order # : " & NewOrderNo, vbInformation
                cancel = 1
                Exit Sub
            End If
            .Close
        End With
        Set rs = Nothing
    End If
    
    
    NewOrderNo = ""
    
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If KeyCode = 46 Then
        If List1.ListCount > 0 Then
            If MsgBox("Do You Want To Delete This BatchNo ?", vbQuestion + vbYesNo, "Deleting") = vbYes Then
                con.Execute "Delete From Batches Where BatchNo='" & List1.Text & "'"
            End If
        End If
        Call RefreshList
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub
