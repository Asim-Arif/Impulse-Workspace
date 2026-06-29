VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMiscGateInAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gate In...."
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveNew 
      Caption         =   "Save && &New"
      Height          =   465
      Left            =   2265
      TabIndex        =   21
      Top             =   6150
      Width           =   1635
   End
   Begin VB.TextBox txtDescription 
      Height          =   405
      Left            =   3330
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtVehicleNo 
      Height          =   405
      Left            =   270
      TabIndex        =   3
      Top             =   1080
      Width           =   3045
   End
   Begin VB.TextBox txtName 
      Height          =   405
      Left            =   3330
      TabIndex        =   2
      Top             =   315
      Width           =   3975
   End
   Begin VB.TextBox txtPacking 
      Height          =   405
      Left            =   5805
      TabIndex        =   8
      Top             =   1785
      Width           =   1485
   End
   Begin VB.TextBox txtUnit 
      Height          =   405
      Left            =   3720
      TabIndex        =   6
      Top             =   1785
      Width           =   1020
   End
   Begin VB.TextBox txtDetail 
      Height          =   405
      Left            =   270
      TabIndex        =   5
      Top             =   1785
      Width           =   3435
   End
   Begin VB.TextBox txtGatePassNo 
      Height          =   405
      Left            =   2130
      TabIndex        =   1
      Top             =   315
      Width           =   1185
   End
   Begin VB.TextBox txtQty 
      Height          =   405
      Left            =   4755
      TabIndex        =   7
      Top             =   1785
      Width           =   1020
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   405
      Left            =   270
      TabIndex        =   0
      Top             =   315
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   714
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy HH:mm"
      Format          =   115212291
      CurrentDate     =   41041
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   465
      Left            =   3960
      TabIndex        =   10
      Top             =   6150
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   5655
      TabIndex        =   11
      Top             =   6150
      Width           =   1635
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3840
      Left            =   270
      TabIndex        =   9
      Top             =   2235
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   6773
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Detail"
         Object.Width           =   6137
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unit"
         Object.Width           =   1773
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Packing"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   3315
      TabIndex        =   20
      Top             =   855
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vehicle No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   300
      TabIndex        =   19
      Top             =   855
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   3375
      TabIndex        =   18
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Packing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   5805
      TabIndex        =   17
      Top             =   1575
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   16
      Top             =   1575
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   270
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   4755
      TabIndex        =   14
      Top             =   1575
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gate Pass #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2130
      TabIndex        =   13
      Top             =   90
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   12
      Top             =   90
      Width           =   405
   End
End
Attribute VB_Name = "frmMiscGateInAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lEntryID As Long
Dim bReturnValue As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Saved Then
        Unload Me
    End If
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    
    If txtGatePassNo = "" Then
        MsgBox "Please Enter Gate Pass No.", vbInformation
        Exit Function
    End If
    
    If LV.ListItems.count = 0 Then
        MsgBox "No Items Added.", vbInformation
        Exit Function
    End If
     
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        con.Execute "INSERT INTO MiscEntries_GateIn(DT,GatePassNo,PersonName,VehicleNo," & _
         "Description,UserName,MachineName) VALUES('" & DTPicker1.Value & "','" & txtGatePassNo & "','" & _
         txtName & "','" & txtVehicleNo & _
         "','" & txtDescription & "','" & CurrentUserName & "','" & strComputerName & "')"
         
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "MiscEntries_GateIn")
    Else
        con.Execute "UPDATE MiscEntries_GateIn SET DT='" & DTPicker1 & "',GatePassNo='" & txtGatePassNo & _
         "',PersonName='" & txtName & "',VehicleNo='" & txtVehicleNo & "',Description='" & txtDescription & "' WHERE EntryID=" & lEntryID
        lmyEntryID = lEntryID
        
        con.Execute "DELETE FROM MiscEntries_GateIn_Details WHERE MEGI_RefID=" & lEntryID
    End If
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO MiscEntries_GateIn_Details(MEGI_RefID,Detail,Unit,Qty,Packing) VALUES(" & _
             lmyEntryID & ",'" & .Text & "','" & .SubItems(1) & "'," & Val(.SubItems(2)) & ",'" & .SubItems(3) & "')"
        End With
    Next
    
    bReturnValue = True
    
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub cmdSaveNew_Click()
    If Saved Then
        lEntryID = 0
        Unload Me
        Me.Show 1
    End If
End Sub

Private Sub Form_Load()

    lEntryID = 0
    DTPicker1.Value = Now
    bReturnValue = False
    
End Sub

Public Function ShowMe(Optional p_lEntryID As Long) As Boolean

    lEntryID = p_lEntryID
    If lEntryID > 0 Then
        Dim rs As New ADODB.Recordset, ITM As ListItem
        With rs
            .Open "SELECT * FROM MiscEntries_GateIn WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            DTPicker1 = !DT & ""
            txtGatePassNo = !GatePassNo & ""
            txtName = !PersonName & ""
            txtVehicleNo = !VehicleNo & ""
            txtDescription = !Description & ""
            .Close
            
            .Open "SELECT * FROM MiscEntries_GateIn_Details WHERE MEGI_RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            LV.ListItems.Clear
            Do Until .EOF
                Set ITM = LV.ListItems.add(, , !Detail & "")
                ITM.ListSubItems.add , , !Unit & ""
                ITM.ListSubItems.add , , Val(!Qty & "")
                ITM.ListSubItems.add , , !Packing & ""
                .MoveNext
            Loop
            .Close
        End With
    End If
    Me.Show 1
    ShowMe = bReturnValue
    
End Function

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub txtPacking_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then
    
        If txtDetail = "" Then
            MsgBox "Please Enter Detail.", vbInformation
            Exit Sub
        ElseIf Val(txtQty) <= 0 Then
            MsgBox "Please Enter Qty.", vbInformation
            Exit Sub
        End If
        
        Dim ITM As ListItem
        Set ITM = LV.ListItems.add(, , txtDetail)
        ITM.ListSubItems.add , , txtUnit
        ITM.ListSubItems.add , , Val(txtQty)
        ITM.ListSubItems.add , , txtPacking
        
    End If
    
End Sub
