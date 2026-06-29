VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGetCustomersAndDateDouble 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Data"
   ClientHeight    =   7785
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetCustomersAndDateDouble.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnSelectAll 
      Caption         =   "Uncheck All"
      Height          =   375
      Left            =   1455
      TabIndex        =   4
      Top             =   7230
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Check All"
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   7230
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6450
      Left            =   135
      TabIndex        =   2
      Top             =   690
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   11377
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Country"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4575
      TabIndex        =   1
      Top             =   7230
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3330
      TabIndex        =   0
      Top             =   7230
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Top             =   315
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   84934659
      CurrentDate     =   39372
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2655
      TabIndex        =   6
      Top             =   315
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   84934659
      CurrentDate     =   39372
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&From Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   75
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&To Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   90
      Width           =   660
   End
End
Attribute VB_Name = "frmGetCustomersAndDateDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim bChecked As Boolean
Dim lBalance As Long
Dim strCustomer As String
Dim myDTFrom As Date, myDTTo As Date

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    LV.ListItems(1).Checked = False
    For i = 2 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdUnSelectAll_Click()
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        LV.ListItems(1).Checked = False
    Next
End Sub

Private Sub Form_Load()

    
    bChecked = False
      
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT CustCode,Country FROM ForeignCustomers ORDER BY CustCode,Country", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Set ITM = LV.ListItems.add(, , "<All Customers>")
        ITM.Tag = "0"
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !CustCode & "")
            ITM.SubItems(1) = !Country & ""
            ITM.Tag = !CustCode & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    LV.ListItems(1).Checked = True
    DTPicker1.Value = DateAdd("m", -1, Date)
    DTPicker2.Value = Date
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    If Item.Tag = "0" Then
        For i = 2 To LV.ListItems.count
            LV.ListItems(i).Checked = False
        Next
    Else
        LV.ListItems(1).Checked = False
    End If
    
End Sub

Private Sub OKButton_Click()
    
    Dim i As Integer, bNothing As Boolean
    bNothing = True
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            bNothing = False
            Exit For
        End If
    Next
    
    If bNothing Then
        MsgBox "Select Customer.", vbInformation
        LV.SetFocus
        Exit Sub
    End If
    
    bCancelled = False
    
    strCustomer = ""
    If LV.ListItems(1).Checked Then
        strCustomer = "All"
    Else
        For i = 2 To LV.ListItems.count
            With LV.ListItems(i)
                If .Checked Then
                    strCustomer = strCustomer & "'" & .Text & "-" & .SubItems(1) & "',"
                End If
            End With
        Next
        If strCustomer <> "" Then
            strCustomer = Left(strCustomer, Len(strCustomer) - 1)
        End If
    End If
    
    myDTFrom = DTPicker1
    myDTTo = DTPicker2
    
    Unload Me
    
End Sub

Public Function getData(ByRef p_strCustomer As String, ByRef p_myDTFrom As Date, ByRef p_myDTTo As Date) As Boolean
     
    Me.Show 1
    
    p_strCustomer = strCustomer
    p_myDTFrom = myDTFrom
    p_myDTTo = myDTTo
    
    getData = Not bCancelled
    
End Function
