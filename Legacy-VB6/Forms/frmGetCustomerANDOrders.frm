VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmGetCustomerANDOrders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Customer & Orders"
   ClientHeight    =   7215
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetCustomerANDOrders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   270
      Width           =   2145
      _ExtentX        =   3784
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
   Begin VB.CommandButton cmdUnSelectAll 
      Caption         =   "Uncheck All"
      Height          =   375
      Left            =   1455
      TabIndex        =   4
      Top             =   6705
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Check All"
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   6705
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6030
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10636
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Order No"
         Object.Width           =   4233
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4575
      TabIndex        =   1
      Top             =   6705
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3330
      TabIndex        =   0
      Top             =   6705
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Country"
      Height          =   195
      Index           =   1
      Left            =   2325
      TabIndex        =   8
      Top             =   75
      Width           =   585
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   2310
      TabIndex        =   6
      Top             =   270
      Width           =   2220
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3916;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   75
      Width           =   690
   End
End
Attribute VB_Name = "frmGetCustomerANDOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim bChecked As Boolean
Dim lBalance As Long
Dim strOrders As String, strCustCode As String, strCountry As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT OrderNo FROM FCustomerOrders WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList)", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'Set ITM = LV.ListItems.add(, , "< All >")
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !OrderNo & "")
            ITM.Tag = !OrderNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    'LV.ListItems(1).Checked = True
    
End Sub

Private Sub cmbCust_matched()
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select Distinct Country From ForeignCustomers "
    If cmbCust.ID <> "0" Then
        SQL = SQL & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            'cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    cmbCountry.ListIndex = 0
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    LV.ListItems(1).Checked = False
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdUnSelectAll_Click()
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
End Sub

Private Sub Form_Load()

    
    bChecked = False
      
    cmbCust.ListHeight = 2500
    'cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
    
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    If Item.Tag = "0" Then
        For i = 1 To LV.ListItems.count
            LV.ListItems(i).Checked = False
        Next
    Else
        'LV.ListItems(1).Checked = False
    End If
    
End Sub

Private Sub OKButton_Click()
    
    If cmbCust.MatchFound = False Then
        MsgBox "Please Select Customer.", vbInformation
        Exit Sub
    End If
    
    If cmbCountry.MatchFound = False Then
        MsgBox "Please Select Country.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer, bNothing As Boolean
    bNothing = True
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            bNothing = False
            Exit For
        End If
    Next
    
    If bNothing Then
        MsgBox "Please Select Order(s).", vbInformation
        LV.SetFocus
        Exit Sub
    End If
    
    bCancelled = False
    
    strOrders = ""
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strOrders = strOrders & "'" & .Tag & "',"
            End If
        End With
    Next
    If strOrders <> "" Then
        strOrders = Left(strOrders, Len(strOrders) - 1)
    End If
    
    strCustCode = cmbCust.ID
    strCountry = cmbCountry.Text
    
    Unload Me
    
End Sub

Public Function getData(ByRef p_strCustCode As String, ByRef p_strCountry As String, ByRef p_strOrders As String) As Boolean
     
    Me.Show 1
    
    p_strCustCode = strCustCode
    p_strCountry = strCountry
    
    p_strOrders = strOrders
    
    getData = Not bCancelled
    
End Function
