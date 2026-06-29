VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCurrencyExchRates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exchange Rates"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
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
   ScaleHeight     =   5160
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2700
      TabIndex        =   4
      Top             =   435
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   465
      Left            =   3840
      TabIndex        =   3
      Top             =   4620
      Width           =   1350
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3840
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   6773
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Currency"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   330
      Left            =   885
      TabIndex        =   1
      Top             =   360
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   52559875
      CurrentDate     =   40124
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   390
      Width           =   645
   End
End
Attribute VB_Name = "frmCurrencyExchRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bFormLoad As Boolean
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DT_Change()

    Dim rs As New ADODB.Recordset
    Dim i As Integer, iIndex As Integer
    With rs
        .Open "SELECT * FROM CurrencyExchangeRates WHERE DT='" & DT.Value & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            For i = 1 To LV.ListItems.count
                LV.ListItems(i).SubItems(2) = GetSingleDoubleValue("ExchRate", "VCurrencyExchangeRates", " WHERE Currency='" & LV.ListItems(i).ListSubItems(1) & "'")
                LV.ListItems(i).Tag = "0"
            Next
        End If
        Do Until .EOF
            iIndex = 0
            For i = 1 To LV.ListItems.count
                With LV.ListItems(i)
                    If .SubItems(1) = rs!Currency & "" Then
                        iIndex = i
                        Exit For
                    End If
                End With
            Next
            If iIndex <> 0 Then
                LV.ListItems(iIndex).SubItems(2) = Val(!ExchRate & "")
                LV.ListItems(iIndex).Tag = !EntryID
            Else
                LV.ListItems(iIndex).SubItems(2) = 0
                LV.ListItems(iIndex).Tag = ""
            End If
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub Form_Load()

    bFormLoad = True
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT DISTINCT Curr FROM ForeignCustomers", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.SubItems(1) = !Curr & ""
            ITM.SubItems(2) = ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    DT = Date
    Call DT_Change
   
    bFormLoad = False
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(3).Left, LV.Top + LV.SelectedItem.Top
        .Text = Val(LV.SelectedItem.SubItems(2))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(LV.SelectedItem.Tag) = 0 Then
            con.Execute "INSERT INTO CurrencyExchangeRates(DT,Currency,ExchRate) VALUES('" & DT & "','" & _
             LV.SelectedItem.SubItems(1) & "'," & Val(txtEdit) & ")"
             LV.SelectedItem.Tag = GetSingleLongValue("MAX(EntryID)", "CurrencyExchangeRates")
        Else
            con.Execute "UPDATE CurrencyExchangeRates SET ExchRate=" & Val(txtEdit) & " WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        End If
        LV.SelectedItem.SubItems(2) = Val(txtEdit)
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
