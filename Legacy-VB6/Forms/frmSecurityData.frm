VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSecurityData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecurityData.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEdit 
      Height          =   330
      Left            =   1695
      TabIndex        =   3
      Top             =   8595
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSComctlLib.ListView LVSecurity 
      Height          =   8460
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   14923
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Security"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   6555
      TabIndex        =   1
      Top             =   8640
      Width           =   1425
      Caption         =   " Close          "
      PicturePosition =   327683
      Size            =   "2514;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   375
      Left            =   5055
      TabIndex        =   0
      Top             =   8640
      Width           =   1425
      Caption         =   "OK           "
      PicturePosition =   327683
      Size            =   "2514;661"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmSecurityData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ShowSecurityData()

    Dim rs As New ADODB.Recordset, ITM As ListItem, strPassword As String
    With rs
        .Open "SELECT * FROM Security", con, adOpenForwardOnly, adLockReadOnly
        LVSecurity.ListItems.Clear
        Do Until .EOF
            Set ITM = LVSecurity.ListItems.add(, , LVSecurity.ListItems.count + 1)
            ITM.Tag = !EntryID
            ITM.ListSubItems.add , , !SettingName & ""
            strPassword = !Password & ""
            strPassword = CryptRC4(FromHexDump(strPassword), "AwmWrangler")
            
            ITM.ListSubItems.add , , strPassword
            'ITM.ListSubItems.add(, , "**********").Tag = !Password & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub Form_Load()

    On Error GoTo err
    
    
    
    Call ShowSecurityData
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LVSecurity_DblClick()

    With txtEdit
        .Move LVSecurity.Left + LVSecurity.ColumnHeaders(3).Left, LVSecurity.Top + LVSecurity.SelectedItem.Top, LVSecurity.ColumnHeaders(3).Width
        .Text = LVSecurity.SelectedItem.SubItems(2)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Dim strNewPassword As String
        strNewPassword = ToHexDump(CryptRC4(txtEdit, "AwmWrangler"))
        con.Execute "UPDATE Security SET Password='" & strNewPassword & "' WHERE EntryID=" & Val(LVSecurity.SelectedItem.Tag)
        LVSecurity.SelectedItem.SubItems(2) = txtEdit
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
