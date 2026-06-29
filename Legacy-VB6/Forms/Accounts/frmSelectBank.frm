VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelectBank 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Double Click / Enter to Select Esc to Cancel"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmSelectBank.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1935
      Top             =   1275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectBank.frx":000C
            Key             =   "BNK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectBank.frx":0460
            Key             =   "ACC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectBank.frx":08B2
            Key             =   "ACCOD"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectBank.frx":0D04
            Key             =   "BCH"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tv 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   5636
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "Select This Bank Account"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "Set As Default Bank Account"
      End
      Begin VB.Menu mnuOD 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "frmSelectBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public BankAcc As String, BankID As Double

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call Tv_DblClick
Else
  Call Tv_DblClick
End If

End Sub


Private Sub RefreshTV()
Screen.MousePointer = vbHourglass
Dim DefaultIndex As Integer


Set TV.ImageList = ImageList1
TV.Nodes.Clear

Dim rsBnk As New ADODB.Recordset
Dim rsBch As New ADODB.Recordset
Dim rsAcc As New ADODB.Recordset

rsBnk.Open "Select Distinct Bank From VBankAccounts Order By Bank", con, adOpenForwardOnly

For i = 0 To rsBnk.RecordCount - 1
   TV.Nodes.add , , rsBnk![Bank] & "", rsBnk![Bank] & "", "BNK"
      
      If rsBch.State = 1 Then rsBch.Close
      rsBch.Open "Select Distinct BankID,Branch From VBankAccounts Where Bank='" & rsBnk![Bank] & "'", con, adOpenForwardOnly
      
      For k = 0 To rsBch.RecordCount - 1
         TV.Nodes.add rsBnk![Bank] & "", tvwChild, rsBch![BankID] & "BCH" & k, rsBch![Branch] & "", "BCH"
                  
                  
               If rsAcc.State = 1 Then rsAcc.Close
                rsAcc.Open "Select * From VBankAccounts Where BankID=" & rsBch![BankID], con, adOpenForwardOnly
                  For l = 0 To rsAcc.RecordCount - 1
                     TV.Nodes.add(rsBch![BankID] & "BCH" & k, tvwChild, rsAcc![AccNo], "[" & rsAcc![AccNo] & "] " & rsAcc![AccTitle], IIf(rsAcc![AllowOD], "ACCOD", "ACC")).Tag = Val(rsBch![BankID])
                     
                     If rsAcc![DefaultBank] Then DefaultIndex = TV.Nodes.count
                      
                     rsAcc.MoveNext
                  Next
                  
         rsBch.MoveNext
      Next
      
      

   rsBnk.MoveNext
Next

If DefaultIndex > 0 Then
   TV.Nodes(DefaultIndex).Bold = True
   TV.Nodes(DefaultIndex).Selected = True
   
End If

Set rsBnk = Nothing
Set rsCty = Nothing
Set rsAcc = Nothing
Screen.MousePointer = vbDefault
End Sub








Private Sub Form_Load()
 Call RefreshTV
 
End Sub

Private Sub mnuDefault_Click()
 con.Execute "Update BankAccounts Set [DefaultBank]=0"
 con.Execute "Update BankAccounts Set [DefaultBank]=1 Where AccNo='" & TV.SelectedItem.key & "'"
 Call RefreshTV
End Sub

Private Sub mnuOD_Click()
con.Execute "Update BankAccounts Set AllowOD=" & CBool(mnuOD.Tag <> "OD") & " Where Accno='" & TV.SelectedItem.key & "'"
Call RefreshTV
End Sub

Private Sub mnuSelect_Click()
Call Tv_DblClick
End Sub

Private Sub Tv_DblClick()

    If TV.Nodes.count = 0 Then GoTo Hide
    If Not TV.SelectedItem.Image Like "ACC*" Then Exit Sub
    With TV.SelectedItem
        BankAcc = .Text
        BankID = Val(.Tag)
    End With
 
Hide:
    Me.Hide
End Sub

Private Sub Tv_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        Call Tv_DblClick
    Else
        BankAcc = ""
        BankID = 0
        Me.Hide
    End If
 
End Sub

Private Sub Tv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Or TV.Nodes.count = 0 Then Exit Sub
    If TV.SelectedItem.Image = "ACC" Then
      mnuOD.Caption = "Allow Over Draft For This Account"
      Me.PopUpMenu mnuPop
   ElseIf TV.SelectedItem.Image = "ACCOD" Then
      mnuOD.Caption = "Disallow Over Draft For This Account"
      Me.PopUpMenu mnuPop
   End If
End Sub
