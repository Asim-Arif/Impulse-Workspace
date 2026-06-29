VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmReminders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Reminders"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmReminders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReminders.frx":000C
            Key             =   "Salaries"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReminders.frx":0326
            Key             =   "AdvSals"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReminders.frx":0480
            Key             =   "Bonus"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2970
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5239
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "frmReminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    If LV.ListItems.Count = 0 Then Unload Me
End Sub

Private Sub Form_Load()
  
  Dim Count As Long
  Set LV.SmallIcons = ImageList2

  Count = con.Execute("select count(distinct EmpID) From MonthlySalaries Where Month(dt)=" & Month(ServerDate) & " and year(dt)=" & year(ServerDate) & " and Payed=0 and AdvSalary=1").Fields(0)
  If Count > 0 Then LV.ListItems.Add , "AdvSals", "Advance Salary For " & Count & " Employee(s) is Payable...", , "AdvSals"
  
  Count = con.Execute("select count(distinct EmpID) From MonthlySalaries Where Month(dt)=" & Month(ServerDate) - 1 & " and year(dt)=" & year(ServerDate) & " and Payed=0 and AdvSalary=0").Fields(0)
  If Count > 0 Then LV.ListItems.Add , "Salaries", "Salary For " & Count & "  Employee(s) is Payable...", , "Salaries"
  
  Count = con.Execute("select count(*) From Employees Where Month(JoinDate)=" & Month(ServerDate) - 1 & " and year(JoinDate)<=" & year(ServerDate) & " and Bonus=1 and Active=1").Fields(0)
  If Count > 0 Then LV.ListItems.Add , "Bonus", "Bonus For " & Count & "  Employee(s) is Payable...", , "Bonus"
  
  
  

End Sub



Private Sub LV_DblClick()

    Dim Item As MSComctlLib.ListItem
    If LV.ListItems.Count = 0 Then Exit Sub
    Set Item = LV.SelectedItem
    Me.Hide
    Select Case Item.key
       Case "AdvSals"
          'If MainForm.mnuAdvSalary.Visible Then Call MainForm.mnuAdvSalary_Click
       Case "Salaries"
          'If MainForm.mnuSalSheet.Visible Then Call MainForm.mnuSalSheet_Click
       Case "Bonus"
          'If MainForm.mnuBonusSheet.Visible Then Call MainForm.mnuBonusSheet_Click
    End Select
    Unload Me
    
End Sub



