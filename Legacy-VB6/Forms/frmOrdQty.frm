VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOrdQty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "For Specific Orders..."
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "frmOrdQty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4305
   Begin MSComctlLib.ListView LV 
      Height          =   4950
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   8731
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Order No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Qty"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Wastage"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmOrdQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotQty As Long, Wastage As Long
Public Sub FillList(Col As Collection, ItemCode As String)
    
    Dim ITM As ListItem
    For i = 1 To Col.Count
        If Col(i).ItemCode = ItemCode Then
            Set ITM = LV.ListItems.Add(, , Col(i).OrderNo)
            ITM.Tag = Col(i).lOrdQty
            ITM.SubItems(1) = Col(i).lQty
            ITM.SubItems(2) = Col(i).lWastage
            TotQty = TotQty + Col(i).lQty
            Wastage = Wastage + Col(i).lWastage
        End If
    Next i
    
    Me.Show 1
    
End Sub
