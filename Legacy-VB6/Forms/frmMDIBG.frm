VERSION 5.00
Begin VB.Form frmMDIBG 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   7170
      Left            =   60
      Top             =   60
      Width           =   8325
   End
End
Attribute VB_Name = "frmMDIBG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Form_Resize()
    Image1.Move 0, 0
    Me.Width = Image1.Width
    Me.Height = Image1.Height
End Sub

