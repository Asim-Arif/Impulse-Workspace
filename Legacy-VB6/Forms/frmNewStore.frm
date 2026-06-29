VERSION 5.00
Begin VB.Form frmNewStore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Store Defination"
   ClientHeight    =   1140
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewStore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStore 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   375
      Width           =   4515
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Store  Name"
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
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   1035
   End
End
Attribute VB_Name = "frmNewStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim lEntryID As Long

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub OKButton_Click()
     
    On Error GoTo err
    If txtStore = "" Then
        MsgBox "Invalid Store.", vbInformation
        Exit Sub
    End If
    
    If lEntryID = 0 Then
        con.Execute "INSERT INTO Stores(StoreName) VALUES('" & txtStore & "')"
    Else
        con.Execute "UPDATE Stores SET StoreName='" & txtStore & "' WHERE EntryID=" & lEntryID
    End If
    bCancelled = False
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Function ShowMe(Optional p_lEntryID As Long = 0) As Boolean
    
    lEntryID = p_lEntryID
    
    If lEntryID <> 0 Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT * FROM Stores WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            txtStore = !StoreName & ""
            .Close
        End With
        Set rs = Nothing
    End If
    
    Me.Show 1

    ShowMe = Not bCancelled
    
End Function
