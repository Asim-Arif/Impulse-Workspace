VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewRack 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rack Defination"
   ClientHeight    =   2400
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
   Icon            =   "frmNewRack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCont 
      Caption         =   "Make Bins by Continous Numbers"
      Height          =   210
      Left            =   105
      TabIndex        =   12
      Top             =   2055
      Width           =   2880
   End
   Begin VB.TextBox txtShelf 
      Height          =   315
      Left            =   105
      TabIndex        =   9
      Top             =   1620
      Width           =   2670
   End
   Begin VB.TextBox txtTotalBins 
      Height          =   315
      Left            =   2790
      TabIndex        =   8
      Top             =   1620
      Width           =   1755
   End
   Begin VB.TextBox txtTotalRacks 
      Height          =   315
      Left            =   3300
      TabIndex        =   6
      Top             =   1050
      Width           =   1245
   End
   Begin ComboList.Usercontrol1 cmbStore 
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   375
      Width           =   4515
      _ExtentX        =   7964
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
   End
   Begin VB.TextBox txtRack 
      Height          =   315
      Left            =   105
      TabIndex        =   3
      Top             =   1050
      Width           =   3165
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1530
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Bin Name"
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
      Index           =   4
      Left            =   135
      TabIndex        =   11
      Top             =   1410
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bins (Each Rack)"
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
      Index           =   2
      Left            =   2835
      TabIndex        =   10
      Top             =   1410
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Racks"
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
      Index           =   3
      Left            =   3300
      TabIndex        =   7
      Top             =   840
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Store Name"
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
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Rack Name"
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
      Left            =   135
      TabIndex        =   2
      Top             =   840
      Width           =   945
   End
End
Attribute VB_Name = "frmNewRack"
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

Private Sub Form_Load()
    cmbStore.ListHeight = 2400
    cmbStore.AddVals con, "StoreName", "Stores", "EntryID", " ORDER BY StoreName"
End Sub

Private Sub OKButton_Click()
     
    On Error GoTo err
    If cmbStore.MatchFound = False Then
        MsgBox "Please Select Store.", vbInformation
        Exit Sub
    ElseIf txtRack = "" Then
        MsgBox "Invalid Rack.", vbInformation
        Exit Sub
    End If
    
    Dim iBinNo As Integer
    iBinNo = 0
    If lEntryID = 0 Then
        Dim i As Integer, j As Integer, lRackEntryID As Long
        Call StartTrans(con)
        For i = 1 To Val(txtTotalRacks)
            con.Execute "INSERT INTO StoreRacks(Store_RefID,RackNo) VALUES(" & cmbStore.ID & ",'" & txtRack & "-" & i & "')"
            lRackEntryID = GetSingleLongValue("MAX(EntryID)", "StoreRacks")
            
            For j = 1 To Val(txtTotalBins)
                iBinNo = iBinNo + 1
                con.Execute "INSERT INTO StoreShelfs(Rack_RefID,ShelfNo) VALUES(" & lRackEntryID & ",'" & txtShelf & "-" & IIf(chkCont.Value = vbChecked, iBinNo, j) & "')"
            Next
            
        Next
        con.CommitTrans
    Else
        con.Execute "UPDATE StoreRacks SET RackNo='" & txtRack & "' WHERE EntryID=" & lEntryID
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
            .Open "SELECT * FROM StoreRacks WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            cmbStore.ID = !Store_RefID
            txtRack = !RackNo & ""
            .Close
        End With
        Set rs = Nothing
    End If
    
    Me.Show 1

    ShowMe = Not bCancelled
    
End Function
