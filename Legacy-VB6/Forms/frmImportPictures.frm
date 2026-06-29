VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportPictures 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Pictures"
   ClientHeight    =   3612
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6204
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3612
   ScaleWidth      =   6204
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLeft 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4524
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3108
      Width           =   1404
   End
   Begin VB.TextBox txtFail 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3096
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3108
      Width           =   1404
   End
   Begin VB.TextBox txtSuccess 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1668
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3108
      Width           =   1404
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   432
      Left            =   216
      TabIndex        =   6
      Top             =   1992
      Width           =   5736
      _ExtentX        =   10118
      _ExtentY        =   762
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtTotalPictures 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3108
      Width           =   1404
   End
   Begin VB.TextBox txtPath 
      Height          =   348
      Left            =   216
      TabIndex        =   3
      Top             =   300
      Width           =   5736
   End
   Begin VB.CommandButton cmdStartImport 
      Caption         =   "Start Import"
      Height          =   396
      Left            =   4596
      TabIndex        =   1
      Top             =   756
      Width           =   1332
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1224
      Left            =   1296
      Pattern         =   "*.jpg"
      TabIndex        =   0
      Top             =   -720
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Left"
      Height          =   192
      Index           =   4
      Left            =   5100
      TabIndex        =   12
      Top             =   2892
      Width           =   252
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Failure"
      Height          =   192
      Index           =   3
      Left            =   3528
      TabIndex        =   10
      Top             =   2892
      Width           =   468
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Success"
      Height          =   192
      Index           =   2
      Left            =   2088
      TabIndex        =   8
      Top             =   2892
      Width           =   552
   End
   Begin VB.Label lblImportStatus 
      Alignment       =   2  'Center
      Caption         =   "Importing"
      Height          =   192
      Left            =   216
      TabIndex        =   7
      Top             =   2472
      Width           =   5736
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Pictures"
      Height          =   192
      Index           =   0
      Left            =   264
      TabIndex        =   4
      Top             =   2892
      Width           =   936
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path :"
      Height          =   192
      Left            =   216
      TabIndex        =   2
      Top             =   84
      Width           =   408
   End
End
Attribute VB_Name = "frmImportPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim rsPic As New ADODB.Stream

Private Sub cmdStartImport_Click()

    File1.Path = txtPath.Text
    File1.Refresh
    txtTotalPictures = File1.ListCount
    
    PB1.Min = 1
    PB1.Max = File1.ListCount
        
    Dim i As Long, lTotalSuccess As Long, lTotalFail As Long, iStatus As Integer
    For i = 0 To File1.ListCount - 1
        iStatus = UpdatePicture(File1.List(i), File1.Path & "\" & File1.List(i))
        If iStatus = 1 Then
            lTotalSuccess = lTotalSuccess + 1
        ElseIf iStatus = 2 Then
            lTotalFail = lTotalFail + 1
        End If
        lblImportStatus.Caption = "Importing " & (i + 1) & " of " & File1.ListCount
        PB1.Value = i + 1
        txtSuccess = lTotalSuccess
        txtFail = lTotalFail
        txtLeft = Val(txtTotalPictures) - lTotalSuccess - lTotalFail
        Me.Refresh
        DoEvents
    Next
    lblImportStatus.Caption = "Done"
    
End Sub

Private Function UpdatePicture(strFileName As String, strFilePath As String) As Integer
    
    
    If rs.State = 1 Then rs.Close
    If rsPic.State = 1 Then rsPic.Close
    Dim iStatus As Integer
    iStatus = 2
    rs.Open "SELECT ItemID,ItemPic FROM Items WHERE ItemID='" & Left(strFileName, Len(strFileName) - 4) & "'", con, adOpenStatic, adLockOptimistic
    
    If rs.EOF = False Then
        rsPic.Open
        
        rsPic.LoadFromFile strFilePath
        rs.Fields(1) = rsPic.Read
        rs.Update
        iStatus = 1
        rsPic.Close
    Else
        iStatus = 2
    End If
    
    rs.Close
    con.Execute "INSERT INTO PicturesImportStatus(ItemID,Status,Path) VALUES('" & strFileName & "'," & iStatus & ",'" & strFilePath & "')"
        
    UpdatePicture = iStatus
    
End Function

Private Sub Form_Load()
    rsPic.Type = adTypeBinary
End Sub
