VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmChangeFixedAssetCatagory 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangeFixedAssetCatagory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   6555
      Index           =   0
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   8430
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "CHANGE HEAD"
      Top             =   0
      Width           =   8430
      Begin ComboList.Usercontrol1 cmbMaterial 
         Height          =   285
         Left            =   3015
         TabIndex        =   4
         Top             =   2130
         Width           =   5100
         _ExtentX        =   8996
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
      Begin ComboList.Usercontrol1 cmbNewCatagory 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   4245
         Width           =   7710
         _ExtentX        =   13600
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
      Begin ComboList.Usercontrol1 cmbCatagory 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   2130
         Width           =   2865
         _ExtentX        =   5054
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Original Details :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   1050
         Left            =   135
         TabIndex        =   9
         Top             =   2685
         Width           =   8010
         Begin VB.TextBox txtHead 
            Height          =   315
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   510
            Width           =   7680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Head"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00B54646&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   10
            Top             =   285
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Details :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   1050
         Left            =   120
         TabIndex        =   5
         Top             =   3735
         Width           =   8010
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Head"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00B54646&
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   8
            Top             =   285
            Width           =   825
         End
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   5625
         Picture         =   "frmChangeFixedAssetCatagory.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5310
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   6915
         Picture         =   "frmChangeFixedAssetCatagory.frx":09A9
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5310
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Head"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B54646&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   6
         Top             =   1920
         Width           =   435
      End
      Begin VB.Image Image2 
         Height          =   1200
         Left            =   150
         Picture         =   "frmChangeFixedAssetCatagory.frx":1208
         Top             =   645
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fixed Asset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B54646&
         Height          =   195
         Index           =   11
         Left            =   3000
         TabIndex        =   3
         Top             =   1920
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmChangeFixedAssetCatagory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    For i = 0 To PicWC.count - 1
        Call DrawCaptions(PicWC(i))
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox)

    Dim lBorderColor As Long, lFillColor
    lBorderColor = RGB(70, 70, 181)
    lFillColor = RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 350)-(Pic.Width, 350), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 600), lFillColor, BF
        Pic.ForeColor = vbWhite
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 180
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 100)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 30, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub

Private Sub cmbCatagory_matched()

    Dim strCondition As String
    If cmbCatagory.id <> "0" Then
        strCondition = " WHERE MasterFA_RefID=" & cmbCatagory.id
    End If
    
    cmbMaterial.ClearVals
    cmbMaterial.AddVals con, "'{' + AssetID + '} ' +  Description", "FixedAssets", "AssetID", strCondition & " ORDER BY Description"
    
End Sub

Private Sub cmbMaterial_matched()

    On Error GoTo err
    
    Dim rs  As New ADODB.Recordset
    
    With rs
        .Open "SELECT EntryID,FixedAssets_MasterDescription FROM VFixedAssets WHERE AssetID='" & cmbMaterial.id & "'", con, adOpenForwardOnly, adLockReadOnly
        txtHead = !FixedAssets_MasterDescription & ""
        txtHead.Tag = Val(!EntryID & "")
        .Close
    End With
    Set rs = Nothing
    
    cmbNewCatagory.ClearVals
    cmbNewCatagory.AddVals con, "Description", "FixedAssets_Master", "EntryID", " ORDER BY Description"
    'cmbNewCatagory.AddVals con, "'{' + CatID + '} ' + Catagory", "MaterialCatagories", "CatID", " WHERE CatID<>'" & txtCatID & "'"
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbNewCatagory_matched()

    Dim strNewAccNo As String
    Dim strAccTitle As String
    'strAccTitle = GetSingleStringValue("AccTitle", "Accounts", " WHERE AccNo='" & cmbMaterial.id & "'")
    
    'strNewAccNo = GetNextAccno(strAccTitle, Left(cmbMaterial.id, 2), cmbNewCatagory.id, False)
   
    
    'txtNewMaterialID = GetNextMaterialID(cmbNewCatagory.ID)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    If cmbMaterial.MatchFound = False Then
        MsgBox "Please Select Account.", vbInformation
        Exit Sub
    ElseIf cmbNewCatagory.MatchFound = False Then
        MsgBox "Please Select New Head.", vbInformation
        Exit Sub
    ElseIf txtHead = cmbNewCatagory.id Then
        MsgBox "Old & New Heads are same.", vbInformation
        Exit Sub
    End If
    
    Call StartTrans(con)
    
    con.Execute "INSERT INTO FixedAssetsHeadChanges(AssetID,FA_Master_RefID,FA_Master_New_RefID,UserName,MachineName) " & _
     "VALUES('" & cmbMaterial.id & "'," & txtHead.Tag & "," & cmbNewCatagory.id & ",'" & CurrentUserName & "','" & _
     strComputerName & "')"
    
    con.Execute "UPDATE FixedAssets SET MasterFA_RefID=" & cmbNewCatagory.id & " WHERE AssetID='" & cmbMaterial.id & "'"
    
     
    con.CommitTrans
    bSaved = True
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Form_Load()

    cmbCatagory.ListHeight = 3200
    cmbMaterial.ListHeight = 3200
    cmbNewCatagory.ListHeight = 1600
    
    cmbCatagory.AddItem "<All Heads>", "0"
    cmbCatagory.AddVals con, "Description", "FixedAssets_Master", "EntryID", " ORDER BY Description"
    cmbCatagory.id = "0"
    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = x
                lDownY = y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
End Sub

Public Function ShowMe(Optional strMaterialID As String) As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show

    If strMaterialID <> "" Then
        cmbMaterial.id = strMaterialID
    End If
    Me.Show 1
    ShowMe = bSaved
    
End Function


Private Function GetNextMaterialID(CatID As String) As String
    Dim NextVal As String
    NextVal = GetMax("Cast(Right(MaterialID,3) as int)", "Material", "Where CatID ='" & CatID & "'")
    GetNextMaterialID = CatID & "-" & Format(Val(NextVal) + 1, "000")
End Function

