VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmViewPictureFSExt 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3930
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.FlatScrollBar VScroll1 
      Height          =   3390
      Left            =   5925
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   5980
      _Version        =   393216
      Orientation     =   1638400
   End
   Begin MSComCtl2.FlatScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3675
      Visible         =   0   'False
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   1935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewPictureFSExt.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   1482
      ButtonWidth     =   1693
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actual Size"
            Key             =   "ActualSize"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fit Screen"
            Key             =   "FitScreen"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   0
      Top             =   870
      Width           =   195
   End
End
Attribute VB_Name = "frmViewPictureFSExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lScreenNetHeight As Long
Dim lScreenNetWidth As Long

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
   m_cScroll_Scroll eBar
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)

    Dim lVal As Long
    If (eBar = efsHorizontal) Then
        lVal = 0 - m_cScroll.Value(eBar)
        lVal = lVal * 50
        Image1.Left = lVal
    Else
        lVal = (Toolbar1.Height / 50) - m_cScroll.Value(eBar)
        lVal = lVal * 50
        Image1.Top = lVal
   End If
   
End Sub

Public Sub ShowFromDBField(strFieldName As String, strTableName As String, strCondition As String)
    
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT " & strFieldName & " FROM " & strTableName & " " & strCondition, con, adOpenForwardOnly, adLockReadOnly
    
    Set Image1.DataSource = rs
    Image1.DataField = strFieldName
    
    rs.Close
    Set rs = Nothing
    Call FitOnScreen
    Call ResizeControls
    Me.Show 1
    
End Sub

Private Sub ResizeControls()

    Image1.Stretch = False
    Image1.Top = Toolbar1.Height
    If (Image1.Width + VScroll1.Width) > Screen.Width Then
        Me.Width = Screen.Width
        m_cScroll.Visible(efsHorizontal) = True
    Else
        Me.Width = Image1.Width + VScroll1.Width
    End If
    
    If Image1.Height + HScroll1.Height + Toolbar1.Height > Screen.Height Then
        Me.Height = Screen.Height
        m_cScroll.Visible(efsVertical) = True
    Else
        Me.Height = Image1.Height + Toolbar1.Height + HScroll1.Height
    End If
    
    lScreenNetHeight = Me.Height - 255 - Toolbar1.Height
    lScreenNetWidth = Me.Width - 240

    m_cScroll.SmallChange(efsVertical) = 15
    m_cScroll.LargeChange(efsVertical) = 100
    m_cScroll.Min(efsVertical) = 0
    m_cScroll.Max(efsVertical) = (Image1.Height - lScreenNetHeight) / 50
    
    m_cScroll.SmallChange(efsHorizontal) = 15
    m_cScroll.LargeChange(efsHorizontal) = 100
    m_cScroll.Min(efsHorizontal) = 0
    m_cScroll.Max(efsHorizontal) = (Image1.Width - lScreenNetWidth) / 50
    
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Public Sub ShowFromPic(myPic As StdPicture)
    
    Set Image1.Picture = myPic
    Image1.Top = Toolbar1.Height
    
    If (Image1.Width + VScroll1.Width) > Screen.Width Then
        Me.Width = Screen.Width
        m_cScroll.Visible(efsHorizontal) = True
    Else
        Me.Width = Image1.Width + VScroll1.Width
    End If
    
    If Image1.Height + HScroll1.Height + Toolbar1.Height > Screen.Height Then
        Me.Height = Screen.Height
        m_cScroll.Visible(efsVertical) = True
    Else
        Me.Height = Image1.Height + Toolbar1.Height + HScroll1.Height
    End If
    
    lScreenNetHeight = Me.Height - 255 - Toolbar1.Height
    lScreenNetWidth = Me.Width - 240

    m_cScroll.SmallChange(efsVertical) = 15
    m_cScroll.LargeChange(efsVertical) = 100
    m_cScroll.Min(efsVertical) = 0
    m_cScroll.Max(efsVertical) = (Image1.Height - lScreenNetHeight) / 50
    
    m_cScroll.SmallChange(efsHorizontal) = 15
    m_cScroll.LargeChange(efsHorizontal) = 100
    m_cScroll.Min(efsHorizontal) = 0
    m_cScroll.Max(efsHorizontal) = (Image1.Width - lScreenNetWidth) / 50
    Me.Show 1
    
End Sub

Private Sub Form_Load()
    ' Set up scroll bars:
    Set m_cScroll = New cScrollBars
    m_cScroll.Create Me.hwnd
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "Close"
            Unload Me
        Case "ActualSize"
            ResizeControls
        Case "FitScreen"
            Call FitOnScreen
    End Select
End Sub

Private Sub FitOnScreen()
    Image1.Width = 800 * 15
    Image1.Height = 600 * 15
    Image1.Stretch = True
    m_cScroll.Visible(efsHorizontal) = False
    m_cScroll.Visible(efsVertical) = False
    Me.Width = Image1.Width
    Me.Height = Image1.Height + Toolbar1.Height
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
