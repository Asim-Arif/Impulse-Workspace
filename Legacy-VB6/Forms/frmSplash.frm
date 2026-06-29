VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7125
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   7125
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Pbar1 
      Height          =   195
      Left            =   1860
      TabIndex        =   2
      Top             =   6900
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   2940
      TabIndex        =   3
      Top             =   3735
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00778E1C&
      Height          =   645
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2340
      Width           =   135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   645
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   2355
      Width           =   135
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim NotFirstTime As Boolean

Private Sub ShowReg()

'
'    DBName = GetSetting(App.Title, "Settings", "DatabasePath", App.Path & "\BloodBank.mdb")
'
'    Call ConnectToDatabase(DBName)
'    Load frmMain
'
'    Dim rec As New ADODB.Recordset
'
'    frmMain.Show
'    Label1.Caption = "Loading Data ..."
'    Me.Refresh

End Sub

Private Sub Form_Activate()

    Call FadeMe(Me, True)
    If NotFirstTime Then Exit Sub
    NotFirstTime = True
Me.Refresh
With Pbar1
If .Value > 0 Then Exit Sub
For i = .Min To .Max - 1 Step 3
 .Value = i
 If i = 18 Then
    'Label1.Caption = "Connecting To Database ..."
'    Call ShowReg
 End If
 
 
 For j = 0 To 30000
 Next
 

  If .Value Mod 15 = 0 Then Sleep (150)
Next
   
    
    
    Call FadeMe(Me)

    Me.Hide
    
    'MainForm.Logo.Picture = ""
    'CompanyWiseSettings
    CompanyWiseMainLinks
    
'    Load MainForm
'    Dim lTBBackColor As Long
'    lTBBackColor = vbWhite 'CreateSolidBrush(RGB(240, 120, 120))        'Creates the background from a Color (Long)
'    ChangeTBBack MainForm.MainToolbar, lTBBackColor, enuTB_STANDARD
    
    MainForm.Show
'    InvalidateRect 0&, 0&, False
    
''Check For Registration
'Dim rec As New ADODB.Recordset
'
    FrmLogin.Show 1
    Unload Me
End With
End Sub

Private Sub CompanyWiseMainLinks()
    
     
    Dim iIndex As Integer
'    If Not (strCompany = "Dr-Frgz") Then
'        iIndex = MainForm.MainToolbar.Buttons("Sampling").Index
'        MainForm.MainToolbar.Buttons("Sampling").Visible = False
'        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
'    End If
    
    If strCompany = "Kami" Then
        
        iIndex = MainForm.MainToolbar.Buttons("Export").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("Sampling").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("Financial").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("FixedAssets").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("QMS").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("Dashboard").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("Help").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
    End If
    
    If strCompany = "Banzai" Then
    
        iIndex = MainForm.MainToolbar.Buttons("Dashboard").Index
        MainForm.MainToolbar.Buttons("Dashboard").Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("QMS").Index
        MainForm.MainToolbar.Buttons("QMS").Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
    End If
    
    If Not (strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Towne") Then
        iIndex = MainForm.MainToolbar.Buttons("FixedAssets").Index
        MainForm.MainToolbar.Buttons("FixedAssets").Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
    End If
    
    If strCompany = "Instrumed" Then
    
        iIndex = MainForm.MainToolbar.Buttons("Payroll").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("Production").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("QMS").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("FixedAssets").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("Dashboard").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
        iIndex = MainForm.MainToolbar.Buttons("Help").Index
        MainForm.MainToolbar.Buttons(iIndex).Visible = False
        MainForm.MainToolbar.Buttons(iIndex + 1).Visible = False
        
    End If
        
    If strCompany = "Scard" Then
    
        For i = 1 To MainForm.MainToolbar.Buttons.count
            MainForm.MainToolbar.Buttons(i).Visible = False
        Next
        
        Dim iFinIndex As Integer
        iFinIndex = MainForm.MainToolbar.Buttons("Financial").Index
        MainForm.MainToolbar.Buttons(iFinIndex).Visible = True
        MainForm.MainToolbar.Buttons(iFinIndex + 1).Visible = True
        
        Dim iExitIndex As Integer
        iExitIndex = MainForm.MainToolbar.Buttons("Exit").Index
        MainForm.MainToolbar.Buttons(iExitIndex).Visible = True
        MainForm.MainToolbar.Buttons(iExitIndex + 1).Visible = True
        
        Dim iMinButtonIndex As Integer
        iMinButtonIndex = MainForm.MainToolbar.Buttons("PlaceHolder1").Index
        MainForm.MainToolbar.Buttons(iMinButtonIndex).Visible = True
        'MainForm.MainToolbar.Buttons(iMinButtonIndex + 1).Visible = True
    End If
    
End Sub

Private Sub Form_Load()
    
    'Registered to : K.M. Ashraf && Sons (Pvt.) Ltd.
 '   Call ApplyColorScheme(Me)
    
'    Dim rgn As Long
'    rgn = CreateRoundRectRgn(3, 3, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 20, 20)
'    Call SetWindowRgn(Me.hwnd, rgn, True)

 Call TransparentMe(frmSplash)
    
End Sub

