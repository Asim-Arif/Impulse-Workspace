VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPorts 
   Appearance      =   0  'Flat
   Caption         =   "Ports"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   825
      Left            =   75
      ScaleHeight     =   765
      ScaleWidth      =   2205
      TabIndex        =   0
      Top             =   7245
      Visible         =   0   'False
      Width           =   2265
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   75
         TabIndex        =   2
         Top             =   375
         Width           =   1020
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   1125
         TabIndex        =   1
         Top             =   375
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackColor       =   &H00CFF0F3&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Currency 1   Currency 2  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   2100
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3585
      Top             =   4095
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":0000
            Key             =   "japan"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":0452
            Key             =   "brazil"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":08A4
            Key             =   "canada"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":0CF6
            Key             =   "denmark"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":1148
            Key             =   "italy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":159A
            Key             =   "australia"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":19EC
            Key             =   "spain"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":1E3E
            Key             =   "usa"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":2290
            Key             =   "pakistan"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":2762
            Key             =   "Country"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":2BB4
            Key             =   "City"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts.frx":3006
            Key             =   "Port"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TVCountries 
      Height          =   7095
      Left            =   15
      TabIndex        =   4
      Top             =   45
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   12515
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   2640
      TabIndex        =   5
      Top             =   7245
      Width           =   1665
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2937;635"
      Picture         =   "frmPorts.frx":3458
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmPorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddCountries()

    Dim rsCount As New ADODB.Recordset
    TVCountries.Nodes.Clear
    With rsCount
        .CursorLocation = adUseClient
        .Open "Select CountryName from Countries Where CountryName<>'Pak'", con, adOpenForwardOnly, adLockReadOnly
        On Error Resume Next
        Do Until .EOF
            TVCountries.Nodes.Add , , "|" & ![CountryName], ![CountryName], LCase(![CountryName])
            If err.Number = 35601 Then
                TVCountries.Nodes.Add , , "|" & ![CountryName], ![CountryName], "Country"
                err.Clear
            End If
            Call AddCities(![CountryName])
            .MoveNext
        Loop
        .Close
    End With
    Set rsCount = Nothing
    
End Sub
Private Sub AddCities(CountryName As String)

    Dim rsCity As New ADODB.Recordset
    With rsCity
        
        .Open "Select City,CityID from Cities Where CityID<> 4 and CountryName='" & CountryName & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            TVCountries.Nodes.Add "|" & CountryName, tvwChild, "||" & ![CityID], ![City] & "", "City" & ""
            Call AddPorts(CountryName, ![City] & "", ![CityID])
            .MoveNext
        Loop
        .Close
    End With
    Set rsCity = Nothing
    
End Sub
Private Sub AddPorts(CountryName As String, CityName As String, CityID As Integer)

    Dim rsPorts As New ADODB.Recordset
    With rsPorts
        .CursorLocation = adUseClient
        .Open "Select * from Ports Where PortID<>4 and CityID=" & CityID, con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            TVCountries.Nodes.Add "||" & CityID, tvwChild, ![PortID] & "'", ![Port] & "", "Port"
            .MoveNext
        Loop
        .Close
    End With
    Set rsPorts = Nothing
    
End Sub



Public Sub del_Click()
    Dim msg As String
    Dim TableName  As String
    With TVCountries.SelectedItem
        If Left$(.Key, 2) = "||" Then
            msg = "City"
            TableName = "Cities"
        ElseIf Left$(.Key, 1) = "|" Then
            msg = "CountryName"
            TableName = "Countries"
        Else
            msg = "Port"
            TableName = "Ports"
        End If
        If MsgBox("Do You Realy Want To Delete This " & msg & "?", vbQuestion + vbYesNo) = vbYes Then
            con.Execute "Delete from  " & TableName & " Where " & msg & "='" & .Text & "'"
            Call AddCountries
        End If
    End With
End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set TVCountries.ImageList = ImageList1
    Call AddCountries
End Sub

Public Sub NewCity_Click()

    On Error GoTo err
    
        
    NewVal = InputBox("Enter City Name :", "New City")
    If NewVal <> "" Then
        If Left$(TVCountries.SelectedItem.Key, 2) = "||" Then 'Means City Is Clicked
            con.Execute "Insert into Cities (CountryName,City) Values('" & TVCountries.SelectedItem.Parent & "','" & NewVal & "')"
        Else
            con.Execute "Insert into Cities (CountryName,City) Values('" & TVCountries.SelectedItem.Text & "','" & NewVal & "')"
        End If
    End If
    
    Call AddCountries
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub NewPort_Click()

    On Error GoTo err
    
    NewVal = InputBox("Enter Port Name :", "New Port")
    
    If NewVal <> "" Then
        If Left$(TVCountries.SelectedItem.Key, 2) = "||" Then 'Means City Is Clicked
            con.Execute "Insert into Ports (CityID,Port) Values(" & Mid(TVCountries.SelectedItem.Key, 3) & ",'" & NewVal & "')"
        Else
            con.Execute "Insert into Ports (CityID,Port) Values(" & Mid(TVCountries.SelectedItem.Parent.Key, 3) & ",'" & NewVal & "')"
        End If
    End If
    
    Call AddCountries
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub NewEntry_Click()

    On Error GoTo err

    Dim NewVal As String
    If TVCountries.SelectedItem Is Nothing Then
        NewVal = InputBox("Enter Country Name :", "New Country")
        If NewVal <> "" Then
            con.Execute "Insert into Countries(CountryName) Values('" & NewVal & "')"
        End If
    ElseIf Left$(TVCountries.SelectedItem.Key, 1) = "|" Then 'Means Country Is Clicked
        'frmNewCountry.Show 1, Me
        NewVal = InputBox("Enter Country Name :", "New Country")
        If NewVal <> "" Then
            con.Execute "Insert into Countries(CountryName) Values('" & NewVal & "')"
        End If
        'TVCountries.Nodes.add , , "|", NewVal, "Country"
    ElseIf Left$(TVCountries.SelectedItem.Key, 2) = "||" Then 'Means City Is Clicked
        NewVal = InputBox("Enter City Name :", "New City")
        If NewVal <> "" Then
            con.Execute "Insert into Cities Values('" & TVCountries.SelectedItem.Parent & "','" & NewVal & "')"
            
        End If
    Else
        NewVal = InputBox("Enter Port Name :", "New Port")
        If NewVal <> "" Then
            'con.Execute "Insert into Ports Values('" & TVCountries.SelectedItem.Parent.Parent & "','" & TVCountries.SelectedItem.Parent & "','" & NewVal & "')"
            con.Execute "Insert into Ports (CityID,Port) Values(" & Right(TVCountries.SelectedItem.Parent.Key, Len(TVCountries.SelectedItem.Parent.Key) - 2) & ",'" & NewVal & "')"
        End If
    End If
    
    Call AddCountries
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub Rename_Click()
    With TVCountries
        '.LabelEdit = tvwManual
        .StartLabelEdit
    End With
End Sub



Private Sub Form_Resize()
    TVCountries.Width = Me.Width - 150
    TVCountries.Height = Me.Height - 600 - cmdClose.Height - TVCountries.Top
    cmdClose.Move TVCountries.Width - cmdClose.Width, TVCountries.Top + TVCountries.Height + 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub Text1_Validate(Cancel As Boolean)

    'On Error Resume Next
    Dim UpdateQry As String
    With TVCountries.SelectedItem
        If Left$(.Key, 2) = "||" Then
            UpdateQry = "update Countries set Curr1='" & Text1 & "' where Country='" & .Parent.Text & "'"
        ElseIf Left$(.Key, 1) = "|" Then
            UpdateQry = "update Countries set Curr1='" & Text1 & "' where Country='" & .Text & "'"
        Else
            UpdateQry = "update Countries set Curr1='" & Text1 & "' where Country='" & .Parent.Parent.Text & "'"
        End If
        con.Execute UpdateQry
    End With
    
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    On Error Resume Next
    Dim UpdateQry As String
    With TVCountries.SelectedItem
        If Left$(.Key, 2) = "||" Then
            UpdateQry = "update Countries set Curr2='" & Text2 & "' where Country='" & .Parent.Text & "'"
        ElseIf Left$(.Key, 1) = "|" Then
            UpdateQry = "update Countries set Curr2='" & Text2 & "' where Country='" & .Text & "'"
        Else
            UpdateQry = "update Countries set Curr2='" & Text2 & "' where Country='" & .Parent.Parent.Text & "'"
        End If
    End With
    con.Execute UpdateQry
End Sub

Private Sub TVCountries_AfterLabelEdit(Cancel As Integer, NewString As String)

On Error GoTo err
    Dim UpdateQuery As String
    With TVCountries
        If Left$(.SelectedItem.Key, 2) = "||" Then 'Means City Is Edited
            UpdateQuery = "Update Cities set City='" & NewString & "' Where City='" & TVCountries.Tag & "'"
        ElseIf Left$(.SelectedItem.Key, 1) = "|" Then 'Means Country Is Edited
            UpdateQuery = "Update Countries set CountryName='" & NewString & "' Where CountryName='" & TVCountries.Tag & "'"
        Else
            UpdateQuery = "Update Ports set Port='" & NewString & "' Where Port='" & TVCountries.Tag & "'"
        End If
    End With
    con.Execute UpdateQuery
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub TVCountries_BeforeLabelEdit(Cancel As Integer)
    TVCountries.Tag = TVCountries.SelectedItem.Text
End Sub

Private Sub TVCountries_Click()
    Exit Sub
    Dim rsCurr As New ADODB.Recordset
    Dim CountryName As String
    On Error Resume Next
    With TVCountries
        If Left$(.SelectedItem.Key, 2) = "||" Then
            CountryName = .SelectedItem.Parent
        ElseIf Left$(.SelectedItem.Key, 1) = "|" Then
            CountryName = .SelectedItem.Text
        Else
            CountryName = .SelectedItem.Parent.Parent.Text
        End If
    End With
    
    'With rsCurr
    '    .CursorLocation = adUseClient
    '    .Open "Select Curr1,Curr2 from Countries Where Country='" & CountryName & "'", con, adOpenForwardOnly, adLockReadOnly
    '    If .RecordCount > 0 Then
    '        Text1.Text = ![Curr1] & ""
    '        Text2.Text = ![Curr2] & ""
    '    End If
    '    .Close
    'End With
    Set rsCurr = Nothing
    Exit Sub

End Sub

Private Sub TVCountries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo err
    
    'Dim strRet As String
    If Button = 2 Then
        With TVCountries
            
            If .SelectedItem Is Nothing Then
                strRet = ShowPopUpMenu1("&Add New Country")
            ElseIf Left$(.SelectedItem.Key, 2) = "||" Then 'Means City Is Clicked
                'MainForm.NewCity.Visible = True
                'MainForm.NewEntry.Caption = "Add New City"
                'MainForm.NewCity.Caption = "Add New Port"
                'MainForm.NewEntry.Tag = .SelectedItem.Key
                
                strRet = ShowPopUpMenu1("Add New &City", "-", "Add New &Port", "&Rename This City", "-", "&Delete")
    
    
            ElseIf Left$(.SelectedItem.Key, 1) = "|" Then 'Means Country Is Clicked
                'MainForm.NewCity.Visible = True
                'MainForm.NewCity.Caption = "Add New City"
                'MainForm.NewEntry.Caption = "Add New Country"
                'MainForm.NewEntry.Tag = ""
                
                strRet = ShowPopUpMenu1("&Add New Country", "-", "Add New &City", "&Rename This Country", "-", "&Delete")
                
            Else
                'MainForm.NewCity.Visible = False
                'MainForm.NewEntry.Caption = "Add New Port"
                'MainForm.NewEntry.Tag = .SelectedItem.Key
                
                strRet = ShowPopUpMenu1("Add New &Port", "-", "&Rename This Port", "-", "&Delete")
            End If
            'MainForm.Rename.Caption = "Rename " & .SelectedItem.Text
        End With
        
        'Me.PopUpMenu MainForm.mnupopPorts
        
            If strRet = "&Add New Country" Then
                Call NewEntry_Click
            ElseIf strRet = "Add New &City" Then
                Call NewCity_Click
            ElseIf strRet = "Add New &Port" Then
                Call NewPort_Click
            ElseIf Left(strRet, 7) = "&Rename" Then
                Call Rename_Click
            ElseIf strRet = "&Delete" Then
                Call del_Click
            End If
            
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub


