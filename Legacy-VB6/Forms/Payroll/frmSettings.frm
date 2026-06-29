VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings..."
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
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
   ScaleHeight     =   7515
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   2100
      TabIndex        =   0
      Top             =   210
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   5505
      Left            =   60
      TabIndex        =   4
      Top             =   615
      Width           =   7425
      Begin VB.Frame Frame4 
         Caption         =   "Timings"
         Height          =   5685
         Left            =   -15
         TabIndex        =   58
         Top             =   0
         Width           =   7440
         Begin VB.Frame Frame6 
            Caption         =   "Payroll Catagory"
            Height          =   3645
            Left            =   75
            TabIndex        =   59
            Top             =   1845
            Width           =   7335
            Begin VB.ComboBox cmbCatagory 
               Height          =   315
               Left            =   1110
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   270
               Width           =   1335
            End
            Begin MSComctlLib.ListView LVCat 
               Height          =   2955
               Left            =   105
               TabIndex        =   61
               Top             =   600
               Width           =   7140
               _ExtentX        =   12594
               _ExtentY        =   5212
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   15199215
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Setting"
                  Object.Width           =   4233
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Value"
                  Object.Width           =   4233
               EndProperty
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Catagory # :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   72
               Left            =   120
               TabIndex        =   62
               Top             =   315
               Width           =   945
            End
         End
         Begin MSComCtl2.DTPicker TExit 
            Height          =   285
            Left            =   4950
            TabIndex        =   63
            Top             =   450
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145752067
            UpDown          =   -1  'True
            CurrentDate     =   37958.7083333333
         End
         Begin MSComCtl2.DTPicker TEnter 
            Height          =   285
            Left            =   1305
            TabIndex        =   64
            Top             =   450
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145752067
            UpDown          =   -1  'True
            CurrentDate     =   37958.3333333333
         End
         Begin MSComCtl2.DTPicker LunchStart 
            Height          =   285
            Left            =   2520
            TabIndex        =   65
            Top             =   450
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145752067
            UpDown          =   -1  'True
            CurrentDate     =   37683.5208333333
         End
         Begin MSComCtl2.DTPicker LunchEnd 
            Height          =   285
            Left            =   3735
            TabIndex        =   66
            Top             =   450
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145752067
            UpDown          =   -1  'True
            CurrentDate     =   37683.5625
         End
         Begin MSComCtl2.DTPicker FridayBreakStart 
            Height          =   285
            Left            =   1305
            TabIndex        =   71
            Top             =   1020
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145752067
            UpDown          =   -1  'True
            CurrentDate     =   37958.7083333333
         End
         Begin MSComCtl2.DTPicker FridayBreakEnd 
            Height          =   285
            Left            =   2925
            TabIndex        =   72
            Top             =   1020
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145752067
            UpDown          =   -1  'True
            CurrentDate     =   37958.7083333333
         End
         Begin MSComCtl2.DTPicker FridayCheckOut 
            Height          =   285
            Left            =   4545
            TabIndex        =   75
            Top             =   1020
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145752067
            UpDown          =   -1  'True
            CurrentDate     =   37958.7083333333
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Friday Check Out :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   4545
            TabIndex        =   76
            Top             =   825
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Friday Break End:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   2925
            TabIndex        =   74
            Top             =   795
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Friday Break Start :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   1305
            TabIndex        =   73
            Top             =   810
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Check In:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   35
            Left            =   1305
            TabIndex        =   70
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Check Out :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   36
            Left            =   4950
            TabIndex        =   69
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Lunch End :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   68
            Left            =   3735
            TabIndex        =   68
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Lunch Start :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   69
            Left            =   2520
            TabIndex        =   67
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "HIDDEN"
         Height          =   5235
         Left            =   90
         TabIndex        =   10
         Top             =   1890
         Visible         =   0   'False
         Width           =   7155
         Begin VB.TextBox TBox 
            Height          =   285
            Index           =   9
            Left            =   3615
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   405
            Width           =   1005
         End
         Begin VB.Frame fraHoliday 
            Height          =   675
            Index           =   0
            Left            =   4890
            TabIndex        =   36
            Top             =   1515
            Width           =   2355
            Begin VB.OptionButton optGeneral 
               Caption         =   "No"
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   1
               Left            =   1680
               TabIndex        =   38
               Top             =   330
               Width           =   615
            End
            Begin VB.OptionButton optGeneral 
               Caption         =   "Yes"
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   0
               Left            =   960
               TabIndex        =   37
               Top             =   330
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "Deduct Absent Before Sunday"
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   19
               Left            =   60
               TabIndex        =   39
               Top             =   150
               Width           =   2265
            End
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   0
            Left            =   3390
            TabIndex        =   34
            Top             =   1560
            Width           =   1200
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   2
            Left            =   3390
            TabIndex        =   33
            Top             =   2505
            Width           =   1200
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   3390
            TabIndex        =   32
            Top             =   2835
            Width           =   1200
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   4
            Left            =   3390
            TabIndex        =   31
            Top             =   3165
            Width           =   1200
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   5
            Left            =   3390
            TabIndex        =   30
            Top             =   3825
            Width           =   1200
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   3390
            TabIndex        =   29
            Top             =   4155
            Width           =   1200
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   7
            Left            =   3390
            MaxLength       =   2
            TabIndex        =   28
            Top             =   4470
            Width           =   1200
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   8
            Left            =   3390
            TabIndex        =   27
            Top             =   4785
            Width           =   1200
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   1
            Left            =   3390
            TabIndex        =   26
            Top             =   1875
            Width           =   1200
         End
         Begin VB.Frame fraHoliday 
            Height          =   705
            Index           =   3
            Left            =   4890
            TabIndex        =   22
            Top             =   3225
            Width           =   2355
            Begin VB.OptionButton optPublic 
               Caption         =   "No"
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   3
               Left            =   1680
               TabIndex        =   24
               Top             =   360
               Width           =   555
            End
            Begin VB.OptionButton optPublic 
               Caption         =   "Yes"
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   2
               Left            =   960
               TabIndex        =   23
               Top             =   360
               Width           =   675
            End
            Begin VB.Label Label1 
               Caption         =   "Deduct Absent After Public Holiday"
               ForeColor       =   &H00800000&
               Height          =   405
               Index           =   21
               Left            =   90
               TabIndex        =   25
               Top             =   120
               Width           =   2205
            End
         End
         Begin VB.Frame fraHoliday 
            Height          =   705
            Index           =   2
            Left            =   4890
            TabIndex        =   18
            Top             =   2640
            Width           =   2355
            Begin VB.OptionButton optPublic 
               Caption         =   "Yes"
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   0
               Left            =   960
               TabIndex        =   20
               Top             =   390
               Width           =   675
            End
            Begin VB.OptionButton optPublic 
               Caption         =   "No"
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   1
               Left            =   1680
               TabIndex        =   19
               Top             =   390
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Deduct Absent Before Public Holiday"
               ForeColor       =   &H00800000&
               Height          =   375
               Index           =   22
               Left            =   90
               TabIndex        =   21
               Top             =   150
               Width           =   2115
            End
         End
         Begin VB.Frame fraHoliday 
            Height          =   675
            Index           =   1
            Left            =   4890
            TabIndex        =   14
            Top             =   2070
            Width           =   2355
            Begin VB.OptionButton optGeneral 
               Caption         =   "Yes"
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   2
               Left            =   960
               TabIndex        =   16
               Top             =   330
               Width           =   675
            End
            Begin VB.OptionButton optGeneral 
               Caption         =   "No"
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   3
               Left            =   1680
               TabIndex        =   15
               Top             =   330
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Deduct Absent After Sunday"
               ForeColor       =   &H00800000&
               Height          =   315
               Index           =   20
               Left            =   90
               TabIndex        =   17
               Top             =   120
               Width           =   2205
            End
         End
         Begin VB.TextBox TBox 
            Height          =   315
            Index           =   10
            Left            =   3390
            TabIndex        =   13
            Top             =   3495
            Width           =   1200
         End
         Begin VB.CheckBox chkAbsents 
            Caption         =   "Deduct Absents"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5040
            TabIndex        =   12
            Top             =   4485
            Width           =   2025
         End
         Begin VB.CheckBox chkLateHrs 
            Caption         =   "Deduct Late Hours"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5040
            TabIndex        =   11
            Top             =   4815
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker dtLHours 
            Height          =   315
            Left            =   3390
            TabIndex        =   35
            Top             =   2190
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145817603
            UpDown          =   -1  'True
            CurrentDate     =   37683
         End
         Begin MSComCtl2.DTPicker DTLunch 
            Height          =   285
            Left            =   3555
            TabIndex        =   55
            Top             =   915
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   145817603
            UpDown          =   -1  'True
            CurrentDate     =   37683
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Working Duration :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   2220
            TabIndex        =   57
            Top             =   435
            Width           =   1350
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Lunch Duration :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   2280
            TabIndex        =   56
            Top             =   975
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "How Many Leaves(Abs. With Pay) Allowed :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   53
            Top             =   1665
            Width           =   3150
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "How  Much Late Time Allowed :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1140
            TabIndex        =   52
            Top             =   2310
            Width           =   2220
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "How Many Times Late Time Allowed :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   720
            TabIndex        =   51
            Top             =   2625
            Width           =   2640
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Over Time Hrs Per Day  :"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   3
            Left            =   780
            TabIndex        =   50
            Top             =   2940
            Width           =   2580
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Long Term Loan Limit :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   1755
            TabIndex        =   49
            Top             =   4905
            Width           =   1605
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "After How Much Time Long Term Loan is App. :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   48
            Top             =   4575
            Width           =   3360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H005F4941&
            Height          =   195
            Index           =   9
            Left            =   4635
            TabIndex        =   47
            Top             =   3915
            Width           =   165
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            ForeColor       =   &H005F4941&
            Height          =   195
            Index           =   10
            Left            =   4635
            TabIndex        =   46
            Top             =   4245
            Width           =   165
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "MM"
            ForeColor       =   &H005F4941&
            Height          =   195
            Index           =   11
            Left            =   4620
            TabIndex        =   45
            Top             =   4545
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Short Term Loan :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   2070
            TabIndex        =   44
            Top             =   3930
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Advance Salary Payable  :"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   8
            Left            =   780
            TabIndex        =   43
            Top             =   4260
            Width           =   2580
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "How Many Abs.(Abs. Without Pay) Allowed :"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   23
            Left            =   135
            TabIndex        =   42
            Top             =   1995
            Width           =   3225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Dinner Amount For Over Time  :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   1080
            TabIndex        =   41
            Top             =   3270
            Width           =   2280
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Social Security Amount :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   1620
            TabIndex        =   40
            Top             =   3600
            Width           =   1740
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Departement Settings :"
      Height          =   645
      Left            =   135
      TabIndex        =   5
      Top             =   -15
      Width           =   6615
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Settings For :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   810
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   135
      TabIndex        =   7
      Top             =   -30
      Visible         =   0   'False
      Width           =   3915
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "of Every Month"
         ForeColor       =   &H005F4941&
         Height          =   195
         Index           =   15
         Left            =   2685
         TabIndex        =   9
         Top             =   270
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Advance Salary Paid After :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1995
      End
   End
   Begin MSForms.CommandButton cmdUpdateAll 
      Height          =   360
      Left            =   495
      TabIndex        =   77
      Top             =   6840
      Visible         =   0   'False
      Width           =   2400
      Caption         =   "Update All Departments"
      PicturePosition =   327683
      Size            =   "4233;635"
      Accelerator     =   85
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSaveN 
      Height          =   360
      Left            =   2910
      TabIndex        =   1
      Top             =   6840
      Width           =   1455
      Caption         =   "Save & New"
      PicturePosition =   327683
      Size            =   "2566;635"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   360
      Left            =   4380
      TabIndex        =   2
      Top             =   6840
      Width           =   1575
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2778;635"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   5970
      TabIndex        =   3
      Top             =   6840
      Width           =   1365
      Caption         =   "Cancel   "
      PicturePosition =   327683
      Size            =   "2408;635"
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
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IsNew As Boolean
Public add As Boolean
Dim FirstEntry As Boolean


Private Sub cmbCatagory_Change()
    Call cmbCatagory_Click
End Sub

Private Sub cmbCatagory_Click()
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    LVCat.ListItems.Clear
    With rs
        .Open "Select * From EmployeeAttendanceCatagories Where EntryID=" & cmbCatagory.Text, con, adOpenForwardOnly, adLockReadOnly
        
        Set ITM = LVCat.ListItems.add(, , "Lunch Scanning")
        ITM.SubItems(1) = IIf(!LunchTimeScanning, "Required", "Not Required")
        
        Set ITM = LVCat.ListItems.add(, , "Time Deduction")
        ITM.SubItems(1) = IIf(!LateDeduction, "Yes", "No")
         
        Set ITM = LVCat.ListItems.add(, , "Time Deduction After")
        ITM.SubItems(1) = !LateDeductionAfter & ""
        
        Set ITM = LVCat.ListItems.add(, , "Gate Pass Required")
        ITM.SubItems(1) = IIf(!GatePassRequired, "Yes", "No")
        
        Set ITM = LVCat.ListItems.add(, , "Leaves Recorded")
        ITM.SubItems(1) = IIf(!LeavesRecorded, "Yes", "No")
        
        Set ITM = LVCat.ListItems.add(, , "Over Time Paid")
        ITM.SubItems(1) = IIf(!OverTimePaid, "Yes", "No")
        
        Set ITM = LVCat.ListItems.add(, , "Sandwich Leaves Deduction")
        ITM.ListSubItems.add(, , IIf(!DeductionSpecialAbsents, "Yes", "No")).Tag = Abs(!DeductionSpecialAbsents)
        
        Set ITM = LVCat.ListItems.add(, , "Over Time Rate")
        ITM.ListSubItems.add(, , Val(!OTRate & "")).Tag = Val(!OTRate & "")
        
        Set ITM = LVCat.ListItems.add(, , "Deduction Percentage")
        ITM.ListSubItems.add(, , Val(!DeductionPercentage & "")).Tag = Val(!DeductionPercentage & "")
        
        Set ITM = LVCat.ListItems.add(, , "Saturday Early Out (Mins.)")
        ITM.ListSubItems.add(, , Val(!Saturday_Early_Out_Mins & "")).Tag = Val(!Saturday_Early_Out_Mins & "")
        
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbdept_matched()
    Call Editrec(cmbDept.ID)
    Frame1.Enabled = True
    cmdUpdateAll.Visible = True
End Sub

Private Sub cmdCancel_Click()
    
'    If IsNew Then
'        NewEmployee.chkSpecial.value = False
'    End If

    Unload Me
    
End Sub

Private Function Saved() As Boolean
     
      
    If cmbCatagory.ListIndex = -1 Then
        MsgBox "Invalid Catagory Selected.", vbInformation
        Exit Function
    End If
'    If Val(TBox(3)) = 0 Then
'        MsgBox "Over Time Per Day Could Not Be Zero !", vbInformation
'        TBox(3).SetFocus
'        Exit Function
'    End If
    
    Dim BGHoliday As Integer, AGHoliday As Integer, _
    BPHoliday As Integer, APHoliday As Integer

    'On Error GoTo ERR
    
    
    If Not cmbDept.MatchFound Then
      MsgBox "Please Select The Departement.", vbInformation
      cmbDept.SetFocus
      Exit Function
    
    End If
    
    Dim SQL As String, SQLEmps As String
    Dim LHours As Double, LunchTime As Double
    
    BGHoliday = IIf(optGeneral(0).Value = True, 1, 0)
    AGHoliday = IIf(optGeneral(2).Value = True, 1, 0)
    BPHoliday = IIf(optPublic(0).Value = True, 1, 0)
    APHoliday = IIf(optPublic(2).Value = True, 1, 0)
    
    LHours = Round((dtLHours.Hour * 60 + dtLHours.Minute) / 60, 2)
    LunchTime = Round((DTLunch.Hour * 60 + DTLunch.Minute) / 60, 2)
    
    SQL = "UPDATE GeneralSettings SET AbsAllowed=" & Val(TBox(0) & "") & ",TotalAbsAllowed=" & Val(TBox(1) & "") & _
     ",TimeAllowed=" & LHours & ",NTimes=" & Val(TBox(2) & "") & ",OverTime=" & Val(TBox(3) & "") & ",DinnerAmt=" & _
     Val(TBox(4) & "") & ",ShortTermPer=" & Val(TBox(5) & "") & ",AdvSalPer=" & Val(TBox(6) & "") & _
     ",LongTermTime=" & Val(TBox(7) & "") & ", LongTermLimit = '" & TBox(8) & "',EnterTime='" & _
     Format(TEnter.Value, "HH:mm") & "',ExitTime='" & Format(TExit.Value, "HH:mm") & "',LunchTime=" & LunchTime & _
     ", SocialAmt=" & Val(TBox(10)) & ", BGHoliday='" & BGHoliday & "', AGHoliday='" & AGHoliday & "', BPHoliday='" & _
     BPHoliday & "', APHoliday='" & APHoliday & "',WorkingHrs=" & Val(TBox(9).Tag) & ",DeductAbsents=" & _
     Val(chkAbsents) & ",DeductLHrs=" & Val(chkLateHrs) & ",LunchInTime='" & Format(LunchStart.Value, "HH:mm") & _
     "',LunchOutTime='" & Format(LunchEnd.Value, "HH:mm") & "',EmpAttCat_RefID=" & cmbCatagory.Text & _
     ",FridayBreakStartTime='" & Format(FridayBreakStart, "HH:mm") & "',FridayBreakEndTime='" & Format(FridayBreakEnd, "HH:mm") & _
     "',FridayCheckOut='" & Format(FridayCheckOut, "HH:mm") & _
     "' WHERE DeptID='" & cmbDept.ID & "'"
         
    SQLEmps = "UPDATE EmpSettings SET AbsAllowed=" & Val(TBox(0) & "") & ",TotalAbsAllowed=" & Val(TBox(1) & "") & _
    ",TimeAllowed=" & LHours & ",NTimes=" & Val(TBox(2) & "") & ",OverTime=" & Val(TBox(3) & "") & _
    ",DinnerAmt=" & Val(TBox(4) & "") & ",ShortTermPer=" & Val(TBox(5) & "") & ",AdvSalPer=" & Val(TBox(6) & "") & _
    ",LongTermTime=" & Val(TBox(7) & "") & ", LongTermLimit = '" & TBox(8) & "',EnterTime='" & _
    Format(TEnter.Value, "HH:mm") & "',ExitTime='" & Format(TExit.Value, "HH:mm") & "',LunchTime=" & LunchTime & _
    ", SocialAmt=" & Val(TBox(10)) & ", BGHoliday='" & BGHoliday & "', AGHoliday='" & AGHoliday & "', BPHoliday='" & _
    BPHoliday & "', APHoliday='" & APHoliday & "',WorkingHrs=" & Val(TBox(9).Tag) & ",DeductAbsents=" & _
    Val(chkAbsents) & ",DeductLHrs=" & Val(chkLateHrs) & ",LunchInTime='" & Format(LunchStart.Value, "HH:mm") & _
    "',LunchOutTime='" & Format(LunchEnd.Value, "HH:mm") & "',EmpAttCat_RefID=" & cmbCatagory.Text & _
    ",FridayBreakStartTime='" & Format(FridayBreakStart, "HH:mm") & "',FridayBreakEndTime='" & Format(FridayBreakEnd, "HH:mm") & _
    "',FridayCheckOut='" & Format(FridayCheckOut, "HH:mm") & _
    "' WHERE EmpID IN (SELECT EmpID FROM Employees WHERE DeptID= '" & cmbDept.ID & "' AND ISNULL(ExemptSettings,0)=0 )"
         
    Saved = True
    Frame1.Enabled = False
    If IsNew Then
        NewEmployee.SqlSettings = SQL
        Me.Hide
    Else
        con.Execute SQL, a
        con.Execute SQLEmps, a
        'Unload Me
    End If
    
    Exit Function
err:
    MsgBox err.Description
End Function


Public Sub Editrec(DeptID As String)
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select * from GeneralSettings Where DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            FirstEntry = False
            TEnter.Value = CDate(IIf(IsNull(!EnterTime), "09:00", !EnterTime))
            TExit.Value = CDate(IIf(IsNull(!ExitTime), "17:00", !ExitTime))
            
            LunchStart.Value = CDate(IIf(IsNull(!LunchInTime), "12:30", !LunchInTime))
            LunchEnd.Value = CDate(IIf(IsNull(!LunchOutTime), "13:30", !LunchOutTime))
            
            FridayBreakStart.Value = CDate(IIf(IsNull(!FridayBreakStartTime), "12:15", !FridayBreakStartTime))
            FridayBreakEnd.Value = CDate(IIf(IsNull(!FridayBreakEndTime), "14:00", !FridayBreakEndTime))
            FridayCheckOut.Value = CDate(IIf(IsNull(!FridayCheckOut), "18:00", !FridayCheckOut))
            
            DTLunch.Value = CDate(IIf(IsNull(!LunchTime), "01:00", MinsToTime(Val(!LunchTime & "") * 60)))
            
            TBox(0) = ![AbsAllowed] & ""
            TBox(1) = ![TotalAbsAllowed] & ""
            
            dtLHours = CDate(IIf(IsNull(!TimeAllowed), "09:00", MinsToTime(!TimeAllowed * 60)))
            TBox(2) = ![NTimes] & ""
            TBox(3) = ![OverTime] & ""
            
            TBox(4) = ![DinnerAmt] & ""
            TBox(5) = ![ShortTermPer] & ""
            TBox(6) = ![AdvSalPer] & ""
            
            TBox(7) = ![longtermtime] & ""
            TBox(8) = ![LongTermLimit] & ""
            TBox(10) = Val(!SocialAmt & "")
            chkAbsents = Abs(![DeductAbsents])
            chkLateHrs = Abs(![DeductLHrs])
            
            optGeneral(0).Value = IIf((!BGHoliday = True), True, False)
            optGeneral(1).Value = IIf((!BGHoliday = False), True, False)
            optGeneral(2).Value = IIf((!AGHoliday = True), True, False)
            optGeneral(3).Value = IIf((!AGHoliday = False), True, False)
            
            optPublic(0).Value = IIf((!BPHoliday = True), True, False)
            optPublic(1).Value = IIf((!BPHoliday = False), True, False)
            optPublic(2).Value = IIf((!APHoliday = True), True, False)
            optPublic(3).Value = IIf((!APHoliday = False), True, False)
            TBox(9) = CalcNet
            If Val(!EmpAttCat_RefID & "") <> 0 Then
                cmbCatagory.Text = Val(!EmpAttCat_RefID & "")
            End If
        Else
            FirstEntry = True
        End If
        .Close
    End With
    
    Set rs = Nothing
End Sub

Private Sub cmdSave_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdSaveN_Click()

If Saved Then
    Form_Load
    
    For i = 0 To TBox.count - 1
        TBox(i) = ""
    Next
    
    TEnter.Value = "00:00"
    TExit.Value = "00:00"
    DTLunch.Value = "00:00"
    dtLHours.Value = "00:00"
    
End If

End Sub

Private Sub DTEnter_Change()
TBox(9) = CalcNet

End Sub

Private Sub DTExit_Change()

TBox(9) = CalcNet

End Sub

Private Sub cmdUpdateAll_Click()

    If MsgBox("Are you sure to Update All Departments Timings?", vbQuestion + vbYesNo) = vbYes Then
        con.Execute "UPDATE GeneralSettings SET EnterTime='" & Format(TEnter.Value, "HH:mm") & "',ExitTime='" & _
         Format(TExit.Value, "HH:mm") & "',LunchInTime='" & Format(LunchStart.Value, "HH:mm") & _
         "',LunchOutTime='" & Format(LunchEnd.Value, "HH:mm") & "',FridayBreakStartTime='" & Format(FridayBreakStart, "HH:mm") & _
         "',FridayBreakEndTime='" & Format(FridayBreakEnd, "HH:mm") & _
         "',FridayCheckOut='" & Format(FridayCheckOut, "HH:mm") & "'"
    End If
    
End Sub

Private Sub DTLunch_Change()
    TBox(9) = CalcNet

End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1500
    
    Call cmbDept.ClearVals
    Call cmbDept.AddVals(con, "Name", "Departments", "DeptID", " Where Active=1")
    
    Call AddToCombo(cmbCatagory, "EntryID", "EmployeeAttendanceCatagories", "")
    
End Sub


Private Sub TBox_GotFocus(Index As Integer)

TBox(Index).SelStart = 0
TBox(Index).SelLength = Len(TBox(Index))

End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub TBox_Validate(Index As Integer, cancel As Boolean)

'If Index = 8 And Val(TBox(8)) > 28 Then
'   MsgBox "Invalid Date to Pay Advance Salary", vbInformation
'   TBox(8) = "15"
'End If

End Sub

Private Sub TEnter_Change()
    TBox(9) = CalcNet
End Sub

Private Sub TExit_Change()

TBox(9) = CalcNet

End Sub

Private Function CalcNet() As String

Dim NetHours As Long

NetHours = DateDiff("n", TEnter.Value, TExit.Value)
NetHours = NetHours - DateDiff("n", CDate("00:00"), DTLunch.Value)

TBox(9).Tag = Round(NetHours / 60, 2)
CalcNet = MinsToTime(NetHours)

End Function

