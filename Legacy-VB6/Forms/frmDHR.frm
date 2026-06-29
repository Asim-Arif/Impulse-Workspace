VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDHR 
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   12390
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      Height          =   10485
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.Frame Frame2 
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3135
         Left            =   60
         TabIndex        =   21
         Top             =   6750
         Width           =   12240
         Begin VB.Frame Frame4 
            Caption         =   "Item Details"
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
            Height          =   2835
            Left            =   6165
            TabIndex        =   23
            Top             =   210
            Width           =   5955
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   17
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   1110
               Width           =   2910
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   16
               Left            =   1545
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   510
               Width           =   4305
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   15
               Left            =   2055
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   2400
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   14
               Left            =   150
               Locked          =   -1  'True
               TabIndex        =   54
               Top             =   2400
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   13
               Left            =   4485
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   1110
               Width           =   1365
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   12
               Left            =   3960
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   1755
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   11
               Left            =   2055
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   1755
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   10
               Left            =   150
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   510
               Width           =   1365
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   9
               Left            =   150
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   1110
               Width           =   1365
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   8
               Left            =   150
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   1755
               Width           =   1875
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Lot No."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   17
               Left            =   1590
               TabIndex        =   59
               Top             =   900
               Width           =   525
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Item Name"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   16
               Left            =   1590
               TabIndex        =   57
               Top             =   300
               Width           =   780
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Man. End Date"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   15
               Left            =   2070
               TabIndex        =   53
               Top             =   2190
               Width           =   1065
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Man. Start Date"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   14
               Left            =   210
               TabIndex        =   52
               Top             =   2190
               Width           =   1155
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Total Steel Used"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   13
               Left            =   4515
               TabIndex        =   51
               Top             =   900
               Width           =   1170
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Qty Released"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   12
               Left            =   3960
               TabIndex        =   49
               Top             =   1530
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Qty Rejected"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   11
               Left            =   2070
               TabIndex        =   47
               Top             =   1530
               Width           =   960
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Product Code"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   10
               Left            =   195
               TabIndex        =   45
               Top             =   300
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Batch No."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   9
               Left            =   195
               TabIndex        =   44
               Top             =   900
               Width           =   705
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Qty Manufactured"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   8
               Left            =   180
               TabIndex        =   43
               Top             =   1530
               Width           =   1320
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Shipment Details"
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
            Height          =   2835
            Left            =   90
            TabIndex        =   22
            Top             =   210
            Width           =   5955
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   2055
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   1755
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   150
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   1755
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   3960
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   525
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   3960
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   1110
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   2055
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   1110
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   150
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   1110
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   2055
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   525
               Width           =   1875
            End
            Begin VB.TextBox txts 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   150
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   525
               Width           =   1875
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Man. End Date"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   7
               Left            =   2070
               TabIndex        =   37
               Top             =   1515
               Width           =   1065
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Man. Start Date"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   180
               TabIndex        =   36
               Top             =   1515
               Width           =   1155
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Total Steel Used"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   5
               Left            =   3960
               TabIndex        =   35
               Top             =   330
               Width           =   1170
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Qty Released"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   4
               Left            =   3990
               TabIndex        =   33
               Top             =   900
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Qty Rejected"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   3
               Left            =   2070
               TabIndex        =   31
               Top             =   900
               Width           =   960
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Qty Manufactured"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   28
               Top             =   900
               Width           =   1320
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Batch No."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   2070
               TabIndex        =   26
               Top             =   300
               Width           =   705
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "No. Of Cartons"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   165
               TabIndex        =   24
               Top             =   300
               Width           =   1095
            End
         End
      End
      Begin VB.TextBox txtPercentage 
         Height          =   315
         Left            =   6990
         TabIndex        =   17
         Top             =   945
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   12375
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "D H R"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   390
            Index           =   8
            Left            =   75
            TabIndex        =   2
            Top             =   165
            Width           =   12225
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "D H R"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   9
            Left            =   90
            TabIndex        =   3
            Top             =   180
            Width           =   12210
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5400
         Left            =   60
         TabIndex        =   5
         Top             =   1290
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   9525
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Process"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Mfgin Period"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty Prod."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty Fail"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker StartDT 
         Height          =   300
         Left            =   10500
         TabIndex        =   15
         Top             =   945
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   53280771
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker FinishDT 
         Height          =   300
         Left            =   12120
         TabIndex        =   16
         Top             =   945
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   53280771
         CurrentDate     =   37250
      End
      Begin MSForms.CommandButton cmdFinal 
         Height          =   405
         Left            =   3810
         TabIndex        =   20
         Top             =   9975
         Width           =   2085
         ForeColor       =   0
         Caption         =   "Make This Final"
         PicturePosition =   327683
         Size            =   "3678;714"
         Accelerator     =   70
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CheckBox chkHide 
         Height          =   360
         Left            =   4875
         TabIndex        =   19
         Top             =   900
         Width           =   2040
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "3598;635"
         Value           =   "0"
         Caption         =   "Show All Proformas"
         Accelerator     =   73
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdHide 
         Height          =   405
         Left            =   5940
         TabIndex        =   18
         Top             =   9975
         Width           =   2085
         ForeColor       =   0
         Caption         =   "Hide This Proforma"
         PicturePosition =   327683
         Size            =   "3678;714"
         Accelerator     =   72
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label5 
         Height          =   240
         Left            =   10500
         TabIndex        =   14
         Top             =   690
         Visible         =   0   'False
         Width           =   1590
         BackColor       =   11517387
         Caption         =   " Man. Start Date"
         Size            =   "2805;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label4 
         Height          =   240
         Left            =   12120
         TabIndex        =   13
         Top             =   690
         Visible         =   0   'False
         Width           =   1560
         BackColor       =   11517387
         Caption         =   " Man. Finish Date"
         Size            =   "2752;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   405
         Left            =   8070
         TabIndex        =   12
         Top             =   9975
         Width           =   2085
         ForeColor       =   0
         Caption         =   "  Print             "
         PicturePosition =   327683
         Size            =   "3678;714"
         Picture         =   "frmDHR.frx":0000
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRCalc 
         Height          =   405
         Left            =   150
         TabIndex        =   11
         Top             =   9975
         Visible         =   0   'False
         Width           =   2580
         ForeColor       =   0
         Caption         =   "Recalculate This Order"
         PicturePosition =   327683
         Size            =   "4551;714"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label3 
         Height          =   240
         Left            =   2130
         TabIndex        =   10
         Top             =   705
         Width           =   2685
         BackColor       =   11517387
         Caption         =   " Item Code"
         Size            =   "4736;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbItems 
         Height          =   315
         Left            =   2115
         TabIndex        =   9
         Top             =   945
         Width           =   2715
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4789;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdCalculate 
         Height          =   405
         Left            =   8055
         TabIndex        =   8
         Top             =   870
         Width           =   1800
         ForeColor       =   0
         Caption         =   "Reset"
         PicturePosition =   327683
         Size            =   "3175;714"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbOrders 
         Height          =   315
         Left            =   75
         TabIndex        =   7
         Top             =   945
         Width           =   2040
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3598;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   705
         Width           =   2010
         BackColor       =   11517387
         Caption         =   " Proforma #"
         Size            =   "3545;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   405
         Left            =   10215
         TabIndex        =   4
         Top             =   9975
         Width           =   2085
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "3678;714"
         Picture         =   "frmDHR.frx":0112
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
End
Attribute VB_Name = "frmDHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkHide_Click()
    Call AddOrders
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()

    If cmbItems.MatchFound = False Then Exit Sub
    'Call FillList
    If cmbOrders.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select ItemCode,IsFinal From DHR Where PInvoice='" & cmbOrders.Text & "' And OrderNo='" & cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & cmbItems.List(cmbItems.ListIndex, 3) & "'", con, adOpenForwardOnly, adLockReadOnly
        
        'Temp = Val(con.Execute("Select Count(ItemCode) From DHR Where PInvoice='" & cmbOrders.Text & "' And ItemCode='" & cmbItems & "'").Fields(0).value & "")
    
        If .RecordCount = 0 Then
            'cmdCalculate.Caption = "Calculate"
        
            'Label3.Visible = True
            'cmbItems.Visible = True
            'cmdRCalc.Visible = True
            cmdCalculate.Visible = False
            txtPercentage.Visible = False
        Else
            'cmdCalculate.Caption = "Make This Final"
        
            'Label3.Visible = True
            'cmbItems.Visible = True
            'cmdRCalc.Visible = True
            Call FillList
            'If .Fields(1) Then
            '    'cmdRCalc.Visible = False
            '    cmdCalculate.Enabled = False
            'Else
            '    'cmdRCalc.Visible = True
            '    cmdCalculate.Enabled = True
            'End If
            If Not CBool(cmbOrders.List(cmbOrders.ListIndex, 2)) Then
                cmdCalculate.Visible = True
                txtPercentage.Visible = True
            End If
        End If
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbOrders_Change()
    Call cmbOrders_Click
End Sub

Private Sub cmbOrders_Click()

    On Error GoTo err
    'Debug.Print "Start : " & Now
    If cmbOrders.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    
    With rs
        '.Open "Select ItemCode,dbo.F_PerIncrease(Qty) From VProformaItems Where PInvoice='" & cmbOrders.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select ItemCode,Qty,OrderNo From VProformaItems Where PInvoice='" & cmbOrders.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        'Debug.Print "Ist rs Opened : " & Now
        cmbItems.Clear
        Do Until .EOF
            cmbItems.AddItem ![ItemCode] & " [" & ![OrderNo] & "]"
            cmbItems.List(.AbsolutePosition - 1, 1) = .Fields(1) & ""
            cmbItems.List(.AbsolutePosition - 1, 2) = .Fields(2) & ""
            cmbItems.List(.AbsolutePosition - 1, 3) = .Fields(0) & ""
            'cmbItems.List(.AbsolutePosition - 1, 2) = .Fields(2) & ""
            .MoveNEXT
        Loop
        .Close
        'Debug.Print "2nd rs Opened : " & Now
        .Open "Select CustomInvoice,TotalCartons,BatchNo,TotalSteelWeight," & _
              "Min(ManStartDate) As ManStartDate,Max(ManEndDate) As ManEndDate," & _
              "Sum(Qty) As ManQty,Sum(TotalWasteges) As Wasteges," & _
              "Sum(Qty-TotalWasteges) As RelQty From VrptDHR Where " & _
              "CustomInvoice='" & cmbOrders.Text & "' Group By " & _
              "CustomInvoice,TotalCartons,BatchNo,TotalSteelWeight", con, adOpenForwardOnly, adLockReadOnly
            
        txts(0) = ![TotalCartons] & ""
        txts(1) = ![BatchNo] & ""
        txts(5) = Round(Val(![TotalSteelWeight] & ""))
        txts(2) = ![ManQty] & ""
        txts(3) = ![Wasteges] & ""
        txts(4) = ![RelQty] & ""
        txts(6) = Format(![ManStartDate] & "", "dd-MMM-yyyy")
        txts(7) = Format(![ManEndDate] & "", "dd-MMM-yyyy")
        
        .Close
    End With
    
    Set rs = Nothing
    
    If Not CBool(cmbOrders.List(cmbOrders.ListIndex, 1)) Then
        cmdHide.Caption = "Hide This Proforma"
    Else
        cmdHide.Caption = "Unhide This Proforma"
    End If
    
    If CBool(cmbOrders.List(cmbOrders.ListIndex, 2)) Then
        cmdFinal.Visible = False
        cmdCalculate.Visible = False
        txtPercentage.Visible = False
    Else
        cmdFinal.Visible = True
        cmdCalculate.Visible = True
        txtPercentage.Visible = True
    End If
    
    LV.ListItems.Clear
    'Debug.Print "End : " & Now
    If cmbItems.ListCount > 0 Then cmbItems.ListIndex = 0
    Debug.Print "List Index Set: " & Now
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub FillItems()
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select ItemCode,dbo.F_PerIncrease(Qty) From FOrderItems Where OrderNo='" & cmbOrders.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbItems.Clear
        Do Until .EOF
            cmbItems.AddItem ![ItemCode] & ""
            cmbItems.List(.AbsolutePosition - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    If cmbItems.ListCount > 0 Then cmbItems.ListIndex = 0
    Set rs = Nothing
    
    'cmdCalculate.Visible = False
    'Label3.Visible = True
    'cmbItems.Visible = True
    'cmdRCalc.Visible = True
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub FillList()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lQty As Long
    
    
    
    'lQty = cmbItems.List(cmbItems.ListIndex, 1)
    lQty = Val(con.Execute("Select Sum(ExtraPer) From DHRDetail Where PInvoice='" & cmbOrders & "' AND OrderNo='" & cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & cmbItems.List(cmbItems.ListIndex, 3) & "'").Fields(0).Value & "")
    'lQty = lQty + Int(Rnd * (2 - (0 - 1)) + (0)) + cmbItems.List(cmbItems.ListIndex, 1)
    lQty = lQty + cmbItems.List(cmbItems.ListIndex, 1)
    
    Dim iTotal As Long
    iTotal = 0
    
    With rs
        .Open "Select * From VDHR Where PInvoice='" & cmbOrders.Text & "' And OrderNo='" & cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & cmbItems.List(cmbItems.ListIndex, 3) & "' Order By SNo", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , ![Description] & "")
            If ![NotApplicable] Then
                ITM.SubItems(1) = "NOT APPLICABLE"
                ITM.SubItems(2) = "-"
                ITM.SubItems(3) = "-"
            Else
                ITM.SubItems(1) = Format(![DTIn] & "", "dd-MMM-yyyy") & "  -  " & Format(![DTOut] & "", "dd-MMM-yyyy")
                lQty = lQty - Val(![WasteQty] & "")
                ITM.SubItems(2) = lQty
                ITM.SubItems(3) = Val(![WasteQty] & "")
                iTotal = iTotal + Val(![WasteQty] & "")
            End If
            
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select CustomInvoice,DT,ItemCode,CompleteItemName,BatchNo,Qty," & _
              "TotalWasteges,ManStartDate,ManEndDate,ForgingWeight From VrptDHR Where " & _
              "CustomInvoice='" & cmbOrders.Text & "' And OrderNo='" & _
              cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & _
              cmbItems.List(cmbItems.ListIndex, 3) & "'", con, adOpenForwardOnly, adLockReadOnly
              
        txts(10) = ![ItemCode] & ""
        txts(16) = ![CompleteItemName] & ""
        txts(9) = ![BatchNo] & ""
        txts(17) = Format(![DT], "MM") & ![BatchNo] & "/" & ![ItemCode]
        txts(13) = (Val(![ForgingWeight] & "") * Val(![Qty] & "")) / 1000
        txts(8) = ![Qty] & ""
        txts(11) = ![TotalWasteges]
        txts(12) = Val(![Qty] & "") - Val(![TotalWasteges] & "")
        txts(14) = ![ManStartDate] & ""
        txts(15) = ![ManStartDate] & ""
        
        .Close
    End With
    Set rs = Nothing
    
    Set ITM = LV.ListItems.add(, , "")
    ITM.ListSubItems.add , , ""
    ITM.ListSubItems.add(, , LV.ListItems(IIf(LV.ListItems.count > 1, LV.ListItems.count - 1, LV.ListItems.count)).SubItems(2)).ForeColor = vbRed
    ITM.ListSubItems.add(, , iTotal).ForeColor = vbRed
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub OldcmdCalculate_Click()

    On Error GoTo err
    
    If cmbOrders.MatchFound = False Then
        MsgBox "Select OrdeNo From The List.", vbInformation
        Exit Sub
    End If
    
    If cmbItems.MatchFound = False Then
        MsgBox "Select Item From The List.", vbInformation
    End If
    
    
    
    
    If cmdCalculate.Caption = "Make This Final" Then
        
        'Check
        
              
        con.Execute "Update DHR Set DHRFinal=1 Where PInvoice='" & cmbOrders.Text & "' And ItemCode='" & cmbItems.Text & "'"
        'cmdRCalc.Visible = False
        cmdCalculate.Enabled = False
        'cmbItems.List(cmbItems.ListIndex, 2) = True
        
    Else
       '
        'Dim PMode As Long
        'PMode = Val(con.Execute("Select PackingMode From FCustomerCatalog Where CustCode+Country=(Select CustCode+Country from FCustomerOrders Where OrderNo='" & cmbOrders.Text & "')").Fields(0).value & "")
        'If PMode = 0 Then
        '    MsgBox "No Packing Mode Define From This Item.", vbInformation
        '    Exit Sub
        'End If
        
        Dim cmd As New ADODB.Command
        With cmd
            Set .ActiveConnection = con
            .CommandType = adCmdStoredProc
            .CommandText = "SP_CalcDHR"
            .Parameters("@PInvoice").Value = cmbOrders.Text
            .Parameters("@ItemID").Value = cmbItems.Text
            .Parameters("@ManStartDT").Value = StartDT
            .Parameters("@ManFinishDT").Value = FinishDT
            .Execute
        End With
        cmdCalculate.Caption = "Make This Final"
        Call FillList
        
    End If
    
    'Call FillItems
   
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCalculate_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    Dim lVal As Long
    Dim AbslVal As Long
    lVal = CLng(Val(txtPercentage))
    AbslVal = Abs(lVal)
    
    
    If lVal = 0 Then
        MsgBox "Invalid Percentage", vbInformation
        Exit Sub
    End If
    
    Call StartTrans(con)
    
    If lVal < 0 Then
        con.Execute "Update DHR Set WasteQty=WasteQty-Cast (WasteQty*" & AbslVal & "/100" & " As Int) Where PInvoice='" & cmbOrders & "' And OrderNo='" & cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & cmbItems.List(cmbItems.ListIndex, 3) & "'"
    Else
        con.Execute "Update DHR Set WasteQty=WasteQty+Cast (WasteQty*" & AbslVal & "/100" & " As Int) Where PInvoice='" & cmbOrders & "' And OrderNo='" & cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & cmbItems.List(cmbItems.ListIndex, 3) & "'"
    End If
    
    Dim lQty As Long
    Dim QtyInStock As Integer
    Dim ProcID As Integer
    
    lQty = Val(con.Execute("Select Sum(WasteQty) From DHR Where PInvoice='" & cmbOrders & "' And OrderNo='" & cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & cmbItems.List(cmbItems.ListIndex, 3) & "'").Fields(0).Value & "")
    
    With rs
        .Open "Select ProcessID From FDAProcesses Where MinWastePer>=2", con, adOpenForwardOnly, adLockReadOnly
        ProcID = Int(Rnd * (.RecordCount))
        .Move ProcID
        ProcID = .Fields(0)
        .Close
    End With
    'Get Number 0,1 or 2 That Should Go In Stock
    QtyInStock = Int(Rnd * (3))
    
    con.Execute "Update DHRDetail Set ExtraPer=" & lQty & ",InStock=" & QtyInStock & ",ProcessID=" & ProcID & "  Where PInvoice='" & cmbOrders & "' And OrderNo='" & cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & cmbItems.List(cmbItems.ListIndex, 3) & "'"
    
    
    Call UpdateRequistion(cmbOrders)
    
    con.CommitTrans
    
    
    Call cmbItems_Click
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFinal_Click()

    On Error GoTo err
    
    If MsgBox("Are You Sure You Want To Make This DHR Final?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    con.Execute "Update DHRDetail1 Set Final=1 Where PInvoice='" & cmbOrders & "'"
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdHide_Click()
    
    On Error GoTo err
    Dim str As String
    If cmdHide.Caption = "Unhide This Proforma" Then
        str = "Update DHRDetail1 Set Hidden=0 Where PInvoice='" & cmbOrders & "'"
    Else
        If MsgBox("Are You Sure to Hide This Entry?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        str = "Update DHRDetail1 Set Hidden=1 Where PInvoice='" & cmbOrders & "'"
    End If
    
    con.Execute str
    
    Call AddOrders
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptDHR.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VRptDHR.CustomInvoice}='" & cmbOrders & "' ", rpt
     
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdRCalc_Click()
    
    On Error GoTo err
    If cmbOrders.MatchFound = False Then Exit Sub
    If cmbItems.MatchFound = False Then Exit Sub
    
    con.Execute "Delete From DHR Where PInvoice='" & cmbOrders.Text & "' And OrderNo='" & cmbItems.List(cmbItems.ListIndex, 2) & "' And ItemCode='" & cmbItems.List(cmbItems.ListIndex, 3) & "'"
    
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_CalcDHR"
        .Parameters("@PInvoice").Value = cmbOrders.Text
        .Parameters("@OrderNo").Value = cmbItems.List(cmbItems.ListIndex, 2)
        .Parameters("@ItemID").Value = cmbItems.List(cmbItems.ListIndex, 3)
        .Parameters("@ManStartDT").Value = StartDT
        .Parameters("@ManFinishDT").Value = FinishDT
        .Execute
    End With
    
    Call FillList
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    Call AddOrders
    StartDT = Date
    FinishDT = Date
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub AddOrders()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    If chkHide.Value Then
        Cond = ""
    Else
        Cond = "Where Hidden=" & Abs(chkHide.Value)
    End If
    
    With rs
        .Open "Select PInvoice,Hidden,Final From DHRDetail1 " & Cond & " Order By PInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbOrders.Clear
        cmbItems.Clear
        LV.ListItems.Clear
        Do Until .EOF
            cmbOrders.AddItem ![PInvoice] & ""
            cmbOrders.List(cmbOrders.ListCount - 1, 1) = ![Hidden] & ""
            cmbOrders.List(cmbOrders.ListCount - 1, 2) = ![Final] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
