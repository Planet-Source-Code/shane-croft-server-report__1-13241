VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report for - "
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6840
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stats"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   6720
      TabIndex        =   24
      Top             =   0
      Width           =   3735
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   36
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   35
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   34
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Total Disabled Accounts"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Total Shares"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Total Services"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Total Print Queues"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Total Groups"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Total Users"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   23
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   21
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   19
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   2775
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   6240
      Width           =   5055
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   6240
      Width           =   5055
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   4680
      Width           =   5055
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   4680
      Width           =   4935
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3120
      Width           =   5055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   37
      Top             =   2040
      Width           =   6495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Installed HAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3360
      TabIndex        =   22
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Processor"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3360
      TabIndex        =   20
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Operating System Version"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Operating System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Organization"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Computer Owner"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Disabled Accounts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   5055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Shares"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   6000
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Services"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   4440
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Print Queues"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Groups"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4935
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuchoose 
         Caption         =   "Choose Another"
      End
      Begin VB.Menu menurefresh 
         Caption         =   "Refresh Report"
      End
      Begin VB.Menu menuprint 
         Caption         =   "Print/Print Preview"
      End
      Begin VB.Menu menuline1 
         Caption         =   "-"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub GetReport()
'Have to have On error resume next or it won't work
On Error Resume Next
Label19.Caption = "Status: Loading..."
DoEvents
Dim Computer
Dim ComputerName
Dim Service
Dim FileService
Dim FileShare
Dim PrintQueue
Dim ContainerPrint
Dim container
Label19.Caption = "Status: Setting..."
DoEvents
ComputerName = FrmChoose.Combo1.Text
Set Computer = GetObject("WinNT://" & ComputerName & ",computer")
Set FileService = GetObject("WinNT://" & ComputerName & "/LanmanServer")
Set ContainerPrint = GetObject("WinNT://" & ComputerName)
Set container = GetObject("WinNT://" & ComputerName)
DoEvents
Label19.Caption = "Status: Getting Owner..."
DoEvents
Text1.Text = Computer.Owner
DoEvents
Label19.Caption = "Status: Getting Organization..."
DoEvents
Text2.Text = Computer.Division
DoEvents
Label19.Caption = "Status: Getting Operating System..."
DoEvents
Text3.Text = Computer.OperatingSystem
DoEvents
Label19.Caption = "Status: Getting Version..."
DoEvents
Text4.Text = Computer.OperatingSystemVersion
DoEvents
Label19.Caption = "Status: Getting Processor..."
DoEvents
Text5.Text = Computer.Processor
DoEvents
Label19.Caption = "Status: Getting Installed HAL..."
DoEvents
Text6.Text = Computer.ProcessorCount
DoEvents
Label19.Caption = "Status: Getting Services..."
DoEvents
For Each Service In Computer
    List4.AddItem Service.DisplayName
Next
DoEvents
Label19.Caption = "Status: Getting File Shares..."
DoEvents
For Each FileShare In FileService
     List5.AddItem FileShare.Name
Next
DoEvents
Label19.Caption = "Status: Getting Print Queue..."
DoEvents
ContainerPrint.Filter = Array("PrintQueue")
For Each PrintQueue In ContainerPrint
     List3.AddItem PrintQueue.Name
Next
DoEvents
Label19.Caption = "Status: Getting Users and Account Disabled Information..."
DoEvents
container.Filter = Array("User")
Dim user As IADsUser
For Each user In container
List1.AddItem user.Name
    If user.AccountDisabled = True Then
        List6.AddItem user.Name
    End If
Next
DoEvents
Label19.Caption = "Status: Getting Groups..."
DoEvents
container.Filter = Array("Group")
Dim group As IADsGroup
For Each group In container
List2.AddItem group.Name
Next
DoEvents
Label19.Caption = "Status: Done!!!"
DoEvents
End Sub
Private Sub CleanUP()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Label19.Caption = "Status:"
End Sub

Private Sub Form_Load()
Timer2.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call CleanUP
DoEvents
FrmChoose.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call CleanUP
DoEvents
FrmChoose.Show
End Sub

Private Sub menuchoose_Click()
Call CleanUP
DoEvents
FrmChoose.Show
Unload Me
End Sub

Private Sub menuexit_Click()
End
End Sub

Private Sub menuprint_Click()
On Error Resume Next
Dim X As Long
Dim X2 As Long
Dim X3 As Long
Dim X4 As Long
Dim X5 As Long
Dim X6 As Long

FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem "Report made on  " & Date & "  at  " & Time
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem "(Domain or Computer) - " & FrmChoose.Combo1.Text
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem vbTab & vbTab & "(Computer Owner) - " & FrmMain.Text1.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Organization) - " & FrmMain.Text2.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Operating System) - " & FrmMain.Text3.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(OS Version) - " & FrmMain.Text4.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Processor) - " & FrmMain.Text5.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Installed HAL) - " & FrmMain.Text6.Text
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem vbTab & "(Stats)"
FrmPrint.List1.AddItem vbTab & vbTab & "(Total Users) - " & FrmMain.Text7.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Total Groups) - " & FrmMain.Text8.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Total Print Queues) - " & FrmMain.Text9.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Total Services) - " & FrmMain.Text10.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Total Shares) - " & FrmMain.Text11.Text
FrmPrint.List1.AddItem vbTab & vbTab & "(Total Accounts Disabled) - " & FrmMain.Text12.Text
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem vbTab & "(Users)"
    For X = 0 To FrmMain.List1.ListCount - 1
FrmPrint.List1.AddItem vbTab & vbTab & FrmMain.List1.List(X)
    Next X
DoEvents
DoEvents
DoEvents
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem vbTab & "(Groups)"
    For X2 = 0 To FrmMain.List2.ListCount - 1
FrmPrint.List1.AddItem vbTab & vbTab & FrmMain.List2.List(X2)
    Next X2
DoEvents
DoEvents
DoEvents
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem vbTab & "(Print Queues)"
    For X3 = 0 To FrmMain.List3.ListCount - 1
FrmPrint.List1.AddItem vbTab & vbTab & FrmMain.List3.List(X3)
    Next X3
DoEvents
DoEvents
DoEvents
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem vbTab & "(Services)"
    For X4 = 0 To FrmMain.List4.ListCount - 1
FrmPrint.List1.AddItem vbTab & vbTab & FrmMain.List4.List(X4)
    Next X4
DoEvents
DoEvents
DoEvents
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem vbTab & "(Shares)"
    For X5 = 0 To FrmMain.List5.ListCount - 1
FrmPrint.List1.AddItem vbTab & vbTab & FrmMain.List5.List(X5)
    Next X5
DoEvents
DoEvents
DoEvents
FrmPrint.List1.AddItem ""
FrmPrint.List1.AddItem vbTab & "(Disabled Accounts)"
    For X6 = 0 To FrmMain.List6.ListCount - 1
FrmPrint.List1.AddItem vbTab & vbTab & FrmMain.List6.List(X6)
    Next X6
DoEvents
DoEvents
DoEvents
FrmPrint.List1.AddItem ""
FrmPrint.Show vbModal, Me
End Sub

Private Sub menurefresh_Click()
Call CleanUP
DoEvents
Call GetReport
DoEvents
End Sub

Private Sub Timer1_Timer()
Text7.Text = List1.ListCount
Text8.Text = List2.ListCount
Text9.Text = List3.ListCount
Text10.Text = List4.ListCount
Text11.Text = List5.ListCount
Text12.Text = List6.ListCount
End Sub

Private Sub Timer2_Timer()
DoEvents
Call GetReport
Timer2.Enabled = False
End Sub
