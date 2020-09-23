VERSION 5.00
Begin VB.Form FrmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Nbr As Long
    Printer.Print
    For Nbr = 0 To FrmPrint.List1.ListCount - 1
        Printer.Print FrmPrint.List1.List(Nbr)
    Next Nbr
    Printer.EndDoc
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
