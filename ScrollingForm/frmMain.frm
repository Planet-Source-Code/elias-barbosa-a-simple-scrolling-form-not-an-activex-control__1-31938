VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   " Example Project"
   ClientHeight    =   3405
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin prjExample.ScrllngFrm ScrllngFrm1 
      Height          =   3075
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5424
      BackColor       =   -2147483633
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4875
         Left            =   0
         ScaleHeight     =   4875
         ScaleWidth      =   3675
         TabIndex        =   1
         Top             =   240
         Width           =   3675
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Text            =   "Text3"
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   855
            Left            =   120
            TabIndex        =   28
            Top             =   3960
            Width           =   3135
            Begin VB.TextBox Text9 
               Height          =   285
               Left            =   1200
               TabIndex        =   29
               Text            =   "Text9"
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   840
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   3600
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Top             =   3600
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   3600
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Text            =   "Text8"
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Text            =   "Text7"
            Top             =   3240
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1800
            TabIndex        =   22
            Text            =   "Text6"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Text            =   "Text5"
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1800
            TabIndex        =   20
            Text            =   "Text4"
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1800
            TabIndex        =   19
            Text            =   "Text2"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   2160
            Width           =   1575
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option11"
            Height          =   255
            Index           =   11
            Left            =   2520
            TabIndex        =   17
            Top             =   840
            Width           =   1035
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option10"
            Height          =   255
            Index           =   10
            Left            =   2520
            TabIndex        =   16
            Top             =   600
            Width           =   1035
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option9"
            Height          =   255
            Index           =   9
            Left            =   2520
            TabIndex        =   15
            Top             =   360
            Width           =   915
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option8"
            Height          =   255
            Index           =   8
            Left            =   2520
            TabIndex        =   14
            Top             =   120
            Width           =   915
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option7"
            Height          =   255
            Index           =   7
            Left            =   1440
            TabIndex        =   13
            Top             =   840
            Width           =   915
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option6"
            Height          =   255
            Index           =   6
            Left            =   1440
            TabIndex        =   12
            Top             =   600
            Width           =   915
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option5"
            Height          =   255
            Index           =   5
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   915
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option4"
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   10
            Top             =   120
            Width           =   915
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option3"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Width           =   915
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option2"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   915
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   915
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   375
            Left            =   1800
            TabIndex        =   5
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Command4"
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   1680
            Width           =   1575
         End
         Begin VB.OptionButton Option 
            Caption         =   "Option0"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   120
            Value           =   -1  'True
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ScrllngFrm1.Attatch Picture1
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ScrllngFrm1.Detatch
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If (Me.Height < 1500) Then
        Me.Height = 1500
    End If
    ScrllngFrm1.Width = Me.Width - 300
    ScrllngFrm1.Height = Me.Height - 600
    
End Sub
