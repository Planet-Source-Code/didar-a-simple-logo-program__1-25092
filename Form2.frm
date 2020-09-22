VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General"
   ClientHeight    =   2070
   ClientLeft      =   3105
   ClientTop       =   2535
   ClientWidth     =   3315
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3315
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   2520
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "Form2.frx":0ECA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   0
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   1320
   End
   Begin VB.Label Label4 
      Caption         =   "CopyRight By General Corporation                  Bangladesh"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "This Program Written by            Didarul Alam."
      Height          =   555
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1845
   End
   Begin VB.Label Label2 
      Caption         =   "Version 2.03"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logo Animation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1320
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.WindowState = 1
End Sub


Private Sub Timer1_Timer()
frmSrc.Show
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
If Me.WindowState = 0 Then
Me.WindowState = 1
End If
End Sub
