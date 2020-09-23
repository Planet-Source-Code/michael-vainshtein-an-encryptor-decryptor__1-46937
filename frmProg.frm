VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProg 
   Caption         =   "Saving File"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   390
      TabIndex        =   3
      Top             =   2205
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   360
      TabIndex        =   0
      Top             =   210
      Width           =   3915
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Please wait a moment..."
         Height          =   195
         Left            =   1065
         TabIndex        =   1
         Top             =   465
         Width           =   1680
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   630
      Picture         =   "frmProg.frx":0000
      Top             =   1365
      Width           =   480
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   1335
      TabIndex        =   2
      Top             =   1530
      Width           =   2595
   End
End
Attribute VB_Name = "frmProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Me.Tag = "esc"
    End If
End Sub

Private Sub Form_Load()
    Label1 = "Please wait a moment..." & vbNewLine & "Press ESCAPE to cancel."
End Sub
