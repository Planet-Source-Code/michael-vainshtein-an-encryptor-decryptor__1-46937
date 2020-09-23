VERSION 5.00
Begin VB.Form frmIncDec 
   Caption         =   "Incriptor \ Decriptor v2.5"
   ClientHeight    =   6660
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10215
   Icon            =   "frmIncriptor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtText 
      Height          =   5505
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   345
      Width           =   9120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuNSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save as..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select all"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmIncDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    txtText.Width = Me.ScaleWidth
    txtText.Height = Me.ScaleHeight
    txtText.Top = 0
    txtText.Left = 0
    mnuNSave.Visible = False
End Sub

Private Sub mnuNew_Click()
    If MsgBox("Are you sure you want to delete all text?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
        txtText.Tag = ""
        txtText.Text = ""
        Me.Tag = ""
        txtText_Change
    End If
End Sub

Private Sub mnuNSave_Click()
    If Me.Tag = "" Then
        mnuSave_Click
    Else
        frmSave.SAFE
    End If
End Sub

Private Sub mnuOpen_Click()
    frmOpen.Show
    Me.Enabled = False
End Sub

Private Sub mnuSave_Click()
    frmSave.Show
    Me.Enabled = False
    frmSave.Visible = True
End Sub

Private Sub mnuSelAll_Click()
    With txtText
        If txtText <> "" Then
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

Private Sub txtText_Change()
    If txtText <> "" Then
        mnuSave.Enabled = True
        mnuNSave.Enabled = True
    Else: mnuSave.Enabled = False: mnuNSave.Enabled = False
    End If
    txtText.Tag = "Changed"
End Sub
