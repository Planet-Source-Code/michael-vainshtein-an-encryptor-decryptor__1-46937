VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5355
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2700
      Pattern         =   "*.coded"
      TabIndex        =   3
      Top             =   495
      Width           =   2685
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   30
      TabIndex        =   2
      Top             =   495
      Width           =   2505
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   4050
      Width           =   2295
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   4110
      Width           =   2625
   End
   Begin VB.Label Label1 
      Caption         =   "File name:"
      Height          =   240
      Left            =   165
      TabIndex        =   5
      Top             =   3840
      Width           =   810
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpen_Click()
    Dim FileText
    Dim nFileText
    Dim A
    Dim PP
    Dim K
    
    
    Me.Hide
    If frmSave.CH(txtFile) Then
        frmProg.Show
        frmProg.ProgressBar1.Value = 0
        frmProg.lblStatus = "Reading data..."
        frmProg.Caption = "Opening " & txtFile
        Open File1.Path & "\" & txtFile For Input As #1
            FileText = Input$(LOF(1), #1)
        Close #1
        
        FileText = Mid(FileText, 2, Len(FileText) - 3)
        
        frmProg.lblStatus = "Decripting data..."
        
        For A = 1 To (Len(FileText) - 1) / 3
            K = K & Chr(Asc(Mid(FileText, A * 3 + 1, 1)) + 2)
            PP = (A / (Len(FileText) - 1) * 3) * 100
            frmProg.ProgressBar1.Value = PP
            frmProg.lblStatus = "Decripting data... " & PP & "%"
            DoEvents
            If frmProg.Tag = "esc" Then A = (Len(FileText) - 1) / 3
        Next
        
        If frmProg.Tag <> "esc" Then
            nFileText = Chr(Asc(Mid(FileText, 1, 1)) + 2) & Mid(K, 1, Len(K) - 1)
            
            frmIncDec.txtText = nFileText
        End If
    Else: MsgBox "File name must not contain the following characters: \ / ? | < > * " & Chr(34), vbCritical, "Error"
    End If
    
    frmProg.Tag = ""
    frmProg.Hide
    Me.Hide
    frmIncDec.Enabled = True
    frmIncDec.Show
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    txtFile = File1.FileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmIncDec.Enabled = True
End Sub
