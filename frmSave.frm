VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   285
      TabIndex        =   4
      Top             =   4260
      Width           =   2625
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   150
      TabIndex        =   2
      Top             =   645
      Width           =   2505
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2820
      TabIndex        =   1
      Top             =   645
      Width           =   2685
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   5355
   End
   Begin VB.Label Label1 
      Caption         =   "File name:"
      Height          =   240
      Left            =   285
      TabIndex        =   5
      Top             =   3990
      Width           =   810
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    SAFE
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo ER
    Dir1.Path = Drive1.Drive
    Exit Sub
ER:
If MsgBox("Device is unavailable. Retry?", vbCritical + vbRetryCancel, "Error") = vbRetry Then
        Drive1_Change
    Else: Dir1.Path = "C:\"
End If
End Sub

Private Sub File1_Click()
    txtFile = File1.FileName
End Sub

Public Function Hirbush(Num) As String
    Dim B
    
    For B = 0 To Num - 1
        Hirbush = Hirbush & Chr(Int(Rnd * 200 + 36))
    Next
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmIncDec.Enabled = True
End Sub


Public Function CH(Stri As String) As Boolean
    Dim O
    CH = True
    If Stri <> "" Then
        For O = 1 To Len(Stri)
            If Mid(Stri, O, 1) = "\" _
            Or Mid(Stri, O, 1) = "/" _
            Or Mid(Stri, O, 1) = "?" _
            Or Mid(Stri, O, 1) = "|" _
            Or Mid(Stri, O, 1) = ":" _
            Or Mid(Stri, O, 1) = "*" _
            Or Mid(Stri, O, 1) = "<" _
            Or Mid(Stri, O, 1) = ">" _
            Or Mid(Stri, O, 1) = "" Then CH = False
        Next
    Else: CH = False
    End If
End Function

Public Sub SAFE()
    Dim FileText
    Dim nFileText
    Dim A
    Dim Prog

    frmSave.Visible = False
    
    If CH(txtFile) Then
        FileText = frmIncDec.txtText
        
        frmProg.Show
        frmProg.ProgressBar1.Value = 0
        frmProg.lblStatus = "Incripting data..."
        frmProg.Caption = "Saving " & frmSave.txtFile
        
        For A = 1 To Len(FileText)
            DoEvents
            nFileText = _
               nFileText & Chr(Asc(Right(Left(FileText, A), 1)) - 2) & Hirbush(2)
            Prog = A / Len(FileText) * 100
            frmProg.ProgressBar1.Value = Prog
            frmProg.lblStatus = "Incripting data... " & Left(Prog, 2) & "%"
            If frmProg.Tag = "esc" Then A = Len(FileText)
        Next
        
        If frmProg.Tag <> "esc" Then
            frmProg.ProgressBar1.Value = 0
            
            frmProg.lblStatus = "Writing to disc..."
            
            If Len(txtFile) > 6 Then
                If Right(txtFile, 6) <> ".coded" Then
                    frmSave.txtFile = frmSave.txtFile & ".coded"
                End If
            Else
                frmSave.txtFile = frmSave.txtFile & ".coded"
            End If
            
            frmIncDec.Tag = frmSave.File1.Path & "\" & frmSave.txtFile: frmIncDec.txtText.Tag = ""
            
            Open frmSave.File1.Path & "\" & frmSave.txtFile For Output As #1
                Write #1, nFileText
            Close #1
        End If
        
        frmIncDec.Enabled = True
        frmIncDec.Show
    
    Else: MsgBox "File name must not contain the following characters: \ / ? | < > * " & Chr(34), vbCritical, "Error"
    End If
    
    frmIncDec.Enabled = True
    frmProg.Tag = ""
    Unload frmProg
End Sub
