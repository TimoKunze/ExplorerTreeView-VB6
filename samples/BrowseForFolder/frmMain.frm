VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "BrowseForFolder Sample"
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub cmdBrowse_Click()
  Dim frm As frmBrowseForFolder

  Set frm = New frmBrowseForFolder
  frm.SelectedPath = txtPath.Text
  frm.Show FormShowConstants.vbModal, Me
  If frm.Aborted Then
    MsgBox "Abbruch!"
  Else
    MsgBox "Sie haben den folgenden Pfad gewählt:" & vbNewLine & frm.SelectedPath
    txtPath.Text = frm.SelectedPath
  End If
  ' remember that the form is still loaded!!
  Unload frm
  Set frm = Nothing
End Sub

Private Sub Form_Initialize()
  ' use XP style
  InitCommonControls
End Sub
