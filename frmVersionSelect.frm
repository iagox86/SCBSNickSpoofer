VERSION 5.00
Begin VB.Form frmVersionSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Brood War Version..."
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVersionSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmVersion 
      Caption         =   "Version Selection"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cmbVersion 
         Height          =   315
         ItemData        =   "frmVersionSelect.frx":030A
         Left            =   1680
         List            =   "frmVersionSelect.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         Caption         =   "Select your Brood War patch version:"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmVersionSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SaveAndExit()
SaveSetting "NameSpoofer", "Version", "Patch Version", BWVersion
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdConfirm_Click()
If cmbVersion.Text = "1.07" Then BWVersion = "107": SaveAndExit: Exit Sub
If cmbVersion.Text = "1.08 Revision O" Then BWVersion = "108O": SaveAndExit: Exit Sub
MsgBox "You must select a Brood War version!", vbOKOnly + vbCritical, "Error"
End Sub
Private Sub Form_Unload(Cancel As Integer)
If BWVersion <> "107" And BWVersion <> "108O" Then
    If MsgBox("Nick Spoofer will terminate if you do not choose a version. Are you sure you want to exit Nick Spoofer?", vbYesNo + vbQuestion, "Must Choose Version") = vbYes Then Unload Me: End
    Cancel = 1
End If
End Sub
