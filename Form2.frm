VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIND FOLDER"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "NEW FOLDER"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "SELECT FOLDER"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   4455
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNew_Click()
    Dim x As String
    x = InputBox("NEW FOLDER NAME: ", "NEW FOLDER")
    If Len(Dir1.Path) = 3 Then
        MkDir Dir1.Path & x
        Dir1.Path = Dir1.Path & x
    Else
        MkDir Dir1.Path & "\" & x
        Dir1.Path = Dir1.Path & "\" & x
    End If
    Dir1.Refresh
End Sub

Private Sub cmdSelect_Click()
    If Form2.Caption = "SOURCE FOLDER:" Then Form1.txtSource.Text = Dir1.Path
    If Form2.Caption = "DESTINATION FOLDER:" Then Form1.txtDestin.Text = Dir1.Path
    Unload Me
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Dir1.Path = App.Path
End Sub
