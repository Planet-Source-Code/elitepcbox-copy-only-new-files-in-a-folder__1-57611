VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COPY NEW FILES ONLY"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFD 
      Caption         =   "..."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdFS 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "EXIT"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "COPY"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtDestin 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2895
   End
   Begin MSComctlLib.ProgressBar ProBar 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar TotBar 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "DESINTATION:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "SOURCE:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   990
   End
   Begin VB.Label lblPro 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WAITING"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Label lblSPer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblTPer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "FILE PROGRESS:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "TOTAL PROGRESS:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   2025
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FolderArray() As String
Dim FileArray() As String
Dim TotalSize As Double
Dim CurrByte As Double
Dim bCancel As Boolean

Private Sub cmdCopy_Click()
    If txtSource.Text <> "" And txtDestin.Text <> "" Then
        cmdCopy.Enabled = False
        cmdExit.Caption = "CANCEL"
        bCancel = False
        CopyNewFiles txtSource.Text, txtDestin.Text
    End If
End Sub

Private Sub cmdExit_Click()
    If cmdExit.Caption = "EXIT" Then End
    If cmdExit.Caption = "CANCEL" Then
        bCancel = True
        lblSPer.Caption = "0%"
        lblTPer.Caption = "0%"
        ProBar.Value = 0
        TotBar.Value = 0
        lblPro.Caption = "CANCELLED"
        cmdCopy.Enabled = True
        cmdExit.Caption = "EXIT"
    End If
End Sub

Private Sub cmdFD_Click()
    Form2.Caption = "DESTINATION FOLDER:"
    Form2.Show
End Sub

Private Sub cmdFS_Click()
    Form2.Caption = "SOURCE FOLDER:"
    Form2.Show
End Sub

Private Sub Form_Load()
    lblSPer.Caption = "0%"
    lblTPer.Caption = "0%"
End Sub

Private Function NewFile(FilePath As String) As Boolean
    Dim FileDate As String
    Dim FileDay As Integer
    Dim FileMonth As Integer
    Dim FileYear As Integer
    FileDate = FileDateTime(FilePath)
    FileDay = Format(FileDate, "dd")
    FileMonth = Format(FileDate, "mm")
    FileYear = Format(FileDate, "yyyy")
    
    NewFile = False
    If FileDay = Day(Date) Then
        If FileMonth = Month(Date) Then
            If FileYear = Year(Date) Then
                NewFile = True
            End If
        End If
    End If
End Function

Private Function CopyNewFiles(RootFolder As String, Destination As String)
    On Error Resume Next
    Dim FSO, rObject, Drive
    Dim x As Variant
    Dim i, j, Count As Double
    Set FSO = CreateObject("Scripting.FileSystemObject")
    lblPro.Caption = "INITIALIZING..."
    If Len(RootFolder) > 3 Then
        Set rObject = FSO.GetFolder(RootFolder)
    Else
        Set Drive = FSO.Drives(Mid(RootFolder, 1, 1))
        Set rObject = Drive.RootFolder
        Destination = Destination & "\"
    End If
    ReDim FolderArray(1)
    ReDim FileArray(1)
    FolderArray(0) = RootFolder
    Count = 1
    For Each x In rObject.SubFolders
        If bCancel = True Then Exit Function
        ReDim Preserve FolderArray(UBound(FolderArray) + 1)
        FolderArray(Count) = x.Path
        Count = Count + 1
    Next x
    i = 1
    Do Until i > UBound(FolderArray) - 1
        DoEvents
        Set rObject = FSO.GetFolder(FolderArray(i))
        For Each x In rObject.SubFolders
            If bCancel = True Then Exit Function
            ReDim Preserve FolderArray(UBound(FolderArray) + 1)
            FolderArray(Count) = x.Path
            Count = Count + 1
        Next x
        i = i + 1
    Loop
    Count = 0
    For i = 0 To UBound(FolderArray) - 1
        DoEvents
        Set rObject = FSO.GetFolder(FolderArray(i))
        For Each x In rObject.Files
            DoEvents
            If bCancel = True Then Exit Function
            If UCase(x.Name) <> "PAGEFILE.SYS" Then
                If NewFile(x.Path) = True Then
                    ReDim Preserve FileArray(UBound(FileArray) + 1)
                    FileArray(Count) = x.Path
                    TotalSize = TotalSize + FileLen(FileArray(Count))
                    Count = Count + 1
                End If
            End If
        Next x
    Next i
    For i = 0 To UBound(FileArray) - 1
        DoEvents
        If bCancel = True Then Exit Function
        CreatePath Destination & Mid(FileArray(i), Len(RootFolder) + 1, Len(FileArray(i)) - Len(RootFolder))
        lblPro.Caption = FileArray(i)
        CopyFile FileArray(i), Destination & Mid(FileArray(i), Len(RootFolder) + 1, Len(FileArray(i)) - Len(RootFolder))
    Next i
    lblSPer.Caption = "0%"
    lblTPer.Caption = "0%"
    ProBar.Value = 0
    TotBar.Value = 0
    lblPro.Caption = "COMPLETED - " & UBound(FileArray) - 1 & " COPIED"
    cmdCopy.Enabled = True
    cmdExit.Caption = "EXIT"
End Function

Private Function CopyFile(Source As String, Destin As String)
    On Error Resume Next
    Dim SrcFile As String
    Dim DestFile As String
    Dim SrcFileLen As Long
    Dim nSF, nDF As Integer
    Dim Chunk As String
    Dim BytesToGet As Integer
    Dim BytesCopied As Long
    Dim TotalPer As Integer
    SrcFile = Source
    DestFile = Destin
    SrcFileLen = FileLen(SrcFile)
    nSF = 1
    nDF = 2
    Open SrcFile For Binary As nSF
    Open DestFile For Binary As nDF
    BytesToGet = 10240 '20kb
    BytesCopied = 0
    ProBar.Value = 0
    lblSPer.Caption = "0%"
    Do While BytesCopied < SrcFileLen
        DoEvents
        If BytesToGet < (SrcFileLen - BytesCopied) Then
            Chunk = Space(BytesToGet)
            Get #nSF, , Chunk
        Else
            Chunk = Space(SrcFileLen - BytesCopied)
            Get #nSF, , Chunk
        End If
        BytesCopied = BytesCopied + Len(Chunk)
        ProBar.Value = Int(BytesCopied / SrcFileLen * 100)
        lblSPer.Caption = ProBar.Value & "%"
        TotalPer = Int(((BytesCopied + CurrByte) / TotalSize) * 100)
        lblTPer.Caption = TotalPer & "%"
        TotBar.Value = TotalPer
        Put #nDF, , Chunk
        Chunk = ""
    Loop
    Close #nSF
    Close #nDF
    CurrByte = CurrByte + BytesCopied
End Function

Private Function PathExists(Path As String) As Boolean
    On Error GoTo MakeF
    Open Path & "\Temp.$$$" For Output As #1
    Close #1
    Kill Path & "\Temp.$$$"
    PathExists = True
    Exit Function
MakeF:
    PathExists = False
End Function

Private Function CreatePath(Path As String)
    Dim ArrayPath() As String
    Dim z As Integer
    Dim FullPath As String
    ArrayPath = Split(Path, "\")
    For z = 0 To UBound(ArrayPath) - 1
        FullPath = FullPath & ArrayPath(z)
        If PathExists(FullPath) = False Then MkDir FullPath
        FullPath = FullPath & "\"
    Next
End Function


