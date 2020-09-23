VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hidden exe trojen scanner in vb projects"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdScan 
      Caption         =   "Scan"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox TrojenList 
      Height          =   3180
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please specify a location for the project files"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.FileListBox Filelist 
         Height          =   1065
         Left            =   1200
         Pattern         =   "*.frm;*.bas;*.cls;*.res;*.ctl;*.pag;*.dsr;*.frx;*.vbw;*.vbp"
         TabIndex        =   5
         Top             =   2880
         Width           =   3135
      End
      Begin VB.DirListBox Dirlist 
         Height          =   2115
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   3135
      End
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project files:"
         Height          =   195
         Left            =   1200
         TabIndex        =   6
         Top             =   2640
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Possible trojens in projects"
      Height          =   195
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastDrive As String
Dim FileNum As Integer
Dim FileBinary As String

Private Sub CmdScan_Click()

    If Filelist.ListCount = 0 Then MsgBox "Their is no project files in the folder you specified.", vbQuestion, "Project scanner": Exit Sub

    

    For M = 0 To Filelist.ListCount - 1
        Filelist.Selected(M) = True
        FileNum = FreeFile
        If Right(Dirlist, 1) = "\" Or Right(Dirlist, 1) = "/" Then
            Open Dirlist.Path & Filelist.List(M) For Binary As #FileNum
        Else
            Open Dirlist.Path & "\" & Filelist.List(M) For Binary As #FileNum
        End If
            FileBinary = Space(2)
            Get #FileNum, 1, FileBinary
            If LCase(FileBinary) = LCase("MZ") Then TrojenList.AddItem (Filelist.List(M))
        Close #FileNum
    Next M
    
    If TrojenList.ListCount = 0 Then
        MsgBox "Files scanning complete. Their was no possible exe trojens detected in source files.", vbOKOnly, "Scan complete"
    Else
        MsgBox "Possible trojens were found in project files. You can view these in the possible trojen list.", vbCritical, "Warning!"
    End If

End Sub

Private Sub Dirlist_Change()
    Filelist.Path = Dirlist.Path
End Sub

Private Sub Drive_Change()
    On Error GoTo FinaliseError
    Dirlist.Path = Drive.Drive
    Filelist.Path = Dirlist.Path
    LastDrive = Drive.Drive
    Exit Sub
FinaliseError:
    MsgBox "Error, drive not ready.", vbCritical, "Drive not ready"
    Drive.Drive = LastDrive
End Sub
