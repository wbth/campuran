VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main 
   BorderStyle     =   0  'None
   Caption         =   "Search and add Device Contents"
   ClientHeight    =   2655
   ClientLeft      =   2160
   ClientTop       =   4290
   ClientWidth     =   4545
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   510
      TabIndex        =   7
      Text            =   "Files Found :"
      Top             =   1500
      Width           =   2565
   End
   Begin VB.TextBox no 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3060
      TabIndex        =   6
      Text            =   "0"
      Top             =   1500
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4710
      Top             =   3810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List2 
      Height          =   2010
      ItemData        =   "main.frx":08CA
      Left            =   7050
      List            =   "main.frx":08CC
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.FileListBox File 
      Height          =   480
      Left            =   3360
      TabIndex        =   3
      Top             =   4620
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Top             =   780
      Width           =   2415
   End
   Begin multimedia.Skin Skin1 
      Align           =   1  'Align Top
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   4683
      Caption         =   "Title Bar"
      Skin            =   "main.frx":08CE
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create List"
         Height          =   315
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   780
         Width           =   1335
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         Height          =   315
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2100
         Width           =   1035
      End
      Begin VB.CommandButton ok 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         Height          =   315
         Left            =   3300
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2100
         Width           =   1035
      End
   End
   Begin VB.Label current 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   1575
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Files Ignored"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2310
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancel As Boolean
Dim globalcounter As Long

Function FindFiles(path As String, SearchStr As String, _
       FileCount As Integer, DirCount As Integer)
      Dim filename As String
      Dim DirName As String
      Dim dirNames() As String
      Dim nDir As Integer
      Dim i As Integer
      Dim olddir As String
      olddir = ""
      Dim atleast As Boolean
      Dim j As Integer
      On Error GoTo sysFileERR
       If Cancel = True Then Exit Function
                
      If Right(path, 1) <> "\" Then path = path & "\"

      nDir = 0
      ReDim dirNames(nDir)
      DirName = Dir(path, vbDirectory Or vbHidden Or vbArchive Or vbReadOnly Or vbSystem)
      Do While Len(DirName) > 0
       
         If (DirName <> ".") And (DirName <> "..") Then
    
            If GetAttr(path & DirName) And vbDirectory Then
               dirNames(nDir) = DirName
               DirCount = DirCount + 1
               nDir = nDir + 1
               ReDim Preserve dirNames(nDir)
               'List2.AddItem path & DirName ' Uncomment to list
            End If
sysFileERRCont:
         End If
         DirName = Dir()
      Loop

  
      filename = Dir(path & SearchStr, vbNormal Or vbHidden Or vbSystem _
      Or vbReadOnly Or vbArchive)
      While Len(filename) <> 0
         FindFiles = FindFiles + FileLen(path & filename)
         FileCount = FileCount + 1

         If Not olddir = path Then
            no.Text = globalcounter
            DoEvents
            If Cancel = True Then Exit Function
            olddir = path
            File.Pattern = "*.*"
            File.path = path
            File.Refresh
 
            Print #1, "@"
            Print #1, Left(path, Len(path) - 1)
            
        End If
    
        Print #1, filename
        globalcounter = globalcounter + 1
        DoEvents
    
      '   List1.AddItem FileName
         List2.AddItem path & filename & vbTab & _
            FileDateTime(path & filename)
         filename = Dir()
      Wend

    
      If nDir > 0 Then
   
         For i = 0 To nDir - 1
           FindFiles = FindFiles + FindFiles(path & dirNames(i) & "\", _
            SearchStr, FileCount, DirCount)
         Next i
      End If

AbortFunction:
      Exit Function
sysFileERR:
      If Right(DirName, 4) = ".sys" Then
        Resume sysFileERRCont
      Else
       
       
        Resume AbortFunction
        
      End If
      
      End Function


Private Sub cmdCancel_Click()
Cancel = True
End Sub

Private Sub cmdsearch_click()

globalcounter = 0
Cancel = False
On Error Resume Next
      Dim SearchPath As String, FindStr As String
      Dim FileSize As Long
      Dim NumFiles As Integer, NumDirs As Integer
      Dim i As Integer
      List2.Clear
      cd.Filter = "CDR file | *.cdr"
      cd.ShowSave
      If Not cd.filename = "" Then
      SearchPath = Left(Drive1.Drive, InStrRev(Drive1.Drive, ":"))
      FindStr = "*.*"
      Open cd.filename For Output As #1
      
      FileSize = FindFiles(SearchPath, FindStr, NumFiles, NumDirs)
      Close #1
      MsgBox "Addition complete !"
      viewer.DisplayFile cd.filename
     
      Unload Me
       Else
       MsgBox "job terminated !"
       End If

End Sub

Private Sub Form_Load()
    Skin1.Init_Skin Me
    Me.Top = viewer.Height / 2 - Me.Height / 2 + viewer.Top
    Me.Left = viewer.Width / 2 - Me.Width / 2 + viewer.Left
End Sub

Private Sub OK_Click()
Unload Me
End Sub
