VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form viewer 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Tugas Multimedia"
   ClientHeight    =   5220
   ClientLeft      =   2565
   ClientTop       =   390
   ClientWidth     =   12045
   Icon            =   "viewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   7800
      Top             =   10740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.CDR Files | *.cdr"
   End
   Begin multimedia.Skin Skin1 
      Align           =   1  'Align Top
      Height          =   5235
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   9234
      Begin MSComctlLib.TreeView tree 
         Height          =   4185
         Left            =   300
         TabIndex        =   4
         Top             =   870
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   7382
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         ImageList       =   "Images"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView filelist 
         Height          =   4185
         Left            =   4140
         TabIndex        =   3
         Top             =   870
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   7382
         View            =   1
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "files"
         SmallIcons      =   "files"
         ColHdrIcons     =   "files"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.ListBox Mainlist 
      Height          =   255
      Left            =   4560
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.ListBox added 
      Height          =   255
      Left            =   5220
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   510
      Visible         =   0   'False
      Width           =   2685
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   4350
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":08CA
            Key             =   "addexistingitem"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":0E64
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":13FE
            Key             =   "ViewOpen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":1998
            Key             =   "PageSetup"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":1F32
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":24CC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":2A66
            Key             =   "Processes"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":3000
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":359A
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":3B34
            Key             =   "ReplaceInFiles"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":40CE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":4668
            Key             =   "SaveAll"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":4C02
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":519C
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":5736
            Key             =   "StartNoDebug"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":5CD0
            Key             =   "opensolution"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":626A
            Key             =   "OpenFile"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":6804
            Key             =   "FindInFiles"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":6D9E
            Key             =   "FindSymbol"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":7338
            Key             =   "FullScreen"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":78D2
            Key             =   "Index"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":7E6C
            Key             =   "Inmediate"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":8406
            Key             =   "NewFile"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":89A0
            Key             =   "NewProject"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":8F3A
            Key             =   "OpenProjectFromWeb"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":94D4
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":9A6E
            Key             =   "Designer"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":A008
            Key             =   "DynamicHelp"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":A5A2
            Key             =   "Exceptions"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":AB3C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":B0D6
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":B670
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":BC0A
            Key             =   "AddNewItem"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":C1A4
            Key             =   "BlankSolution"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":C73E
            Key             =   "BreakPoints"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":CCD8
            Key             =   "closesolution"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":D272
            Key             =   "Code"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":D80C
            Key             =   "Contents"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   -30
      Top             =   6810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":DDA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":E680
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":EF5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":F4F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList files 
      Left            =   420
      Top             =   6780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":FDCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":10368
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":10682
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":1099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":10CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":11590
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":118AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":11BC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":11EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":121F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":12512
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":12AAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSkins 
      Left            =   -210
      Top             =   8550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   207
      ImageHeight     =   52
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":13386
            Key             =   "Blue (default)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":1B298
            Key             =   "Ghost"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":231AA
            Key             =   "Noir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":2B0BC
            Key             =   "Simile XP Blue"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":32FCE
            Key             =   "Cyan Neon"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":3AEE0
            Key             =   "Digital"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "viewer.frx":42DF2
            Key             =   "BlueSteel"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type NodeInfo
    Nodename As String
End Type

Dim NodeCount As Long
Dim Node() As NodeInfo
Dim laststr As String

Private Sub Creater_Click()
    GetList_Click
End Sub

Private Sub Form_Load()
Skin1.Init_Skin Me
Skin1.MenuVisible = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
filelist.Width = Me.Width - filelist.Left - 300
Skin1.allocate
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload main
    
End Sub

Private Sub GetList_Click()
main.Show vbModal
End Sub


Private Sub Mainlist_Click()
filelist.ListItems.Clear
Dim i As Integer
Dim j As Integer
Dim Temp As String
j = Mainlist.ListIndex


If j > -1 And j < Mainlist.ListCount And Not cd.filename = "" Then
    Open cd.filename For Input As #1
        While Not EOF(1)
        Input #1, Temp
        If Temp = "@" Then
            i = i + 1
            If i = j + 1 Then
                Input #1, Temp
                GoTo found
            End If
        End If
        Wend
        
        GoTo finish
found:
    
    While Not EOF(1)
        Input #1, Temp
        If Not Temp = "@" Then
            filelist.ListItems.Add , , Temp, GetDeviceIcon(Temp), GetDeviceIcon(Temp)
        Else
            GoTo finish
        End If
    Wend
finish: Close #1
    
End If
End Sub

Private Function GetDeviceIcon(tmp As String) As Integer
    tmp = LCase(tmp)
    Dim tmp1 As String
    tmp1 = LCase(Mid(tmp, Len(tmp) - 2))
    Select Case tmp1
        Case "exe":
                    If InStr(tmp, "setup") Then
                        GetDeviceIcon = 7
                    Else
                        GetDeviceIcon = 1
                    End If
        Case "hlp", "hlp": GetDeviceIcon = 2
        Case "frm": GetDeviceIcon = 4
        Case "vbp": GetDeviceIcon = 9
        Case "mpg", "mpeg", "dat", "divx", "avi": GetDeviceIcon = 6
        Case "mp3", "midi", "wav", "wma", "cda": GetDeviceIcon = 5
        Case "txt": GetDeviceIcon = 8
        Case "doc": GetDeviceIcon = 10
        Case "psd", "jpg", "jpeg", "bmp": GetDeviceIcon = 11
        Case "zip", "rar", "ace": GetDeviceIcon = 12
        Case Else: GetDeviceIcon = 3
    End Select
    
    
End Function

Private Sub mnu_Open_Click()
 open_Click
End Sub


Private Sub open_Click()

cd.ShowOpen
DisplayFile cd.filename
    
    
End Sub

Public Sub DisplayFile(filename As String)
Mainlist.Clear
added.Clear
tree.Nodes.Clear
filelist.ListItems.Clear
Dim Temp As String
Dim i As Integer
    If Not filename = "" Then
        Open filename For Input As #1
            Mainlist.Clear
            While Not EOF(1)
                Input #1, Temp
                If Temp = "@" Then
                    Input #1, Temp
                    Mainlist.AddItem Temp
                    DoEvents
                End If
            Wend
        Close #1
    End If
Dim maxslash As Long

Dim fail As Boolean
maxslash = GetMaxSlashes()

Dim k As Long
Dim j As Long

'------------------------------------------------------------------
                


For k = 1 To maxslash + 1
    For j = 0 To Mainlist.ListCount - 1
            If GetNthSlashLocation(Mainlist.List(j), k, Temp) = True Then
            For i = 0 To added.ListCount - 1
                If added.List(i) = Temp Then GoTo done1
            Next i
                
                added.AddItem Temp
done1:
        End If
    Next j

Next k

'-----------------------------------------------------------------

tree.Tag = added.List(0) + "\"

'-----------Add the master Tree Content--------------------------
tree.Nodes.Add , , , "Drive", 1, 2
For i = 0 To added.ListCount - 1
    If GetSlashes(added.List(i)) = 1 Then
        tree.Nodes.Add 1, tvwChild, , Mid(added.List(i), 4), 3, 4
        
    End If
Next i
DoEvents
Dim lastcount As Long
'-----------------------Add sub tree---------------------------
For k = 2 To maxslash + 1
    For i = 0 To added.ListCount - 1
        If GetSlashes(added.List(i)) = k Then
            GetNthSlashLocation Mid(added.List(i), 4), k - 1, Temp
            Temp = GetLast(Temp)
            
            For j = 1 To tree.Nodes.Count - 1
                If tree.Nodes.Item(j) = Temp Then GoTo intree
                DoEvents
            Next j
intree:
                tree.Nodes.Add j, tvwChild, , GetLast(added.List(i)), 3, 4
                
        End If
    Next i
Next k
End Sub
Private Function GetLast(ByVal tmp As String) As String
    GetLast = Mid(tmp, InStrRev(tmp, "\") + 1)
   
End Function

Private Function GetSlashes(Str As String) As Long
Dim i As Long
Dim ctr As Long
For i = 1 To Len(Str)
    If Mid(Str, i, 1) = "\" Then ctr = ctr + 1
    
Next i
GetSlashes = ctr
End Function
Private Function GetMaxSlashes() As Long
Dim i As Long
Dim ctr As Long
ctr = 0
ctr1 = 0
For i = 0 To Mainlist.ListCount - 1
    ctr1 = GetSlashes(Mainlist.List(i))
    If ctr < ctr1 Then ctr = ctr1
Next i
GetMaxSlashes = ctr

End Function

Private Function GetNthSlashLocation(Str As String, Loc As Long, Optional ByRef RetString As String) As Boolean
Dim i As Long
Dim ctr As Long

For i = 1 To Len(Str)
    If Mid(Str, i, 1) = "\" Then ctr = ctr + 1
    If ctr = Loc Then GoTo success
Next i

If ctr = Loc - 1 Then
    GetNthSlashLocation = True
    RetString = Str

Else
GetNthSlashLocation = False
End If
Exit Function
success:
GetNthSlashLocation = True
RetString = Mid(Str, 1, i - 1)

End Function

Private Sub Report_Click()
filelist.View = lvwReport
End Sub

Private Sub search_Click()
Dim i As Integer
Dim Temp As String
If cd.filename = "" Then cd.ShowOpen
    If Not cd.filename = "" Then
        Open cd.filename For Input As #1
            While Not EOF(1)
                Input #1, Temp
                If Temp = "@" Then
                    i = i + 1
                
                End If
            Wend
        Close #1
    End If
End Sub


Private Sub Skinner1_BeforeShowChangeSkinDialog(FormName As String, Cancel As Boolean)
    Cancel = True
    frmChangeSkin.Show 1
End Sub
Private Sub Skin1_MenuClick(ItemName As String)
    Select Case ItemName
    Case "Open":
            open_Click
    Case "Save":
            main.Show vbModal

    Case "Exit":
            Unload Me
    End Select

End Sub

Private Sub tree_DBlClick()
    On Error Resume Next
    Dim i As Integer
    Dim Temp As String
    Temp = tree.Tag & Mid(tree.SelectedItem.FullPath, 7)
    For i = 0 To Mainlist.ListCount - 1
    If Temp = Mainlist.List(i) Then
        Mainlist.Selected(i) = True
        Mainlist_Click
        Exit Sub
    End If
    Next i
    Mainlist.Selected(0) = True
    Mainlist_Click
End Sub
