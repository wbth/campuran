VERSION 5.00
Begin VB.UserControl Skin 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   ControlContainer=   -1  'True
   DataSourceBehavior=   1  'vbDataSource
   EditAtDesignTime=   -1  'True
   HasDC           =   0   'False
   PropertyPages   =   "Skin.ctx":0000
   ScaleHeight     =   2730
   ScaleWidth      =   8850
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title Bar"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   90
         Width           =   585
      End
   End
   Begin VB.PictureBox MenuBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   8760
      TabIndex        =   3
      Top             =   315
      Visible         =   0   'False
      Width           =   8790
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.PictureBox Skinpic 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   2100
      Picture         =   "Skin.ctx":000F
      ScaleHeight     =   795
      ScaleWidth      =   3135
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1410
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   0
      Begin VB.Menu a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   1
      Begin VB.Menu b 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   2
      Begin VB.Menu c 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   3
      Begin VB.Menu d 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   4
      Begin VB.Menu e 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   5
      Begin VB.Menu f 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   6
      Begin VB.Menu g 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu Core 
      Caption         =   "Core"
      Index           =   7
      Begin VB.Menu h 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "Skin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"

Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Type ButtonLocation
    Close_Left As Integer
    Close_Right As Integer
    Min_Left As Integer
    Min_Right As Integer
    Max_Left As Integer
    Max_Right As Integer
End Type

Private skinned As Boolean
Private skinfile As String
Private Bool_Min As Boolean
Private Bool_Max As Boolean
Private Bool_Remember  As Boolean
Private Bool_Menu As Boolean

Private frm As Form

Private initok As Boolean 'init check variable
Private locate As ButtonLocation

'Public Evenets
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MenuClick(ItemName As String)
Private inrgn As Boolean
Private MenuNameArray() As String
Private MenuCaptionArray() As String
Private rounded As Boolean
Private a1() As String, b1() As String, c1() As String, d1() As String, e1() As String, f1() As String, g1() As String, h1() As String



Private Sub InitAllocate() 'avoiding some redundant code from allocating effort to increase speed
    frm.BorderStyle = 0
    TitleBar.Top = 0
    TitleBar.Left = 0
    CapLabel.Top = 75
    CapLabel.Caption = frm.Caption
    UserControl.Height = frm.Height
End Sub

Public Sub allocate() ' allocate locations of various buttons depending on settings
    UserControl.Width = frm.Width
    UserControl.Height = frm.Height
    initok = True
End Sub


Private Sub DrawTitleBar() 'Draws Title Bar, iteration avoided !

 TitleBar.PaintPicture Skinpic.Picture, 0, 0, TitleBar.ScaleWidth, 375, 300, 210, 270, 435
 If rounded = True Then TitleBar.PaintPicture Skinpic.Picture, 0, 0, 235, 375, 0, 210, 270, 435


End Sub
Private Sub a_Click(index As Integer)
    RaiseEvent MenuClick(a1(index))
End Sub
Private Sub b_Click(index As Integer)
    RaiseEvent MenuClick(b1(index))
End Sub
Private Sub c_Click(index As Integer)
    RaiseEvent MenuClick(c1(index))
End Sub
Private Sub d_Click(index As Integer)
    RaiseEvent MenuClick(d1(index))
End Sub
Private Sub e_Click(index As Integer)
    RaiseEvent MenuClick(e1(index))
End Sub
Private Sub f_Click(index As Integer)
    RaiseEvent MenuClick(f1(index))
End Sub
Private Sub g_Click(index As Integer)
    RaiseEvent MenuClick(g1(index))
End Sub
Private Sub h_Click(index As Integer)
    RaiseEvent MenuClick(h1(index))
End Sub


Private Sub Label_Click(index As Integer)
    PopupMenu Core(index), , Label(index).Left, MenuBar.Top + MenuBar.Height
End Sub

Private Sub TitleBar_DblClick()
If Bool_Max = True Then
    If frm.WindowState = 2 Then
        frm.WindowState = 0
        
    Else
        frm.WindowState = 2
    End If
End If
End Sub


Private Sub UserControl_Initialize()
    TitleBar.Width = UserControl.Width
    LoadMenu
End Sub

' Various propertis of skinner
Public Property Let MenuVisible(Temp As Boolean)
    Bool_Menu = Temp
    MenuBar.Visible = Temp
    
End Property

Public Property Let Caption(Temp As String)
    CapLabel = Temp
End Property
Public Property Get MenuVisible() As Boolean
    MenuVisible = Bool_Menu
End Property


Public Property Get Caption() As String
    Caption = CapLabel.Caption
End Property



Public Sub Init_Skin(frm1 As Form)
    initok = True
    Set frm = frm1
    Dim filename As String
    filename = GetSetting(frm.Caption, "skin", "main", "")
    InitAllocate
    allocate
    'If Me.RememberSkin = True And Not filename = "" Then ChangeSkin filename
    
End Sub

Private Sub LoadMenu()

If Dir(App.path + "\menu.txt") = "" Then Exit Sub
Dim Temp As String, Temp1 As String
Dim i As Integer
Dim level As Integer
Dim ctr As Integer

Open App.path + "\menu.txt" For Input As #1
While Not EOF(1)
    Line Input #1, Temp
    On Error Resume Next
    Temp1 = Mid(Temp, 1, InStr(Temp, "!") - 1)
    Temp = Mid(Temp, InStr(Temp, "!") + 1)
    level = GetLevel(Temp)
    If level = 0 Then
        i = 0
        If Not ctr = 0 Then
            Load Label(ctr)
            Label(ctr).Left = Label(ctr - 1).Left + Label(ctr - 1).Width + 200
            Label(ctr).Caption = Temp
            Label(ctr).Visible = True
            ctr = ctr + 1
        Else
            Label(ctr).Caption = Temp
            ctr = ctr + 1
        End If
    Else
    Temp = Trimed(Temp)
        Select Case (ctr - 1)
    Case 0:
            Load a(i)
            a(i).Caption = Temp
            ReDim Preserve a1(i)
            a1(i) = Temp1
    Case 1:
            Load b(i)
            b(i).Caption = Temp
            ReDim Preserve b1(i)
            b1(i) = Temp1
    Case 2:
            Load c(i)
            c(i).Caption = Temp
            ReDim Preserve c1(i)
            c1(i) = Temp1
    Case 3:
            Load d(i)
            d(i).Caption = Temp
            ReDim Preserve d1(i)
            d1(i) = Temp1
    Case 4:
            Load e(i)
            e(i).Caption = Temp
            ReDim Preserve e1(i)
            e1(i) = Temp1
    Case 5:
            Load f(i)
            f(i).Caption = Temp
            ReDim Preserve f1(i)
            f1(i) = Temp1
    Case 6:
            Load g(i)
            g(i).Caption = Temp
            ReDim Preserve g1(i)
            g1(i) = Temp1
    Case 7:
            Load h(i)
            h(i).Caption = Temp
            ReDim Preserve h1(i)
            h1(i) = Temp1
    End Select
    i = i + 1
    End If
    
    
Wend
Close #1
End Sub
Private Function Trimed(Temp As String) As String
again:
    If InStr(Temp, "....") Then
        Temp = Mid(Temp, 5)
        GoTo again
    End If
    Trimed = Temp
End Function


Private Function GetLevel(ByVal Temp As String) As Integer
    Dim i As Integer
again:
    If InStr(Temp, "....") Then
        Temp = Mid(Temp, 5)
        i = i + 1
        GoTo again
    End If
    GetLevel = i

End Function

