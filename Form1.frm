VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Info Grabber"
   ClientHeight    =   5415
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ClassNametxt 
      BackColor       =   &H80000006&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   25
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox DCtxt 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   21
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox childtxt 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   19
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox hInstancetxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox ProcIDtxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox UserDatatxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Styletxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox IDtxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox ExStyletxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Parenttxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Classtxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox hWndtxt 
      BackColor       =   &H80000007&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   5280
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Name:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Pixel:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Dc:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Child:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "hInstance:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Proc ID:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "UserData:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Style:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ExStyle:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Parent:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Handle:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuMinWindow 
         Caption         =   "Minimize Window"
      End
      Begin VB.Menu mnuMaxWindow 
         Caption         =   "Maximize Window"
      End
      Begin VB.Menu mnuHideWindow 
         Caption         =   "Hide Window"
      End
      Begin VB.Menu mnuShowWindow 
         Caption         =   "ShowWindow"
      End
      Begin VB.Menu mnuwinStayOnTop 
         Caption         =   "Window StayOnTop"
      End
      Begin VB.Menu mnuSetWinPos 
         Caption         =   "Set Window Position"
      End
      Begin VB.Menu mnuChangeWindowTitle 
         Caption         =   "Change Window Title"
      End
      Begin VB.Menu mnuSuperBannerHider 
         Caption         =   "Super Banner Hider"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label10_Click()
On Error Resume Next
Dim TTime As Integer
TTime = InputBox("Enter In How Many Secs You Want To Wait To Get The Window Information", "Timer")
Timer1.Interval = TTime & "000"
Timer1.Enabled = True
End Sub

Private Sub mnuAbout_Click()
MsgBox "This Program Was Coded By: Catdaddy187", vbOKOnly, "About"
End Sub

Private Sub mnuChangeWindowTitle_Click()
Dim ChangeWindowTitle As String
ChangeWindowTitle = InputBox("Enter in a new title name", "New Title")
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuHideWindow_Click()
ShowWindow hWndtxt.Text, SW_HIDE
End Sub

Private Sub mnuMaxWindow_Click()
ShowWindow hWndtxt.Text, 3
End Sub

Private Sub mnuMinWindow_Click()
ShowWindow hWndtxt.Text, SW_MINIMIZE
End Sub

Private Sub mnuSetWinPos_Click()
Dim XPos As Integer
Dim YPos As Integer
XPos = InputBox("Set Window X Axis", "SetWindowPos")
YPos = InputBox("Set Window Y Axis", "SetWindowPos")
SetWindowPos hWndtxt.Text, 0, XPos, YPos, 0, 0, SWP_NOSIZE
End Sub

Private Sub mnuShowWindow_Click()
ShowWindow hWndtxt.Text, SW_NORMAL
End Sub

Private Sub mnuSuperBannerHider_Click()
SetWindowPos hWndtxt.Text, HWND_TOPMOST, 0, 0, 0, 0, 0
End Sub

Private Sub mnuwinStayOnTop_Click()
SetWindowPos hWndtxt.Text, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim P As POINTAPI
Dim lpClassName As String * 255
Dim nMaxCount As Integer
Dim lpString As String * 244
Dim cch As Integer
Dim rc As Long
nMaxCount = 255
cch = 255
GetCursorPos P
hWndtxt.Text = WindowFromPoint(P.x, P.y)
GetWindowText hWndtxt, lpString, cch
Classtxt.Text = lpString
Parenttxt.Text = GetParent(hWndtxt.Text)
ExStyletxt.Text = GetWindowLong(hWndtxt.Text, GWL_EXSTYLE)
IDtxt.Text = GetWindowLong(hWndtxt.Text, GWL_ID)
Styletxt.Text = GetWindowLong(hWndtxt.Text, GWL_STYLE)
UserDatatxt.Text = GetWindowLong(hWndtxt.Text, GWL_USERDATA)
ProcIDtxt.Text = GetWindowLong(hWndtxt.Text, GWL_WNDPROC)
hInstancetxt.Text = GetWindowLong(hWndtxt.Text, GWL_HINSTANCE)
childtxt.Text = ChildWindowFromPoint(hWndtxt.Text, P.x, P.y)
rc = GetDC(hWndtxt.Text)
DCtxt.Text = rc
Label14.Visible = True
Label14.BackColor = GetPixel(rc, x / 15, y / 15)
rc = GetClassName(hWndtxt.Text, lpClassName, nMaxCount)
ClassNametxt.Text = lpClassName
Text2.Text = GetClassWord(hWndtxt.Text, 255)
Timer1.Enabled = False
End Sub
