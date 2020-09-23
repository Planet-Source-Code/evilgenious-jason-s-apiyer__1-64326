VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00DDE6E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jason's APIyER"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picColor 
      Height          =   470
      Left            =   5160
      ScaleHeight     =   405
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   7995
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   2880
      TabIndex        =   31
      Top             =   8160
      Width           =   2175
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   2880
      TabIndex        =   29
      Top             =   6480
      Width           =   2895
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2880
      TabIndex        =   27
      Top             =   7320
      Width           =   2895
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   2880
      TabIndex        =   24
      Top             =   7680
      Width           =   2895
   End
   Begin VB.TextBox Text11 
      Height          =   735
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      Height          =   735
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDE6E2&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Left            =   0
      TabIndex        =   10
      Top             =   -120
      Width           =   6015
      Begin VB.CommandButton Command4 
         BackColor       =   &H00DDE6E2&
         Caption         =   "About!"
         Height          =   300
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DDE6E2&
         Caption         =   "Top Window"
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picHandler 
         AutoSize        =   -1  'True
         BackColor       =   &H00DDE6E2&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5280
         Picture         =   "Form1.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   250
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   1005
         Left            =   0
         Picture         =   "Form1.frx":1604
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Drag it"
         Height          =   195
         Left            =   5310
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D5B98A&
         BackStyle       =   0  'Transparent
         Caption         =   "A Small tool to play with the OS handle"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   6120
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00DDE6E2&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1755
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DDE6E2&
         Caption         =   "Enable"
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2760
         TabIndex        =   37
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2760
         TabIndex        =   36
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00DDE6E2&
         Caption         =   "Disable"
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Text            =   "08-4C23-EEF"
         Top             =   460
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00DDE6E2&
         Caption         =   "Set Text"
         Height          =   255
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   80
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Handle of Control to enable it:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Handle of Control to disable it:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Handle of Control to Set the Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   4680
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Color from Pointer:"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "DeviceContext:"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Window Handle:"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cursor Position:"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   7710
      Width           =   1335
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Window Text:"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Window Text:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Process ID:"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   5760
      Width           =   825
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Window Handle:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class name:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Handle of Control from Pointer:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   2250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Handle of Control from Draging:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   6840
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Name: Jason's APIyER
' Date: 2005 Dec 27

' BULLSHITS

' (****************************************************************)
' This application I have made just for educational purpose only
' -----------------------------------------------------------------
' Special thanks to Jason and Freddy for their ugly pictures
' (****************************************************************)

' API DECLARATIONS
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long

' TYPES AND VARIABLES
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Const MAX_PATH = 260
Const WM_GETTEXTLENGTH = &HE
Const WM_GETTEXT = &HD
Const WM_SETTEXT = &HC

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private CursorLoc As POINTAPI

Dim intRed, intGreen, intBlue As Byte
Dim colorFromPoint As Long





' ENABLE WINDOW WITH ITS HANDLE
Private Sub Command1_Click()
    EnableWindow Text2.Text, True
End Sub
' DISABLE WINDOW WITH ITS HANDLE
Private Sub Command2_Click()
    EnableWindow Text3.Text, False
End Sub
' SEND TEXT MESSAGE TO WINDOW WITH ITS HANDLE
Private Sub Command3_Click()
    Call SendMessageByString(Text4.Text, WM_SETTEXT, 0&, Text5.Text)
End Sub



' DRAG N DROP OPERATION
Private Sub picHandler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
    Me.MouseIcon = picHandler.Picture
    Text1.Text = vbNullString
End Sub
' DRAG N DROP OPERATION
Private Sub picHandler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 1
    Dim scaption As String
    Dim ihandle As Long
    Dim lstrlen As Long
    
    lstrlen = MAX_PATH
    scaption = Space$(MAX_PATH)
    
    ' get window handle from draging
    Call GetCursorPos(CursorLoc)
    ihandle = WindowFromPoint(CursorLoc.X, CursorLoc.Y)
    Text1.Text = ihandle
    Text2.Text = ihandle
    Text3.Text = ihandle
    Text4.Text = ihandle
End Sub
' START CAPTURING HANDLES
Private Sub Timer1_Timer()
    Dim scaption As String
    Dim ihandle As Long
    Dim phandle As Long
    Dim lstrlen As Long
    Dim sClassName As String * 50
    Dim sHwndText As String * 256
    Dim sParentHwndText As String * 256
    
    lstrlen = MAX_PATH
    scaption = Space$(MAX_PATH)
    
    'get cursor location (X,Y co-ordinates)
    Call GetCursorPos(CursorLoc)
    Text12.Text = "X=" & CursorLoc.X & "   Y=" & CursorLoc.Y
    'get window handle under cursor position
    ihandle = WindowFromPoint(CursorLoc.X, CursorLoc.Y)
    Text6.Text = ihandle
    'get parent window handle of ihandle
    phandle = GetParent(ihandle)
    Text8.Text = phandle
    'get class name
    GetClassName ihandle, sClassName, 50
    Text7.Text = sClassName
    'get ihandle window text
    GetWindowText ihandle, sHwndText, 256
    Text11.Text = sHwndText
    'get phandle window text
    GetWindowText phandle, sParentHwndText, 256
    Text10.Text = sParentHwndText
    'get process id
    GetWindowThreadProcessId ihandle, PROC1
    Text9.Text = PROC1
    'get desktop window handle
    Text13.Text = GetDesktopWindow()
    'get device context of ihandle
    dc = GetWindowDC(CLng(Text13.Text))
    Text14.Text = dc
    
    'get color from mouse pointer
    colorFromPoint = GetPixel(dc, CursorLoc.X, CursorLoc.Y)
        ' Get the red value
        intRed = colorFromPoint Mod 256
        ' Get the green value
        intGreen = ((colorFromPoint And &HFF00) / 256&) Mod 256&
        ' Get the blue value
        intBlue = (colorFromPoint And &HFF0000) / 65536
    Text15.Text = "R=" & intRed & "  G=" & intGreen & "  B=" & intBlue
    picColor.BackColor = RGB(intRed, intGreen, intBlue)
    
    
End Sub






































' SET WINDOW ON TOP
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub Command4_Click()
    Form2.Show
End Sub
