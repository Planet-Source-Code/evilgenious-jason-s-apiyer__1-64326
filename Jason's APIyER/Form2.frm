VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1920
      Top             =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "About!"
      BeginProperty Font 
         Name            =   "MeraKhutt2"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2220
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   2235
      Left            =   3600
      Picture         =   "Form2.frx":000C
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2235
      Left            =   0
      Picture         =   "Form2.frx":8E52
      Top             =   0
      Width           =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Is All About"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1990
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SECRET"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()

If Label1.ForeColor = vbYellow Then
    Label1.ForeColor = vbGreen
ElseIf Label1.ForeColor = vbGreen Then
    Label1.ForeColor = vbYellow
End If

End Sub
