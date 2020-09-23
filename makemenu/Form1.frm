VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Munky's menu Writer"
   ClientHeight    =   2235
   ClientLeft      =   3795
   ClientTop       =   3345
   ClientWidth     =   3405
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "The Submenu within the menu"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Caption"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "The caption of the menu"
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "notepad"
      ToolTipText     =   "This is the class of the window you want to add the menu to"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Menu Caption:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Caption:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim FndWnd As String
FndWnd = FindWindow(Text1.Text, vbNullString)
CreateMenu FndWnd, Text2.Text, Text3.Text

End Sub

Private Sub Form_Load()
App.Title = "Menu Creator"
Form1.Picture = LoadPicture(App.path & "\mnucr.bmp")
windowontop Form1.hwnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
WindowDrag Me.hwnd
End Sub

Private Sub Label4_Click()
End
End Sub
