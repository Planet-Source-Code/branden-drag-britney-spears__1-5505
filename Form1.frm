VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Drag Britney Spears....       By Branden"
   ClientHeight    =   6720
   ClientLeft      =   1170
   ClientTop       =   1875
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6720
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   3480
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   6435
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   1
      Top             =   7080
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DID you know  that Britney Spears favorite
'basketball team is the bulls?
Dim xBritneySpears As Single
Dim yBritneySpears As Single
Private Sub Command1_Click()
'Britney was born on December 2
Unload Me
End Sub
Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
'her favorite icecream is CoOkie Doe!!
Source.Move x - xBritneySpears, y - yBritneySpears
End Sub
Private Sub Picture1_DragDrop(Source As Control, x As Single, y As Single)
'she likes long walks on the beach
Source.Move Picture1.Left + x - xBritneySpears, Picture1.Top + y - yBritneySpears
End Sub
Private Sub Picture1_MouseDown(ButtBritneySpears As Integer, Shift As Integer, x As Single, y As Single)
'britney was on her highschool
'basketball team, and often goes
'to church
xBritneySpears = x
yBritneySpears = y
Picture1.Drag 1
End Sub
