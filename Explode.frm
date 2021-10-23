VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "This is a test!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem define type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Rem declare api calls
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


Private Sub Explode(Newform As Form, Increment As Integer)

Dim Size As RECT                                                                                                ' setup form as rect type
GetWindowRect Newform.hwnd, Size

Dim FormWidth, FormHeight As Integer                                                                 ' establish dimension variables
FormWidth = (Size.Right - Size.Left)
FormHeight = (Size.Bottom - Size.Top)

Dim TempDC
TempDC = GetDC(ByVal 0&)                                                                                 ' obtain memory dc for resizing

Dim Count, LeftPoint, TopPoint, nWidth, nHeight As Integer                                      ' establish resizing variables
For Count = 1 To Increment                                                                               ' loop to new sizes
    nWidth = FormWidth * (Count / Increment)
    nHeight = FormHeight * (Count / Increment)
    LeftPoint = Size.Left + (FormWidth - nWidth) / 2
    TopPoint = Size.Top + (FormHeight - nHeight) / 2
    Rectangle TempDC, LeftPoint, TopPoint, LeftPoint + nWidth, TopPoint + nHeight     ' draw rectangles to build form
Next Count

DeleteDC (TempDC)                                                                                           ' release  memory resource

End Sub


Private Sub Form_Load()

Explode Me, 1000                                                                                             ' open this form by number of desired increment

End Sub

