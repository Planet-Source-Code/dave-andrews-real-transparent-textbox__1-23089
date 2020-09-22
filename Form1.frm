VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   262
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMakeTransparent 
      Caption         =   "GO!"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1215
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":9BD6
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DoIt As Boolean
Private Const DT_EDITCONTROL = &H2000&
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long


Sub TextTrans(MyTB As TextBox)
Dim TempDC As Long
Dim Temp As String
Dim MyLoc As RECT
Temp = MyTB.Text
MyLoc.Left = MyTB.Left
MyLoc.Top = MyTB.Top
MyLoc.Right = MyLoc.Left + MyTB.Width
MyLoc.Bottom = MyLoc.Top + MyTB.Height
MyTB.Parent.Cls
MyTB.Parent.ForeColor = MyTB.ForeColor
Set MyTB.Parent.Font = MyTB.Font
DrawText MyTB.Parent.hdc, Temp, Len(Temp), MyLoc, DT_EDITCONTROL
TempDC = GetDC(MyTB.hWnd)
BitBlt TempDC, 0, 0, MyTB.Width, MyTB.Height, MyTB.Parent.hdc, MyTB.Left, MyTB.Top, vbSrcCopy
End Sub

Private Sub cmdMakeTransparent_Click()
TextTrans txtInfo
DoIt = True
End Sub




Private Sub txtInfo_Change()
If Not DoIt Then Exit Sub
TextTrans txtInfo
End Sub


Private Sub txtInfo_KeyPress(KeyAscii As Integer)
If Not DoIt Then Exit Sub
TextTrans txtInfo
End Sub

Private Sub txtInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DoIt Then Exit Sub
TextTrans txtInfo
End Sub


