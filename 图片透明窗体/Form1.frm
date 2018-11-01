VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "透明颜色007F7F7F"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10890
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   522
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   726
   StartUpPosition =   3  '窗口缺省
   Begin VB.Image Image1 
      Height          =   525
      Left            =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H007F7F7F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   1575
      Left            =   750
      Top             =   5070
      Visible         =   0   'False
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Dim x1 As Long
Dim y1 As Long
Dim buttonB As Boolean

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'SetWindowPos(&wndTopMost,0,0,0,0, SWP_NOMOVE | SWP_NOSIZE);
'
'SetWindowPos(&wndNoTopMost,0,0,0,0, SWP_NOMOVE | SWP_NOSIZE);
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    x1 = x
    y1 = y
    buttonB = True
     Shape1.BorderStyle = 3
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x2 As Long, y2 As Long
If (buttonB) Then

    x2 = x
    y2 = y
    
    If Shape1.Visible = False Then
        Shape1.Visible = buttonB
    End If
    Dim left1 As Long, top1 As Long, wid As Long, hei As Long
    If (x1 < x2) Then
        left1 = x1
        wid = x2 - x1
        Shape1.BorderColor = &HFF
    Else
        left1 = x2
        wid = x1 - x2
                Shape1.BorderColor = &HFF0000

    End If
    If (y1 < y2) Then
        top1 = y1
        hei = y2 - y1
    Else
        top1 = y2
        hei = y1 - y2
    End If
    
    Shape1.Move left1, top1, wid, hei
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x2 As Long, y2 As Long
If Button = 1 Then
    x2 = x
    y2 = y
    buttonB = False
    
        Dim left1 As Long, top1 As Long, wid As Long, hei As Long
    If (x1 < x2) Then
        left1 = x1
        wid = x2 - x1
    Else
        left1 = x2
        wid = x1 - x2
    End If
    If (y1 < y2) Then
        top1 = y1
        hei = y2 - y1
    Else
        top1 = y2
        hei = y1 - y2
    End If
    
    Command1Click
    'Shape1.Visible = buttonB
    Shape1.BorderStyle = 0
End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo err1
Dim pic As String
If (Data.Files.Count <> 1) Then
    Exit Sub
End If
pic = Data.Files.Item(1)

Image1.Picture = VB.LoadPicture(pic)
Me.Width = Image1.Width * 15 + 300
Me.Height = Image1.Height * 15 + 500

Me.Picture = Image1.Picture
Command1Click
err1:
End Sub
Private Sub Command1Click()

    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, RGB(127, 127, 127), 0, LWA_COLORKEY '将扣去窗口中的蓝色
End Sub
