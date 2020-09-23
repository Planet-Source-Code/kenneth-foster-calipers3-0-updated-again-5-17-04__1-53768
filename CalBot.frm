VERSION 5.00
Begin VB.Form CalBot 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CalBot.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   240
      Left            =   855
      TabIndex        =   7
      Top             =   705
      Width           =   150
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<----->"
      Height          =   195
      Left            =   570
      TabIndex        =   6
      Top             =   960
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      Height          =   210
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   465
      Width           =   225
   End
   Begin VB.CommandButton cmdBB 
      Caption         =   "v"
      Height          =   240
      Left            =   1155
      TabIndex        =   4
      Top             =   705
      Width           =   135
   End
   Begin VB.CommandButton cmdBT 
      Caption         =   "^"
      Height          =   240
      Left            =   1020
      TabIndex        =   3
      Top             =   705
      Width           =   135
   End
   Begin VB.CommandButton cmdTB 
      Caption         =   "v"
      Height          =   240
      Left            =   690
      TabIndex        =   2
      Top             =   705
      Width           =   135
   End
   Begin VB.CommandButton cmdTT 
      Caption         =   "^"
      Height          =   240
      Left            =   540
      TabIndex        =   1
      Top             =   705
      Width           =   135
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   615
      TabIndex        =   0
      Top             =   1170
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   60
      Top             =   1185
   End
End
Attribute VB_Name = "CalBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Delcares
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim bytRegion(815) As Byte
Dim nBytes As Long
Private OldX As Integer
Private OldY As Integer
Private Sub Form_Load()
Dim rgnMain As Long

nBytes = 816

LoadBytes

rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
SetWindowRgn Me.hwnd, rgnMain, True

SetWindowPos hwnd, conHwndTopmost, 100, 100, 400, 141, conSwpNoActivate Or conSwpShowWindow


  
   CalBot.Show
   NoFocusRect cmdBB, True
    NoFocusRect cmdBT, True
    NoFocusRect cmdTB, True
    NoFocusRect cmdTT, True
    NoFocusRect Command1, True
    NoFocusRect Command2, True
End Sub
Private Sub cmdTB_Click()
CalTop.Top = CalTop.Top + 15
End Sub

Private Sub cmdTT_Click()
CalTop.Top = CalTop.Top - 15
End Sub

Private Sub cmdBB_Click()
CalBot.Top = CalBot.Top + 15
End Sub

Private Sub cmdBT_Click()
CalBot.Top = CalBot.Top - 15
End Sub

Private Sub Command1_Click()
Unload CalBot
Unload CalTop
Unload LeftCal
Unload RightCal
End Sub

Private Sub Command2_Click()
Unload CalBot
Unload CalTop
LeftCal.Show
RightCal.Show

End Sub

Private Sub Command3_Click()
If CalBot.Line1.BorderColor = vbBlack Then
     CalBot.Line1.BorderColor = vbWhite
     CalBot.Line2.BorderColor = vbWhite
     CalTop.Line1.BorderColor = vbWhite
     CalTop.Line2.BorderColor = vbWhite
 Else
     CalBot.Line1.BorderColor = vbBlack
     CalBot.Line2.BorderColor = vbBlack
     CalTop.Line1.BorderColor = vbBlack
     CalTop.Line2.BorderColor = vbBlack
 End If
End Sub

Private Sub Form_Activate()
CalTop.Left = CalBot.Left

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      OldX = x
      OldY = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      If Button = 1 Then
           ' Me.Left = Me.Left + (X - OldX)
            Me.Top = Me.Top + (y - OldY)
      End If
End Sub

Private Sub Timer1_Timer()
If CalTop.Top > CalBot.Top Then CalTop.Top = CalBot.Top
Text1.Text = (CalBot.Top - (CalTop.Top)) \ 15
End Sub

Private Sub LoadBytes()
bytRegion(0) = 32
bytRegion(4) = 1
bytRegion(8) = 49
bytRegion(12) = 16
bytRegion(13) = 3
bytRegion(24) = 120
bytRegion(28) = 99
bytRegion(32) = 59
bytRegion(40) = 60
bytRegion(44) = 1
bytRegion(48) = 59
bytRegion(52) = 1
bytRegion(56) = 62
bytRegion(60) = 2
bytRegion(64) = 59
bytRegion(68) = 2
bytRegion(72) = 64
bytRegion(76) = 3
bytRegion(80) = 59
bytRegion(84) = 3
bytRegion(88) = 66
bytRegion(92) = 4
bytRegion(96) = 59
bytRegion(100) = 4
bytRegion(104) = 68
bytRegion(108) = 5
bytRegion(112) = 59
bytRegion(116) = 5
bytRegion(120) = 70
bytRegion(124) = 6
bytRegion(128) = 59
bytRegion(132) = 6
bytRegion(136) = 72
bytRegion(140) = 7
bytRegion(144) = 59
bytRegion(148) = 7
bytRegion(152) = 74
bytRegion(156) = 8
bytRegion(160) = 59
bytRegion(164) = 8
bytRegion(168) = 76
bytRegion(172) = 9
bytRegion(176) = 59
bytRegion(180) = 9
bytRegion(184) = 78
bytRegion(188) = 10
bytRegion(192) = 59
bytRegion(196) = 10
bytRegion(200) = 80
bytRegion(204) = 11
bytRegion(208) = 59
bytRegion(212) = 11
bytRegion(216) = 82
bytRegion(220) = 12
bytRegion(224) = 59
bytRegion(228) = 12
bytRegion(232) = 84
bytRegion(236) = 13
bytRegion(240) = 59
bytRegion(244) = 13
bytRegion(248) = 86
bytRegion(252) = 14
bytRegion(256) = 59
bytRegion(260) = 14
bytRegion(264) = 88
bytRegion(268) = 15
bytRegion(272) = 59
bytRegion(276) = 15
bytRegion(280) = 90
bytRegion(284) = 16
bytRegion(288) = 59
bytRegion(292) = 16
bytRegion(296) = 92
bytRegion(300) = 17
bytRegion(304) = 59
bytRegion(308) = 17
bytRegion(312) = 94
bytRegion(316) = 18
bytRegion(320) = 59
bytRegion(324) = 18
bytRegion(328) = 96
bytRegion(332) = 19
bytRegion(336) = 59
bytRegion(340) = 19
bytRegion(344) = 98
bytRegion(348) = 20
bytRegion(352) = 59
bytRegion(356) = 20
bytRegion(360) = 100
bytRegion(364) = 21
bytRegion(368) = 59
bytRegion(372) = 21
bytRegion(376) = 102
bytRegion(380) = 22
bytRegion(384) = 59
bytRegion(388) = 22
bytRegion(392) = 104
bytRegion(396) = 23
bytRegion(400) = 59
bytRegion(404) = 23
bytRegion(408) = 106
bytRegion(412) = 24
bytRegion(416) = 59
bytRegion(420) = 24
bytRegion(424) = 108
bytRegion(428) = 25
bytRegion(432) = 59
bytRegion(436) = 25
bytRegion(440) = 110
bytRegion(444) = 26
bytRegion(448) = 59
bytRegion(452) = 26
bytRegion(456) = 112
bytRegion(460) = 27
bytRegion(464) = 59
bytRegion(468) = 27
bytRegion(472) = 114
bytRegion(476) = 28
bytRegion(480) = 59
bytRegion(484) = 28
bytRegion(488) = 116
bytRegion(492) = 29
bytRegion(496) = 59
bytRegion(500) = 29
bytRegion(504) = 120
bytRegion(508) = 30
bytRegion(516) = 30
bytRegion(520) = 88
bytRegion(524) = 31
bytRegion(528) = 1
bytRegion(532) = 31
bytRegion(536) = 88
bytRegion(540) = 32
bytRegion(544) = 3
bytRegion(548) = 32
bytRegion(552) = 88
bytRegion(556) = 33
bytRegion(560) = 5
bytRegion(564) = 33
bytRegion(568) = 88
bytRegion(572) = 34
bytRegion(576) = 7
bytRegion(580) = 34
bytRegion(584) = 88
bytRegion(588) = 35
bytRegion(592) = 9
bytRegion(596) = 35
bytRegion(600) = 88
bytRegion(604) = 36
bytRegion(608) = 11
bytRegion(612) = 36
bytRegion(616) = 88
bytRegion(620) = 37
bytRegion(624) = 13
bytRegion(628) = 37
bytRegion(632) = 88
bytRegion(636) = 38
bytRegion(640) = 15
bytRegion(644) = 38
bytRegion(648) = 88
bytRegion(652) = 39
bytRegion(656) = 17
bytRegion(660) = 39
bytRegion(664) = 88
bytRegion(668) = 40
bytRegion(672) = 19
bytRegion(676) = 40
bytRegion(680) = 88
bytRegion(684) = 41
bytRegion(688) = 21
bytRegion(692) = 41
bytRegion(696) = 88
bytRegion(700) = 42
bytRegion(704) = 23
bytRegion(708) = 42
bytRegion(712) = 88
bytRegion(716) = 43
bytRegion(720) = 25
bytRegion(724) = 43
bytRegion(728) = 88
bytRegion(732) = 44
bytRegion(736) = 27
bytRegion(740) = 44
bytRegion(744) = 88
bytRegion(748) = 45
bytRegion(752) = 29
bytRegion(756) = 45
bytRegion(760) = 88
bytRegion(764) = 46
bytRegion(768) = 31
bytRegion(772) = 46
bytRegion(776) = 88
bytRegion(780) = 47
bytRegion(784) = 33
bytRegion(788) = 47
bytRegion(792) = 88
bytRegion(796) = 48
bytRegion(800) = 35
bytRegion(804) = 48
bytRegion(808) = 88
bytRegion(812) = 99
End Sub
