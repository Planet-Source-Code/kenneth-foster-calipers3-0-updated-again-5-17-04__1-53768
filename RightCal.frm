VERSION 5.00
Begin VB.Form RightCal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RightCal.frx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVH 
      Caption         =   "<---->"
      Height          =   195
      Left            =   510
      TabIndex        =   7
      Top             =   795
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "C"
      Height          =   225
      Left            =   840
      TabIndex        =   6
      Top             =   1005
      Width           =   270
   End
   Begin VB.CommandButton cmdRR 
      Caption         =   ">"
      Height          =   225
      Left            =   1320
      TabIndex        =   5
      Top             =   1005
      Width           =   135
   End
   Begin VB.CommandButton cmdRL 
      Caption         =   "<"
      Height          =   225
      Left            =   1155
      TabIndex        =   4
      Top             =   1005
      Width           =   135
   End
   Begin VB.CommandButton cmdLR 
      Caption         =   ">"
      Height          =   225
      Left            =   660
      TabIndex        =   3
      Top             =   1005
      Width           =   135
   End
   Begin VB.CommandButton cmdLL 
      Caption         =   "<"
      Height          =   225
      Left            =   495
      TabIndex        =   2
      Top             =   1005
      Width           =   135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      Height          =   225
      Left            =   630
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   525
      Width           =   225
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   885
      TabIndex        =   0
      Top             =   525
      Width           =   525
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   930
      Top             =   1335
   End
   Begin VB.Line Line2 
      X1              =   30
      X2              =   30
      Y1              =   60
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   29
      X2              =   29
      Y1              =   32
      Y2              =   0
   End
End
Attribute VB_Name = "RightCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Delcares
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim bytRegion(847) As Byte
Dim nBytes As Long
Private OldX1 As Integer
Private OldY1 As Integer
Private Sub Form_Load()
Dim rgnMain As Long

nBytes = 848

LoadBytes

rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
SetWindowRgn Me.hwnd, rgnMain, True
SetWindowPos hwnd, conHwndTopmost, 100, 100, 400, 141, conSwpNoActivate Or conSwpShowWindow

  RightCal.Show
    
    NoFocusRect cmdLL, True
    NoFocusRect cmdLR, True
    NoFocusRect cmdRL, True
    NoFocusRect cmdRR, True
    NoFocusRect cmdVH, True
    NoFocusRect Command1, True
    NoFocusRect Command2, True
    Command1.SetFocus
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    OldX1 = x
    OldY1 = y
End Sub
    
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.Left = Me.Left + (x - OldX1)
      ' Me.Top = Me.Top + (Y - OldY1)
        LeftCal.Top = Me.Top
    End If
    
End Sub
Private Sub Timer1_Timer()
    If LeftCal.Left > RightCal.Left Then LeftCal.Left = RightCal.Left - 10
    Text1.Text = (RightCal.Left - (LeftCal.Left)) \ 15
End Sub
Private Sub cmdLL_Click()
    LeftCal.Left = LeftCal.Left - 15
    
End Sub
    
Private Sub cmdLR_Click()
    LeftCal.Left = LeftCal.Left + 15
End Sub
    
Private Sub cmdRL_Click()
    Me.Left = Me.Left - 15
End Sub
    
Private Sub cmdRR_Click()
    Me.Left = Me.Left + 15
End Sub
    
Private Sub cmdVH_Click()
    Unload LeftCal
    Unload RightCal
    CalTop.Show
    CalBot.Show
End Sub
    
Private Sub Command1_Click()
    Unload Me
    Unload LeftCal
    Unload CalTop
    Unload CalBot
End Sub
    
Private Sub Command2_Click()
If RightCal.Line1.BorderColor = vbBlack Then
     RightCal.Line1.BorderColor = vbWhite
     RightCal.Line2.BorderColor = vbWhite
     LeftCal.Line1.BorderColor = vbWhite
     LeftCal.Line2.BorderColor = vbWhite
 Else
     RightCal.Line1.BorderColor = vbBlack
     RightCal.Line2.BorderColor = vbBlack
     LeftCal.Line1.BorderColor = vbBlack
     LeftCal.Line2.BorderColor = vbBlack
 End If
End Sub
Private Sub LoadBytes()
bytRegion(0) = 32
bytRegion(4) = 1
bytRegion(8) = 51
bytRegion(12) = 48
bytRegion(13) = 3
bytRegion(24) = 99
bytRegion(28) = 120
bytRegion(32) = 29
bytRegion(40) = 30
bytRegion(44) = 4
bytRegion(48) = 28
bytRegion(52) = 4
bytRegion(56) = 30
bytRegion(60) = 6
bytRegion(64) = 27
bytRegion(68) = 6
bytRegion(72) = 30
bytRegion(76) = 8
bytRegion(80) = 26
bytRegion(84) = 8
bytRegion(88) = 30
bytRegion(92) = 10
bytRegion(96) = 25
bytRegion(100) = 10
bytRegion(104) = 26
bytRegion(108) = 11
bytRegion(112) = 27
bytRegion(116) = 10
bytRegion(120) = 30
bytRegion(124) = 11
bytRegion(128) = 25
bytRegion(132) = 11
bytRegion(136) = 30
bytRegion(140) = 12
bytRegion(144) = 24
bytRegion(148) = 12
bytRegion(152) = 30
bytRegion(156) = 14
bytRegion(160) = 23
bytRegion(164) = 14
bytRegion(168) = 30
bytRegion(172) = 16
bytRegion(176) = 22
bytRegion(180) = 16
bytRegion(184) = 30
bytRegion(188) = 18
bytRegion(192) = 21
bytRegion(196) = 18
bytRegion(200) = 30
bytRegion(204) = 20
bytRegion(208) = 20
bytRegion(212) = 20
bytRegion(216) = 30
bytRegion(220) = 22
bytRegion(224) = 19
bytRegion(228) = 22
bytRegion(232) = 30
bytRegion(236) = 24
bytRegion(240) = 18
bytRegion(244) = 24
bytRegion(248) = 30
bytRegion(252) = 26
bytRegion(256) = 17
bytRegion(260) = 26
bytRegion(264) = 30
bytRegion(268) = 28
bytRegion(272) = 16
bytRegion(276) = 28
bytRegion(280) = 30
bytRegion(284) = 30
bytRegion(288) = 15
bytRegion(292) = 30
bytRegion(296) = 30
bytRegion(300) = 32
bytRegion(304) = 14
bytRegion(308) = 32
bytRegion(312) = 99
bytRegion(316) = 34
bytRegion(320) = 13
bytRegion(324) = 34
bytRegion(328) = 99
bytRegion(332) = 36
bytRegion(336) = 12
bytRegion(340) = 36
bytRegion(344) = 99
bytRegion(348) = 38
bytRegion(352) = 11
bytRegion(356) = 38
bytRegion(360) = 99
bytRegion(364) = 40
bytRegion(368) = 10
bytRegion(372) = 40
bytRegion(376) = 99
bytRegion(380) = 42
bytRegion(384) = 9
bytRegion(388) = 42
bytRegion(392) = 99
bytRegion(396) = 44
bytRegion(400) = 8
bytRegion(404) = 44
bytRegion(408) = 99
bytRegion(412) = 46
bytRegion(416) = 7
bytRegion(420) = 46
bytRegion(424) = 99
bytRegion(428) = 48
bytRegion(432) = 6
bytRegion(436) = 48
bytRegion(440) = 99
bytRegion(444) = 50
bytRegion(448) = 5
bytRegion(452) = 50
bytRegion(456) = 99
bytRegion(460) = 52
bytRegion(464) = 4
bytRegion(468) = 52
bytRegion(472) = 99
bytRegion(476) = 54
bytRegion(480) = 3
bytRegion(484) = 54
bytRegion(488) = 99
bytRegion(492) = 56
bytRegion(496) = 2
bytRegion(500) = 56
bytRegion(504) = 99
bytRegion(508) = 58
bytRegion(512) = 1
bytRegion(516) = 58
bytRegion(520) = 99
bytRegion(524) = 60
bytRegion(532) = 60
bytRegion(536) = 99
bytRegion(540) = 61
bytRegion(544) = 30
bytRegion(548) = 61
bytRegion(552) = 99
bytRegion(556) = 85
bytRegion(560) = 30
bytRegion(564) = 85
bytRegion(568) = 48
bytRegion(572) = 87
bytRegion(576) = 30
bytRegion(580) = 87
bytRegion(584) = 47
bytRegion(588) = 89
bytRegion(592) = 30
bytRegion(596) = 89
bytRegion(600) = 46
bytRegion(604) = 91
bytRegion(608) = 30
bytRegion(612) = 91
bytRegion(616) = 45
bytRegion(620) = 93
bytRegion(624) = 30
bytRegion(628) = 93
bytRegion(632) = 44
bytRegion(636) = 95
bytRegion(640) = 30
bytRegion(644) = 95
bytRegion(648) = 43
bytRegion(652) = 97
bytRegion(656) = 30
bytRegion(660) = 97
bytRegion(664) = 42
bytRegion(668) = 99
bytRegion(672) = 30
bytRegion(676) = 99
bytRegion(680) = 41
bytRegion(684) = 101
bytRegion(688) = 30
bytRegion(692) = 101
bytRegion(696) = 40
bytRegion(700) = 103
bytRegion(704) = 30
bytRegion(708) = 103
bytRegion(712) = 39
bytRegion(716) = 105
bytRegion(720) = 30
bytRegion(724) = 105
bytRegion(728) = 38
bytRegion(732) = 107
bytRegion(736) = 30
bytRegion(740) = 107
bytRegion(744) = 37
bytRegion(748) = 109
bytRegion(752) = 30
bytRegion(756) = 109
bytRegion(760) = 36
bytRegion(764) = 111
bytRegion(768) = 30
bytRegion(772) = 111
bytRegion(776) = 35
bytRegion(780) = 113
bytRegion(784) = 30
bytRegion(788) = 113
bytRegion(792) = 34
bytRegion(796) = 115
bytRegion(800) = 30
bytRegion(804) = 115
bytRegion(808) = 33
bytRegion(812) = 117
bytRegion(816) = 30
bytRegion(820) = 117
bytRegion(824) = 32
bytRegion(828) = 119
bytRegion(832) = 30
bytRegion(836) = 119
bytRegion(840) = 31
bytRegion(844) = 120
End Sub
