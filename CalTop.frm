VERSION 5.00
Begin VB.Form CalTop 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CalTop.frx":0000
   ScaleHeight     =   900
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "CalTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Delcares
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim bytRegion(959) As Byte
Dim nBytes As Long
Private OldX As Integer
Private OldY As Integer
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      OldX = x
      OldY = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      If Button = 1 Then
            Me.Left = Me.Left + (x - OldX)
             Me.Top = Me.Top + (y - OldY)
           CalBot.Left = CalTop.Left
           CalBot.Top = CalTop.Top + Int(CalBot.Text1.Text) * 15
       If CalBot.Top < CalTop.Top Then CalBot.Top = CalTop.Top
        
      End If
End Sub
Private Sub Form_Load()
Dim rgnMain As Long

nBytes = 960

LoadBytes
CalTop.Show
CalBot.Show
 CalTop.Left = CalBot.Left
 CalTop.Top = CalBot.Top - 1000
rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
SetWindowRgn Me.hwnd, rgnMain, True

End Sub
Private Sub LoadBytes()
bytRegion(0) = 32
bytRegion(4) = 1
bytRegion(8) = 58
bytRegion(12) = 160
bytRegion(13) = 3
bytRegion(20) = 1
bytRegion(24) = 120
bytRegion(28) = 60
bytRegion(32) = 57
bytRegion(36) = 1
bytRegion(40) = 60
bytRegion(44) = 2
bytRegion(48) = 55
bytRegion(52) = 2
bytRegion(56) = 60
bytRegion(60) = 3
bytRegion(64) = 53
bytRegion(68) = 3
bytRegion(72) = 60
bytRegion(76) = 4
bytRegion(80) = 51
bytRegion(84) = 4
bytRegion(88) = 60
bytRegion(92) = 5
bytRegion(96) = 49
bytRegion(100) = 5
bytRegion(104) = 60
bytRegion(108) = 6
bytRegion(112) = 47
bytRegion(116) = 6
bytRegion(120) = 60
bytRegion(124) = 7
bytRegion(128) = 45
bytRegion(132) = 7
bytRegion(136) = 60
bytRegion(140) = 8
bytRegion(144) = 43
bytRegion(148) = 8
bytRegion(152) = 60
bytRegion(156) = 9
bytRegion(160) = 41
bytRegion(164) = 9
bytRegion(168) = 60
bytRegion(172) = 10
bytRegion(176) = 39
bytRegion(180) = 10
bytRegion(184) = 60
bytRegion(188) = 11
bytRegion(192) = 37
bytRegion(196) = 11
bytRegion(200) = 60
bytRegion(204) = 12
bytRegion(208) = 35
bytRegion(212) = 12
bytRegion(216) = 60
bytRegion(220) = 13
bytRegion(224) = 33
bytRegion(228) = 13
bytRegion(232) = 60
bytRegion(236) = 14
bytRegion(240) = 31
bytRegion(244) = 14
bytRegion(248) = 60
bytRegion(252) = 15
bytRegion(256) = 29
bytRegion(260) = 15
bytRegion(264) = 60
bytRegion(268) = 16
bytRegion(272) = 27
bytRegion(276) = 16
bytRegion(280) = 60
bytRegion(284) = 17
bytRegion(288) = 25
bytRegion(292) = 17
bytRegion(296) = 60
bytRegion(300) = 18
bytRegion(304) = 23
bytRegion(308) = 18
bytRegion(312) = 60
bytRegion(316) = 19
bytRegion(320) = 21
bytRegion(324) = 19
bytRegion(328) = 60
bytRegion(332) = 20
bytRegion(336) = 19
bytRegion(340) = 20
bytRegion(344) = 60
bytRegion(348) = 21
bytRegion(352) = 17
bytRegion(356) = 21
bytRegion(360) = 60
bytRegion(364) = 22
bytRegion(368) = 15
bytRegion(372) = 22
bytRegion(376) = 60
bytRegion(380) = 23
bytRegion(384) = 13
bytRegion(388) = 23
bytRegion(392) = 60
bytRegion(396) = 24
bytRegion(400) = 11
bytRegion(404) = 24
bytRegion(408) = 60
bytRegion(412) = 25
bytRegion(416) = 9
bytRegion(420) = 25
bytRegion(424) = 60
bytRegion(428) = 26
bytRegion(432) = 7
bytRegion(436) = 26
bytRegion(440) = 60
bytRegion(444) = 27
bytRegion(448) = 5
bytRegion(452) = 27
bytRegion(456) = 60
bytRegion(460) = 28
bytRegion(464) = 3
bytRegion(468) = 28
bytRegion(472) = 60
bytRegion(476) = 29
bytRegion(484) = 29
bytRegion(488) = 60
bytRegion(492) = 30
bytRegion(496) = 59
bytRegion(500) = 30
bytRegion(504) = 120
bytRegion(508) = 32
bytRegion(512) = 59
bytRegion(516) = 32
bytRegion(520) = 116
bytRegion(524) = 33
bytRegion(528) = 59
bytRegion(532) = 33
bytRegion(536) = 114
bytRegion(540) = 34
bytRegion(544) = 59
bytRegion(548) = 34
bytRegion(552) = 112
bytRegion(556) = 35
bytRegion(560) = 59
bytRegion(564) = 35
bytRegion(568) = 110
bytRegion(572) = 36
bytRegion(576) = 59
bytRegion(580) = 36
bytRegion(584) = 108
bytRegion(588) = 37
bytRegion(592) = 59
bytRegion(596) = 37
bytRegion(600) = 106
bytRegion(604) = 38
bytRegion(608) = 59
bytRegion(612) = 38
bytRegion(616) = 104
bytRegion(620) = 39
bytRegion(624) = 59
bytRegion(628) = 39
bytRegion(632) = 102
bytRegion(636) = 40
bytRegion(640) = 59
bytRegion(644) = 40
bytRegion(648) = 100
bytRegion(652) = 41
bytRegion(656) = 59
bytRegion(660) = 41
bytRegion(664) = 98
bytRegion(668) = 42
bytRegion(672) = 59
bytRegion(676) = 42
bytRegion(680) = 96
bytRegion(684) = 43
bytRegion(688) = 59
bytRegion(692) = 43
bytRegion(696) = 94
bytRegion(700) = 44
bytRegion(704) = 59
bytRegion(708) = 44
bytRegion(712) = 92
bytRegion(716) = 45
bytRegion(720) = 59
bytRegion(724) = 45
bytRegion(728) = 90
bytRegion(732) = 46
bytRegion(736) = 59
bytRegion(740) = 46
bytRegion(744) = 88
bytRegion(748) = 47
bytRegion(752) = 59
bytRegion(756) = 47
bytRegion(760) = 86
bytRegion(764) = 48
bytRegion(768) = 59
bytRegion(772) = 48
bytRegion(776) = 84
bytRegion(780) = 49
bytRegion(784) = 59
bytRegion(788) = 49
bytRegion(792) = 82
bytRegion(796) = 50
bytRegion(800) = 59
bytRegion(804) = 50
bytRegion(808) = 80
bytRegion(812) = 51
bytRegion(816) = 59
bytRegion(820) = 51
bytRegion(824) = 78
bytRegion(828) = 52
bytRegion(832) = 59
bytRegion(836) = 52
bytRegion(840) = 76
bytRegion(844) = 53
bytRegion(848) = 59
bytRegion(852) = 53
bytRegion(856) = 74
bytRegion(860) = 54
bytRegion(864) = 59
bytRegion(868) = 54
bytRegion(872) = 72
bytRegion(876) = 55
bytRegion(880) = 59
bytRegion(884) = 55
bytRegion(888) = 70
bytRegion(892) = 56
bytRegion(896) = 59
bytRegion(900) = 56
bytRegion(904) = 68
bytRegion(908) = 57
bytRegion(912) = 59
bytRegion(916) = 57
bytRegion(920) = 66
bytRegion(924) = 58
bytRegion(928) = 59
bytRegion(932) = 58
bytRegion(936) = 64
bytRegion(940) = 59
bytRegion(944) = 59
bytRegion(948) = 59
bytRegion(952) = 62
bytRegion(956) = 60
End Sub
