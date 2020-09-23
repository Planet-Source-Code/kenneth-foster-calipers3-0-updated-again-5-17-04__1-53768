VERSION 5.00
Begin VB.Form LeftCal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LeftCal.frx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   60
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line2 
      X1              =   31
      X2              =   31
      Y1              =   1
      Y2              =   61
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   30
      Y1              =   61
      Y2              =   120
   End
End
Attribute VB_Name = "LeftCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Delcares
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim bytRegion(975) As Byte
Dim nBytes As Long
Private OldX As Integer   'used to move forms
Private OldY As Integer   'used to move forms

Private Sub Form_Activate()
    RightCal.Left = Me.Left + 1000  'sets up some perimeters
    RightCal.Top = Me.Top
End Sub
'*******************************************************************************
' *  Title            :     Calipers
' *  Author         : Ken Foster
' *  Purpose        : Measure objects on screen in pixels.
' *
' *                  : This is freeware.No copyrights ,no license,no agreements.
' *                   Just hope it is helpful. Have fun and Howdy from Texas.
' *  Date           : 2004
'*******************************************************************************
Private Sub Form_Load()
Dim rgnMain As Long

nBytes = 976

LoadBytes

rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
SetWindowRgn Me.hwnd, rgnMain, True

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    OldX = x
    OldY = y
End Sub
    
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then   'left mouse button
    Me.Left = Me.Left + (x - OldX)
    Me.Top = Me.Top + (y - OldY)
    RightCal.Top = Me.Top
    RightCal.Left = LeftCal.Left + Int(RightCal.Text1.Text * 15) + 10
End If

End Sub
Private Sub LoadBytes()
bytRegion(0) = 32
bytRegion(4) = 1
bytRegion(8) = 59
bytRegion(12) = 176
bytRegion(13) = 3
bytRegion(16) = 1
bytRegion(24) = 60
bytRegion(28) = 120
bytRegion(32) = 31
bytRegion(40) = 32
bytRegion(44) = 4
bytRegion(48) = 31
bytRegion(52) = 4
bytRegion(56) = 33
bytRegion(60) = 6
bytRegion(64) = 31
bytRegion(68) = 6
bytRegion(72) = 34
bytRegion(76) = 8
bytRegion(80) = 31
bytRegion(84) = 8
bytRegion(88) = 35
bytRegion(92) = 10
bytRegion(96) = 31
bytRegion(100) = 10
bytRegion(104) = 36
bytRegion(108) = 12
bytRegion(112) = 31
bytRegion(116) = 12
bytRegion(120) = 37
bytRegion(124) = 14
bytRegion(128) = 31
bytRegion(132) = 14
bytRegion(136) = 38
bytRegion(140) = 16
bytRegion(144) = 31
bytRegion(148) = 16
bytRegion(152) = 39
bytRegion(156) = 18
bytRegion(160) = 31
bytRegion(164) = 18
bytRegion(168) = 40
bytRegion(172) = 20
bytRegion(176) = 31
bytRegion(180) = 20
bytRegion(184) = 41
bytRegion(188) = 22
bytRegion(192) = 31
bytRegion(196) = 22
bytRegion(200) = 42
bytRegion(204) = 24
bytRegion(208) = 31
bytRegion(212) = 24
bytRegion(216) = 43
bytRegion(220) = 26
bytRegion(224) = 31
bytRegion(228) = 26
bytRegion(232) = 44
bytRegion(236) = 28
bytRegion(240) = 31
bytRegion(244) = 28
bytRegion(248) = 45
bytRegion(252) = 30
bytRegion(256) = 31
bytRegion(260) = 30
bytRegion(264) = 46
bytRegion(268) = 32
bytRegion(272) = 31
bytRegion(276) = 32
bytRegion(280) = 47
bytRegion(284) = 34
bytRegion(288) = 31
bytRegion(292) = 34
bytRegion(296) = 48
bytRegion(300) = 36
bytRegion(304) = 31
bytRegion(308) = 36
bytRegion(312) = 49
bytRegion(316) = 38
bytRegion(320) = 31
bytRegion(324) = 38
bytRegion(328) = 50
bytRegion(332) = 40
bytRegion(336) = 31
bytRegion(340) = 40
bytRegion(344) = 51
bytRegion(348) = 42
bytRegion(352) = 31
bytRegion(356) = 42
bytRegion(360) = 52
bytRegion(364) = 44
bytRegion(368) = 31
bytRegion(372) = 44
bytRegion(376) = 53
bytRegion(380) = 46
bytRegion(384) = 31
bytRegion(388) = 46
bytRegion(392) = 54
bytRegion(396) = 48
bytRegion(400) = 31
bytRegion(404) = 48
bytRegion(408) = 55
bytRegion(412) = 50
bytRegion(416) = 31
bytRegion(420) = 50
bytRegion(424) = 56
bytRegion(428) = 52
bytRegion(432) = 31
bytRegion(436) = 52
bytRegion(440) = 57
bytRegion(444) = 54
bytRegion(448) = 31
bytRegion(452) = 54
bytRegion(456) = 58
bytRegion(460) = 56
bytRegion(464) = 31
bytRegion(468) = 56
bytRegion(472) = 59
bytRegion(476) = 58
bytRegion(480) = 31
bytRegion(484) = 58
bytRegion(488) = 60
bytRegion(492) = 60
bytRegion(496) = 1
bytRegion(500) = 60
bytRegion(504) = 60
bytRegion(508) = 61
bytRegion(512) = 1
bytRegion(516) = 61
bytRegion(520) = 31
bytRegion(524) = 63
bytRegion(528) = 2
bytRegion(532) = 63
bytRegion(536) = 31
bytRegion(540) = 65
bytRegion(544) = 3
bytRegion(548) = 65
bytRegion(552) = 31
bytRegion(556) = 67
bytRegion(560) = 4
bytRegion(564) = 67
bytRegion(568) = 31
bytRegion(572) = 69
bytRegion(576) = 5
bytRegion(580) = 69
bytRegion(584) = 31
bytRegion(588) = 71
bytRegion(592) = 6
bytRegion(596) = 71
bytRegion(600) = 31
bytRegion(604) = 73
bytRegion(608) = 7
bytRegion(612) = 73
bytRegion(616) = 31
bytRegion(620) = 75
bytRegion(624) = 8
bytRegion(628) = 75
bytRegion(632) = 31
bytRegion(636) = 77
bytRegion(640) = 9
bytRegion(644) = 77
bytRegion(648) = 31
bytRegion(652) = 79
bytRegion(656) = 10
bytRegion(660) = 79
bytRegion(664) = 31
bytRegion(668) = 81
bytRegion(672) = 11
bytRegion(676) = 81
bytRegion(680) = 31
bytRegion(684) = 83
bytRegion(688) = 12
bytRegion(692) = 83
bytRegion(696) = 31
bytRegion(700) = 85
bytRegion(704) = 13
bytRegion(708) = 85
bytRegion(712) = 31
bytRegion(716) = 87
bytRegion(720) = 14
bytRegion(724) = 87
bytRegion(728) = 31
bytRegion(732) = 89
bytRegion(736) = 15
bytRegion(740) = 89
bytRegion(744) = 31
bytRegion(748) = 91
bytRegion(752) = 16
bytRegion(756) = 91
bytRegion(760) = 31
bytRegion(764) = 93
bytRegion(768) = 17
bytRegion(772) = 93
bytRegion(776) = 31
bytRegion(780) = 95
bytRegion(784) = 18
bytRegion(788) = 95
bytRegion(792) = 31
bytRegion(796) = 97
bytRegion(800) = 19
bytRegion(804) = 97
bytRegion(808) = 31
bytRegion(812) = 99
bytRegion(816) = 20
bytRegion(820) = 99
bytRegion(824) = 31
bytRegion(828) = 101
bytRegion(832) = 21
bytRegion(836) = 101
bytRegion(840) = 31
bytRegion(844) = 103
bytRegion(848) = 22
bytRegion(852) = 103
bytRegion(856) = 31
bytRegion(860) = 105
bytRegion(864) = 23
bytRegion(868) = 105
bytRegion(872) = 31
bytRegion(876) = 107
bytRegion(880) = 24
bytRegion(884) = 107
bytRegion(888) = 31
bytRegion(892) = 109
bytRegion(896) = 25
bytRegion(900) = 109
bytRegion(904) = 31
bytRegion(908) = 111
bytRegion(912) = 26
bytRegion(916) = 111
bytRegion(920) = 31
bytRegion(924) = 113
bytRegion(928) = 27
bytRegion(932) = 113
bytRegion(936) = 31
bytRegion(940) = 115
bytRegion(944) = 28
bytRegion(948) = 115
bytRegion(952) = 31
bytRegion(956) = 117
bytRegion(960) = 29
bytRegion(964) = 117
bytRegion(968) = 31
bytRegion(972) = 120
End Sub
