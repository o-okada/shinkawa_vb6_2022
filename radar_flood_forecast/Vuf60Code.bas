Attribute VB_Name = "Vuf60Code"
Option Explicit

Public sp(60)       As Double
Public tt(60)       As Double
Public mfday(12)    As Integer
Public ug(60)       As Double
Public vg(60)       As Double
Public tnm(60)      As String
Public AnlObsLong   As ObsDegree
Public AnlBaseTime  As ObsTime
'調和定数 60分調
Public amp(60)      As Double  'h
Public phs(60)      As Double  'κ

Public Th0          As Double  '平均潮位
Public rad          As Double  'π/180.

Type ObsDegree

     Deg As Integer
     Min As Integer
     Sec As Integer
     Tmz As Integer

End Type

Type ObsTime

     Year    As Integer
     Month   As Integer
     Day     As Integer
     Hour    As Integer
     Minute  As Integer

End Type

Public Const PI = 3.1415926535
'
'dw = 天文潮位計算現時刻
't0 = 現時刻天文潮位
't1 = 1時間後天文潮位
't2 = 2時間後天文潮位
't3 = 3時間後天文潮位
'
Sub Cal_Tide(dw As Date, t0 As Single, t1 As Single, t2 As Single, t3 As Single)

    Dim i            As Long
    Dim j            As Long
    Dim v(60)        As Double
    Dim F(60)        As Double
    Dim Ti           As Double
    Dim zz           As Double
    Dim dd           As Date
    Dim buf          As String

    '名古屋湾験潮位置
    AnlObsLong.Deg = 136#
    AnlObsLong.Min = 53#
    AnlObsLong.Sec = 0#
    AnlObsLong.Tmz = -9#

    For j = 0 To 3
        dd = DateAdd("h", j, dw)
        VufSetting v, F, dd
        zz = 0#
        For i = 1 To 60
            zz = zz + F(i) * amp(i) * Cos((v(i) - phs(i)) * rad)   '+ sgunit * sp(i) * CDbl(tt))
        Next i
        Ti = (zz + Th0) * 0.01
        Select Case j
            Case 0
                t0 = Ti
            Case 1
                t1 = Ti
            Case 2
                t2 = Ti
            Case 3
                t3 = Ti
        End Select
    Next j

    LOG_Out "OUT   Cal_Tide dw=" & TIMEC(dw) & "  T0=" & fmt(t0)
    LOG_Out "                                     T1=" & fmt(t1)
    LOG_Out "                                     T2=" & fmt(t2)
    LOG_Out "                                     T3=" & fmt(t3)

    If t0 = 0# And t1 = 0# Then
        LOG_Out "                                    rad=" & fmt(rad)
        LOG_Out "                                    Th0=" & fmt(Th0)
        For i = 1 To 60
            buf = "                                    i=" & Format(str(i), "@@@@") & "  "
            buf = buf & "    amp=" & fmt(amp(i))
            buf = buf & "    Cos=" & fmt(Cos(i))
            buf = buf & "    phs=" & fmt(phs(i))
            LOG_Out buf
        Next i
     End If

End Sub
Public Sub VufSetting(ByRef v() As Double, ByRef F() As Double, ByVal AnlTime As Date)
'Public Sub VufSetting(ByRef v() As Double, ByRef f() As Double, ByVal AnlLong As ObsDegree, ByVal AnlTime As Date)

    Dim iyr As Integer
    Dim imn As Integer
    Dim idy As Integer
    Dim ihr As Integer
    Dim mmn As Integer
    Dim kd  As Integer
    Dim km  As Integer
    Dim ks  As Integer
    Dim tz  As Integer

    kd = AnlObsLong.Deg
    km = AnlObsLong.Min
    ks = AnlObsLong.Sec
    tz = AnlObsLong.Tmz

    iyr = Year(AnlTime)
    imn = Month(AnlTime)
    idy = Day(AnlTime)
    ihr = Hour(AnlTime)
    mmn = Minute(AnlTime)

    Call vuf60(iyr, imn, idy, ihr, mmn, kd, km, ks, tz, v, F)

End Sub
Sub 分調読み込み()

    Dim i   As Long
    Dim nf  As Long
    Dim F   As String
    Dim buf As String
    Dim w
'
'amp(i)=ｈ
'phs(i)=κ
'
'
'60分調データ例
'            h         κ
'Sa    ,   16.0   ,  153.1
'Ssa   ,    1.4   ,  348.4
'Mm    ,    1.1   ,  207.1
'MSf   ,    3.0   ,   68.8
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
' ・        ・        ・
'
'
    LOG_Out "IN   分調読み込み"

    F = App.Path & "\data\60分調.dat"
    nf = FreeFile
    Open F For Input As #nf
    Line Input #nf, buf     'データタイトル
    For i = 1 To 60
        Line Input #nf, buf
        w = Split(buf, ",")
        amp(i) = w(1) '------------------ ｈ
        phs(i) = w(2) '------------------ κ
    Next i
    Line Input #nf, buf   '平均潮位
    w = Split(buf, ",")
    Th0 = w(1)
    Close #nf

    LOG_Out "OUT  分調読み込み"

End Sub
Public Sub vuf60(ByVal iy As Integer, ByVal im As Integer, ByVal id As Integer, ByVal ih As Integer, ByVal MM As Integer, _
                 ByVal kd As Integer, ByVal km As Integer, ByVal ks As Integer, ByVal tz As Integer, _
                 ByRef v() As Double, ByRef F() As Double)

'                                              modified 2006 Dec.1
'                          based on TIDAL HARMONIC CONSTANTS TABLES
'                                             (PUB.No.742 1992 Feb)

'(IY, IM, ID, IH, MM) : NEN, GATU, HI, JI, FUN
'(KD,KM,KS)           : KEIDO: DO,FUN,BYO
'(TZ)                 : TIME(ZONE(Hour))
'(V, F)               : V0(+U And F)

    Dim d     As Double
    Dim dfl   As Double
    Dim y     As Double
    Dim S     As Double
    Dim H     As Double
    Dim p     As Double
    Dim fn    As Double
    Dim fkd   As Double
    Dim ttt   As Double
    Dim C1    As Double
    Dim C2    As Double
    Dim C3    As Double
    Dim s1    As Double
    Dim s2    As Double
    Dim s3    As Double
    Dim a12   As Double
    Dim b12   As Double
    Dim a34   As Double
    Dim b34   As Double
    Dim v0u   As Double
    Dim leap  As Long
    Dim L     As Long
    Dim i     As Long

    spsetting

'    rad = Math.PI / 180#  'Change By MK 2008/04/04 17:30
    rad = PI / 180#

'うるう年の判別　サンプルはうそ
'    leap = 0
'    If iy Mod 4 = 0 And iy Mod 100 = 0 Then leap = 1
'    If iy Mod 400 = 0 Then leap = 1
'うるう年の判別　菊地修正
    leap = 0
    If iy Mod 4 = 0 Then
        leap = 1
    End If
    If iy Mod 100 = 0 Then
       leap = 0
    End If
    If iy Mod 400 = 0 Then
        leap = 1
    End If
    d = mfday(im) + id
    If leap = 1 And im >= 3 Then d = d + 1#
    L = Int((iy + 3) / 4) - 500
    dfl = d + CDbl(L)
    y = CDbl(iy)

    S = 211.728 + 129.38471 * (y - 2000#) + 13.176396 * dfl
    H = 279.974 - 0.23871 * (y - 2000#) + 0.985647 * dfl
    p = 83.298 + 40.66229 * (y - 2000#) + 0.111404 * dfl
    fn = 125.071 - 19.32812 * (y - 2000#) - 0.052954 * dfl

    C1 = Cos(fn * rad)
    C2 = Cos(2# * fn * rad)
    C3 = Cos(3# * fn * rad)
    s1 = Sin(fn * rad)
    s2 = Sin(2# * fn * rad)
    s3 = Sin(3# * fn * rad)

    ' - these values are valid even after 2000 -----------
    getv60 vg, S, H, p
    getu60 ug, p, fn, s1, s2, s3, a12, b12, a34, b34
    getf60 F, C1, C2, C3, a12, b12, a34, b34
    ' ------------------------------------------------------

    fkd = (CDbl(ks) / 60# + CDbl(km)) / 60# + CDbl(kd)
    For i = 1 To 60
        ttt = tt(i) * fkd + sp(i) * (CDbl(tz) + CDbl(ih) + CDbl(MM) / 60#)
        v0u = vg(i) + ug(i) + ttt
        v0u = v0u Mod 360#
        v(i) = (v0u + 360#) Mod 360#
    Next i

End Sub
Public Sub spsetting()

    Dim i As Long

     tnm(1) = "SA"
     tnm(2) = "SSA"
     tnm(3) = "MM"
     tnm(4) = "MSF"
     tnm(5) = "MF"
     tnm(6) = "2Q1"
     tnm(7) = "SGM1"
     tnm(8) = "Q1"
     tnm(9) = "RHO1"
    tnm(10) = "O1"
    tnm(11) = "MP1"
    tnm(12) = "M1"
    tnm(13) = "CHI1"
    tnm(14) = "PAI1"
    tnm(15) = "P1"
    tnm(16) = "S1"
    tnm(17) = "K1"
    tnm(18) = "PSI1"
    tnm(19) = "PHI1"
    tnm(20) = "THT1"
    tnm(21) = "J1"
    tnm(22) = "SO1"
    tnm(23) = "OO1"
    tnm(24) = "OQ2"
    tnm(25) = "MNS2"
    tnm(26) = "2N2"
    tnm(27) = "MU2"
    tnm(28) = "N2"
    tnm(29) = "NU2"
    tnm(30) = "OP2"
    tnm(31) = "M2"
    tnm(32) = "MKS2"
    tnm(33) = "LAM2"
    tnm(34) = "L2"
    tnm(35) = "T2"
    tnm(36) = "S2"
    tnm(37) = "R2"
    tnm(38) = "K2"
    tnm(39) = "MSN2"
    tnm(40) = "KJ2"
    tnm(41) = "2SM2"
    tnm(42) = "MO3"
    tnm(43) = "M3"
    tnm(44) = "SO3"
    tnm(45) = "MK3"
    tnm(46) = "SK3"
    tnm(47) = "MN4"
    tnm(48) = "M4"
    tnm(49) = "SN4"
    tnm(50) = "MS4"
    tnm(51) = "MK4"
    tnm(52) = "S4"
    tnm(53) = "SK4"
    tnm(54) = "2MN6"
    tnm(55) = "M6"
    tnm(56) = "MSN6"
    tnm(57) = "2MS6"
    tnm(58) = "2MK6"
    tnm(59) = "2SM6"
    tnm(60) = "MSK6"

'定数チェック 2008/01/07 09:50 MK
     sp(1) = 0.0410686
     sp(2) = 0.0821373
     sp(3) = 0.5443747
     sp(4) = 1.0158958
     sp(5) = 1.0980331
     sp(6) = 12.8542862
     sp(7) = 12.9271398
     sp(8) = 13.3986609
     sp(9) = 13.4715145
    sp(10) = 13.9430356
    sp(11) = 14.0251729
    sp(12) = 14.4920521
    sp(13) = 14.5695476
    sp(14) = 14.9178647
    sp(15) = 14.9589314
    sp(16) = 15#
    sp(17) = 15.0410686
    sp(18) = 15.0821353
    sp(19) = 15.1232059
    sp(20) = 15.5125897
    sp(21) = 15.5854433
    sp(22) = 16.0569644
    sp(23) = 16.1391017
    sp(24) = 27.3416964
    sp(25) = 27.4238337
    sp(26) = 27.8953548
    sp(27) = 27.9682084
    sp(28) = 28.4397295
    sp(29) = 28.5125831
    sp(30) = 28.9019669
    sp(31) = 28.9841042
    sp(32) = 29.0662415
    sp(33) = 29.4556253
    sp(34) = 29.5284789
    sp(35) = 29.9589333
    sp(36) = 30#
    sp(37) = 30.0410667
    sp(38) = 30.0821373
    sp(39) = 30.5443747
    sp(40) = 30.626512
    sp(41) = 31.0158958
    sp(42) = 42.9271398
    sp(43) = 43.4761563
    sp(44) = 43.9430356
    sp(45) = 44.0251729
    sp(46) = 45.0410686
    sp(47) = 57.4238337
    sp(48) = 57.9682084
    sp(49) = 58.4397295
    sp(50) = 58.9841042
    sp(51) = 59.0662415
    sp(52) = 60#
    sp(53) = 60.0821373
    sp(54) = 86.407938
    sp(55) = 86.9523127
    sp(56) = 87.4238337
    sp(57) = 87.9682084
    sp(58) = 88.0503457
    sp(59) = 88.9841042
    sp(60) = 89.0662415

    For i = 1 To 5
        tt(i) = 0#
    Next i
    For i = 6 To 23
        tt(i) = 1#
    Next i
    For i = 24 To 41
        tt(i) = 2#
    Next i
    For i = 42 To 46
        tt(i) = 3#
    Next i
    For i = 47 To 53
        tt(i) = 4#
    Next i
    For i = 54 To 60
        tt(i) = 6#
    Next i

    mfday(1) = -1
    mfday(2) = 30
    mfday(3) = 58
    mfday(4) = 89
    mfday(5) = 119
    mfday(6) = 150
    mfday(7) = 180
    mfday(8) = 211
    mfday(9) = 242
    mfday(10) = 272
    mfday(11) = 303
    mfday(12) = 333

End Sub
Public Function TideNumber(ByVal code) As Integer

    If code = "C0" Then TideNumber = 0
    If code = "SA" Then TideNumber = 1
    If code = "SSA" Then TideNumber = 2
    If code = "MM" Then TideNumber = 3
    If code = "MSF" Then TideNumber = 4
    If code = "MF" Then TideNumber = 5
    If code = "2Q1" Then TideNumber = 6
    If code = "SGM1" Then TideNumber = 7
    If code = "Q1" Then TideNumber = 8
    If code = "RHO1" Then TideNumber = 9
    If code = "O1" Then TideNumber = 10
    If code = "MP1" Then TideNumber = 11
    If code = "M1" Then TideNumber = 12
    If code = "CHI1" Then TideNumber = 13
    If code = "PAI1" Then TideNumber = 14
    If code = "P1" Then TideNumber = 15
    If code = "S1" Then TideNumber = 16
    If code = "K1" Then TideNumber = 17
    If code = "PSI1" Then TideNumber = 18
    If code = "PHI1" Then TideNumber = 19
    If code = "THT1" Then TideNumber = 20
    If code = "J1" Then TideNumber = 21
    If code = "SO1" Then TideNumber = 22
    If code = "OO1" Then TideNumber = 23
    If code = "OQ2" Then TideNumber = 24
    If code = "MNS2" Then TideNumber = 25
    If code = "2N2" Then TideNumber = 26
    If code = "MU2" Then TideNumber = 27
    If code = "N2" Then TideNumber = 28
    If code = "NU2" Then TideNumber = 29
    If code = "OP2" Then TideNumber = 30
    If code = "M2" Then TideNumber = 31
    If code = "MKS2" Then TideNumber = 32
    If code = "LAM2" Then TideNumber = 33
    If code = "L2" Then TideNumber = 34
    If code = "T2" Then TideNumber = 35
    If code = "S2" Then TideNumber = 36
    If code = "R2" Then TideNumber = 37
    If code = "K2" Then TideNumber = 38
    If code = "MSN2" Then TideNumber = 39
    If code = "KJ2" Then TideNumber = 40
    If code = "2SM2" Then TideNumber = 41
    If code = "MO3" Then TideNumber = 42
    If code = "M3" Then TideNumber = 43
    If code = "SO3" Then TideNumber = 44
    If code = "MK3" Then TideNumber = 45
    If code = "SK3" Then TideNumber = 46
    If code = "MN4" Then TideNumber = 47
    If code = "M4" Then TideNumber = 48
    If code = "SN4" Then TideNumber = 49
    If code = "MS4" Then TideNumber = 50
    If code = "MK4" Then TideNumber = 51
    If code = "S4" Then TideNumber = 52
    If code = "SK4" Then TideNumber = 53
    If code = "2MN6" Then TideNumber = 54
    If code = "M6" Then TideNumber = 55
    If code = "MSN6" Then TideNumber = 56
    If code = "2MS6" Then TideNumber = 57
    If code = "2MK6" Then TideNumber = 58
    If code = "2SM6" Then TideNumber = 59
    If code = "MSK6" Then TideNumber = 60

End Function
Public Sub getv60(ByRef vg() As Double, ByRef S As Double, ByRef H As Double, ByRef p As Double)
'
'                                              modified 2006 Dec.1
'                          based on TIDAL HARMONIC CONSTANTS TABLES
'                                             (PUB.No.742 1992 Feb)
'
    Dim A2(60)   As Double
    Dim a3(60)   As Double
    Dim a4(60)   As Double
    Dim c(60)    As Double
    Dim i        As Long
'                                                      数値チェックOK 2007/01/09 13:40
     A2(1) = 0#:    a3(1) = 1#:    a4(1) = 0#:    c(1) = 0#
     A2(2) = 0#:    a3(2) = 2#:    a4(2) = 0#:    c(2) = 0#
     A2(3) = 1#:    a3(3) = 0#:    a4(3) = -1#:   c(3) = 0#
     A2(4) = 2#:    a3(4) = -2#:   a4(4) = 0#:    c(4) = 0#
     A2(5) = 2#:    a3(5) = 0#:    a4(5) = 0#:    c(5) = 0#
     A2(6) = -4#:   a3(6) = 1#:    a4(6) = 2#:    c(6) = 270#
     A2(7) = -4#:   a3(7) = 3#:    a4(7) = 0#:    c(7) = 270#
     A2(8) = -3#:   a3(8) = 1#:    a4(8) = 1#:    c(8) = 270#
     A2(9) = -3#:   a3(9) = 3#:    a4(9) = -1#:   c(9) = 270#
    A2(10) = -2#:  a3(10) = 1#:   a4(10) = 0#:   c(10) = 270#
    A2(11) = -2#:  a3(11) = 3#:   a4(11) = 0#:   c(11) = 90#
    A2(12) = -1#:  a3(12) = 1#:   a4(12) = 0#:   c(12) = 90#
    A2(13) = -1#:  a3(13) = 3#:   a4(13) = -1#:  c(13) = 90#
    A2(14) = 0#:   a3(14) = -2#:  a4(14) = 0#:   c(14) = 193#
    A2(15) = 0#:   a3(15) = -1#:  a4(15) = 0#:   c(15) = 270#
    A2(16) = 0#:   a3(16) = 0#:   a4(16) = 0#:   c(16) = 180#
    A2(17) = 0#:   a3(17) = 1#:   a4(17) = 0#:   c(17) = 90#
    A2(18) = 0#:   a3(18) = 2#:   a4(18) = 0#:   c(18) = 167#
    A2(19) = 0#:   a3(19) = 3#:   a4(19) = 0#:   c(19) = 90#
    A2(20) = 1#:   a3(20) = -1#:  a4(20) = 1#:   c(20) = 90#
    A2(21) = 1#:   a3(21) = 1#:   a4(21) = -1#:  c(21) = 90#
    A2(22) = 2#:   a3(22) = -1#:  a4(22) = 0#:   c(22) = 90#
    A2(23) = 2#:   a3(23) = 1#:   a4(23) = 0#:   c(23) = 90#
    A2(24) = -5#:  a3(24) = 2#:   a4(24) = 1#:   c(24) = 180#
    A2(25) = -5#:  a3(25) = 4#:   a4(25) = 1#:   c(25) = 0#
    A2(26) = -4#:  a3(26) = 2#:   a4(26) = 2#:   c(26) = 0#
    A2(27) = -4#:  a3(27) = 4#:   a4(27) = 0#:   c(27) = 0#
    A2(28) = -3#:  a3(28) = 2#:   a4(28) = 1#:   c(28) = 0#
    A2(29) = -3#:  a3(29) = 4#:   a4(29) = -1#:  c(29) = 0#
    A2(30) = -2#:  a3(30) = 0#:   a4(30) = 0#:   c(30) = 180#
    A2(31) = -2#:  a3(31) = 2#:   a4(31) = 0#:   c(31) = 0#
    A2(32) = -2#:  a3(32) = 4#:   a4(32) = 0#:   c(32) = 0#
    A2(33) = -1#:  a3(33) = 0#:   a4(33) = 1#:   c(33) = 180#
    A2(34) = -1#:  a3(34) = 2#:   a4(34) = -1#:  c(34) = 180#
    A2(35) = 0#:   a3(35) = -1#:  a4(35) = 0#:   c(35) = 283#
    A2(36) = 0#:   a3(36) = 0#:   a4(36) = 0#:   c(36) = 0#
    A2(37) = 0#:   a3(37) = 1#:   a4(37) = 0#:   c(37) = 257#
    A2(38) = 0#:   a3(38) = 2#:   a4(38) = 0#:   c(38) = 0#
    A2(39) = 1#:   a3(39) = 0#:   a4(39) = -1#:  c(39) = 0#
    A2(40) = 1#:   a3(40) = 2#:   a4(40) = -1#:  c(40) = 180#
    A2(41) = 2#:   a3(41) = -2#:  a4(41) = 0#:   c(41) = 0#
    A2(42) = -4#:  a3(42) = 3#:   a4(42) = 0#:   c(42) = 270#
    A2(43) = -3#:  a3(43) = 3#:   a4(43) = 0#:   c(43) = 180#
    A2(44) = -2#:  a3(44) = 1#:   a4(44) = 0#:   c(44) = 270#
    A2(45) = -2#:  a3(45) = 3#:   a4(45) = 0#:   c(45) = 90#
    A2(46) = 0#:   a3(46) = 1#:   a4(46) = 0#:   c(46) = 90#
    A2(47) = -5#:  a3(47) = 4#:   a4(47) = 1#:   c(47) = 0#
    A2(48) = -4#:  a3(48) = 4#:   a4(48) = 0#:   c(48) = 0#
    A2(49) = -3#:  a3(49) = 2#:   a4(49) = 1#:   c(49) = 0#
    A2(50) = -2#:  a3(50) = 2#:   a4(50) = 0#:   c(50) = 0#
    A2(51) = -2#:  a3(51) = 4#:   a4(51) = 0#:   c(51) = 0#
    A2(52) = 0#:   a3(52) = 0#:   a4(52) = 0#:   c(52) = 0#
    A2(53) = 0#:   a3(53) = 2#:   a4(53) = 0#:   c(53) = 0#
    A2(54) = -7#:  a3(54) = 6#:   a4(54) = 1#:   c(54) = 0#
    A2(55) = -6#:  a3(55) = 6#:   a4(55) = 0#:   c(55) = 0#
    A2(56) = -5#:  a3(56) = 4#:   a4(56) = 1#:   c(56) = 0#
    A2(57) = -4#:  a3(57) = 4#:   a4(57) = 0#:   c(57) = 0#
    A2(58) = -4#:  a3(58) = 6#:   a4(58) = 0#:   c(58) = 0#
    A2(59) = -2#:  a3(59) = 2#:   a4(59) = 0#:   c(59) = 0#
    A2(60) = -2#:  a3(60) = 4#:   a4(60) = 0#:   c(60) = 0#


    For i = 1 To 60
        vg(i) = A2(i) * S + a3(i) * H + a4(i) * p + c(i)
        vg(i) = (vg(i) + 3600#) Mod 360#
    Next i


End Sub
Public Function Pythag(ByVal a As Double, ByVal b As Double) As Double

    Dim absa As Double
    Dim absb As Double

    absa = Abs(a)
    absb = Abs(b)
    If absa > absb Then
        Pythag = absa * Sqr(1# + (absb / absa) ^ 2#)
    Else
        If absb = 0# Then
            Pythag = 0#
        Else
            Pythag = absb * Sqr(1# + (absa / absb) ^ 2#)
        End If
    End If

End Function
Public Sub getu60(ByRef ug() As Double, ByRef p As Double, ByRef fn As Double, ByRef s1 As Double, ByRef s2 As Double, ByRef s3 As Double, ByRef a12 As Double, ByRef b12 As Double, ByRef a34 As Double, ByRef b34 As Double)
'
'                                              modified 1996 Feb.12
'                          based on TIDAL HARMONIC CONSTANTS TABLES
'                                             (PUB.No.742 1992 Feb)
'
    Dim ug5       As Double
    Dim ug10      As Double
    Dim ug12      As Double
    Dim ug17      As Double
    Dim ug21      As Double
    Dim ug23      As Double
    Dim ug31      As Double
    Dim ug34      As Double
    Dim ug38      As Double

    ' --- u of basic tide ( Mf,O1,K1,J1,OO1,M2,K2 )
    ug5 = -23.74 * s1 + 2.68 * s2 - 0.38 * s3
    ug10 = 10.8 * s1 - 1.34 * s2 + 0.19 * s3
    ug17 = -8.86 * s1 + 0.68 * s2 - 0.07 * s3
    ug21 = -12.94 * s1 + 1.34 * s2 - 0.19 * s3
    ug23 = -36.68 * s1 + 4.02 * s2 - 0.57 * s3
    ug31 = -2.14 * s1
    ug38 = -17.74 * s1 + 0.68 * s2 - 0.04 * s3

    '    - u of M1 -
    a12 = 2# * Cos(p * rad) + 0.4 * Cos((p - fn) * rad)
    b12 = Sin(p * rad) + 0.2 * Sin((p - fn) * rad)
    ug12 = Atan2(b12, a12) / rad

    '   - u of L2 -
    a34 = 1# - 0.2505 * Cos(2# * p * rad) _
              - 0.1102 * Cos((2# * p - fn) * rad) _
              - 0.0156 * Cos((2# * p - 2# * fn) * rad) _
              - 0.037 * Cos(fn * rad)

    b34 = -0.2505 * Sin(2# * p * rad) _
          - 0.1102 * Sin((2# * p - fn) * rad) _
          - 0.0156 * Sin((2# * p - 2# * fn) * rad) _
          - 0.037 * Sin(fn * rad)
    ug34 = Atan2(b34, a34) / rad

    ug(1) = 0#
    ug(2) = 0#
    ug(3) = 0#
    ug(4) = -ug31
    ug(5) = ug5
    ug(6) = ug10
    ug(7) = ug10
    ug(8) = ug10
    ug(9) = ug10
    ug(10) = ug10

    ug(11) = ug31
    ug(12) = ug12
    ug(13) = ug21
    ug(14) = 0#
    ug(15) = 0#
    ug(16) = 0#
    ug(17) = ug17
    ug(18) = 0#
    ug(19) = 0#
    ug(20) = ug21

    ug(21) = ug21
    ug(22) = ug21
    ug(23) = ug23
    ug(24) = ug10 + ug(8)
    ug(25) = ug31 * 2#
    ug(26) = ug31
    ug(27) = ug31
    ug(28) = ug31
    ug(29) = ug31
    ug(30) = ug10 + ug(15)
    ug(31) = ug31
    ug(32) = ug31 + ug38
    ug(33) = ug31
    ug(34) = ug34
    ug(35) = 0#
    ug(36) = 0#
    ug(37) = 0#
    ug(38) = ug38
    ug(39) = ug31 * 2#
    ug(40) = ug17 + ug21

    ug(41) = -ug31
    ug(42) = ug31 + ug10
    ug(43) = ug31 * 1.5
    ug(44) = ug10
    ug(45) = ug31 + ug17
    ug(46) = ug17
    ug(47) = ug31 * 2#
    ug(48) = ug31 * 2#
    ug(49) = ug31
    ug(50) = ug31

    ug(51) = ug31 + ug38
    ug(52) = 0#
    ug(53) = ug38
    ug(54) = ug31 * 3#
    ug(55) = ug31 * 3#
    ug(56) = ug31 * 2#
    ug(57) = ug31 * 2#
    ug(58) = ug31 * 2# + ug38
    ug(59) = ug31
    ug(60) = ug31 + ug38

End Sub
Sub getf60(ByRef F, ByRef C1, ByRef C2, ByRef C3, ByRef a12, ByRef b12, ByRef a34, ByRef b34)

'
'                                              modified 2006 Dec.1
'                          based on TIDAL HARMONIC CONSTANTS TABLES
'                                             (PUB.No.742 1992 Feb)
'
    Dim F3    As Double
    Dim f5    As Double
    Dim f10   As Double
    Dim f12   As Double
    Dim f17   As Double
    Dim f21   As Double
    Dim f23   As Double
    Dim f31   As Double
    Dim f34   As Double
    Dim f38   As Double

    ' --- f of basic tide ( Mm,Mf,O1,K1,J1,OO1,M2,K2 )
    F3 = 1# - 0.13 * C1 + 0.0013 * C2
    f5 = 1.0429 + 0.4135 * C1 - 0.004 * C2
    f10 = 1.0089 + 0.1871 * C1 - 0.0147 * C2 + 0.0014 * C3
    f17 = 1.006 + 0.115 * C1 - 0.0088 * C2 + 0.0006 * C3
    f21 = 1.0129 + 0.1676 * C1 - 0.017 * C2 + 0.0016 * C3
    f23 = 1.1027 + 0.6504 * C1 + 0.0317 * C2 - 0.0014 * C3
    f31 = 1.0004 - 0.0373 * C1 + 0.0002 * C2
    f38 = 1.0241 + 0.2863 * C1 + 0.0083 * C2 - 0.0015 * C3

    '   - f of M1 and L2
    f12 = Pythag(a12, b12)
    f34 = Pythag(a34, b34)

    F(1) = 1#
    F(2) = 1#
    F(3) = F3
    F(4) = f31
    F(5) = f5
    F(6) = f10
    F(7) = f10
    F(8) = f10
    F(9) = f10
    F(10) = f10

    F(11) = f31
    F(12) = f12
    F(13) = f21
    F(14) = 1#
    F(15) = 1#
    F(16) = 1#
    F(17) = f17
    F(18) = 1#
    F(19) = 1#
    F(20) = f21

    F(21) = f21
    F(22) = f21
    F(23) = f23
    F(24) = f10 * F(8)
    F(25) = f31 * f31
    F(26) = f31
    F(27) = f31
    F(28) = f31
    F(29) = f31
    F(30) = f10 * F(15)

    F(31) = f31
    F(32) = f31 * f38
    F(33) = f31
    F(34) = f34
    F(35) = 1#
    F(36) = 1#
    F(37) = 1#
    F(38) = f38
    F(39) = f31 * f31
    F(40) = f17 * f21

    F(41) = f31
    F(42) = f31 * f10
    F(43) = Sqr(f31) * Sqr(f31) * Sqr(f31)
    F(44) = f10
    F(45) = f31 * f17
    F(46) = f17
    F(47) = f31 * f31
    F(48) = f31 * f31
    F(49) = f31
    F(50) = f31

    F(51) = f31 * f38
    F(52) = 1#
    F(53) = f38
    F(54) = f31 * f31 * f31
    F(55) = f31 * f31 * f31
    F(56) = f31 * f31
    F(57) = f31 * f31
    F(58) = f31 * f31 * f38
    F(59) = f31
    F(60) = f31 * f38

End Sub
Public Sub get60_2(ByVal iy As Integer, ByVal im As Integer, ByVal id As Integer, ByVal ih As Integer, ByVal MM As Integer, ByRef F() As Double)

    Dim d      As Double
    Dim dfl    As Double
    Dim y      As Double
    Dim S      As Double
    Dim H      As Double
    Dim p      As Double
    Dim fn     As Double
    Dim C1     As Double
    Dim C2     As Double
    Dim C3     As Double
    Dim s1     As Double
    Dim s2     As Double
    Dim s3     As Double
    Dim a12    As Double
    Dim b12    As Double
    Dim a34    As Double
    Dim b34    As Double
    Dim leap   As Long
    Dim L      As Long


'うるう年の判別　サンプルはうそ
'    leap = 0
'    If iy Mod 4 = 0 And iy Mod 100 = 0 Then leap = 1
'    If iy Mod 400 = 0 Then leap = 1
'うるう年の判別　菊地修正
    leap = 0
    If iy Mod 4 = 0 Then
        leap = 1
    End If
    If iy Mod 100 = 0 Then
       leap = 0
    End If
    If iy Mod 400 = 0 Then
        leap = 1
    End If
    d = mfday(im) + id
    If leap = 1 And im >= 3 Then d = d + 1#
    L = Int((iy + 3) / 4) - 500
    dfl = d + CDbl(L)
    y = CDbl(iy)

    S = 211.728 + 129.38471 * (y - 2000#) + 13.176396 * dfl
    H = 279.974 - 0.23871 * (y - 2000#) + 0.985647 * dfl
    p = 83.298 + 40.66229 * (y - 2000#) + 0.111404 * dfl
    fn = 125.071 - 19.32812 * (y - 2000#) - 0.052954 * dfl

    C1 = Cos(fn * rad)
    C2 = Cos(2# * fn * rad)
    C3 = Cos(3# * fn * rad)
    s1 = Sin(fn * rad)
    s2 = Sin(2# * fn * rad)
    s3 = Sin(3# * fn * rad)

    ' - these values are valid even after 2000 -----------

    a12 = 2# * Cos(p * rad) + 0.4 * Cos((p - fn) * rad)
    b12 = Sin(p * rad) + 0.2 * Sin((p - fn) * rad)

    '   - u of L2 -
    a34 = 1# - 0.2505 * Cos(2# * p * rad) _
              - 0.1102 * Cos((2# * p - fn) * rad) _
              - 0.0156 * Cos((2# * p - 2# * fn) * rad) _
              - 0.037 * Cos(fn * rad)

    b34 = -0.2505 * Sin(2# * p * rad) _
          - 0.1102 * Sin((2# * p - fn) * rad) _
          - 0.0156 * Sin((2# * p - 2# * fn) * rad) _
          - 0.037 * Sin(fn * rad)

    getf60 F, C1, C2, C3, a12, b12, a34, b34

End Sub
Sub dumy()
'Public Function sin(ByVal x As Double) As Double
'
'    sin = Math.sin(x)
'
'End Function
'Public Function cos(ByVal x As Double) As Double
'
'    cos = Math.cos(x)
'
'End Function
'Public Function sqrt(ByVal x As Double) As Double
'
'    sqrt = Math.sqrt(x)
'
'End Function
'Public Function atan2(ByVal y As Double, ByVal x As Double) As Double
'
'   Dim f  As Double
'
'    f = Math.Atan(y / x) * 180 / Math.PI
'    If x >= 0 And y >= 0 Then
'    ElseIf x < 0 And y > 0 Then
'        f = f + 180#
'    ElseIf x < 0 And y <= 0 Then
'        f = f + 180#
'    ElseIf x >= 0 And y < 0 Then
'        f = f + 360#
'    End If
'
'    atan2 = f
'
'End Function
End Sub
Public Function Atan2(y, x)
'
' Atan2 For Visual Basic
'
    If x = 0# Then
        Atan2 = IIf(y <= 0#, 1.570796327, -1.570796327)
    Else
        If x > 0# Then
            Atan2 = Atn(y / x)
        Else
            Atan2 = IIf(y <= 0#, Atn(y / x) - 3.1415926535, Atn(y / x) + 3.1415926535)
        End If
    End If

End Function
