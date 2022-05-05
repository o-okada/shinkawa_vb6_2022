Attribute VB_Name = "Tidal"
Option Explicit
Option Base 1
Public Ti(10000)    As Single
Public Ti_Time      As Date    '天文潮位データの開始時刻

'************************************
'
'気象庁予測潮位を読む
'
'
'
'************************************
Sub Pre_Tidal(d1 As Date)

    Dim i    As Long
    Dim j    As Long
    Dim k    As Long
    Dim nf   As Long
    Dim buf  As String
    Dim T    As Single
    Dim File As String

    buf = Trim(str(Year(d1)))
    File = App.Path & "\DATA\" & buf & "名古屋天文潮位.txt"

    If Len(Dir(File)) = 0 Then
       isRAIN = "01"
       jgd = Now
       ORA_Message_Out "天文潮位表読み込み", File & " のファイルが有りません、計算を終了します。早急にデータを入力・保存してください。", 2
       End
    End If

    nf = FreeFile
    Open File For Input As #nf

    Line Input #nf, buf
    Ti_Time = CDate(Mid(buf, 1, 20))
    j = 0
    Do Until EOF(nf)
        Line Input #nf, buf
        For i = 1 To 24
            T = CSng(Mid(buf, 3 * (i - 1) + 1, 3)) * 0.01
            j = j + 1
            Ti(j) = T
        Next i
    Loop

    Close #nf

End Sub
'*************************************************
'
'潮位予測
'
'Timet=現時刻
'
'
'*************************************************
Sub TidalY(Timet As Date, C0 As Single, C1 As Single, C2 As Single, C3 As Single)

    Dim i        As Long
    Dim tw       As Date
    Dim tt(4)    As Single
    Dim cc(5)    As Single
    Dim Cminute  As Integer

    If Timet < "2002/01/01 00:00" Then Exit Sub

    LOG_Out "IN   TidalY  現時刻=" & Format(Timet, "yyyy/mm/dd/hh:nn")


    i = DateDiff("h", Ti_Time, Timet) + 1


    tw = Format(Timet, "yyyy/mm/dd hh:00")


    cc(1) = Ti(i)
    cc(2) = Ti(i + 1)
    cc(3) = Ti(i + 2)
    cc(4) = Ti(i + 3)
    cc(5) = Ti(i + 4)
    Cminute = Minute(Timet)

    Tide_InterP cc(), tw, Cminute, tt()

    C0 = tt(1)
    C1 = tt(2)
    C2 = tt(3)
    C3 = tt(4)
'---------------- 対気象庁演習用 2002年 朔望平均満潮位 ------
'    c0 = 1.118
'    c1 = 1.118
'    c2 = 1.118
'    c3 = 1.118
'----------------------------------------------------------
    LOG_Out "OUT   TidalY dw=" & TIMEC(Timet) & "  C0=" & fmt(C0)
    LOG_Out "                                      C1=" & fmt(C1)
    LOG_Out "                                      C2=" & fmt(C2)
    LOG_Out "                                      C3=" & fmt(C3)

    LOG_Out "OUT  TidalY "

End Sub
'*********************************************************
'天文潮位データを内挿する。
'Tid(4) -- 天文潮位データ４個入力
'Tide_Time Tid(1)の時刻これは必ず正時で分は０分です。入力
''CMinute  xx分入力
'
'T(1) = Tid(1)のCMinute分の値 出力
'T(2) = Tid(2)のCMinute分の値 出力
'T(3) = Tid(3)のCMinute分の値 出力
'T(4) = Tid(4)のCMinute分の値 出力
'
'*********************************************************
Sub Tide_InterP(Tid() As Single, Tide_Time As Date, Cminute As Integer, T() As Single)

    Dim cnt  As Long
    Dim Tida As Single
    Dim Tidb As Single

    For cnt = 1 To 4

        Tida = Tid(cnt)
        Tidb = Tid(cnt + 1)
       
        T(cnt) = (Tida * (60 - Cminute) + Tidb * Cminute) / 60

    Next cnt

End Sub
