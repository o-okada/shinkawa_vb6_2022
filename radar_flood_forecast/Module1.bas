Attribute VB_Name = "Module1"
'******************************************************************************
'モジュール名：Module1
'
'******************************************************************************
Option Explicit
Option Base 1

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'******************************************************************************
'
'******************************************************************************
Public Run_Mode             As Boolean                                  'False=デバッグモード時  True=通常モード時
Public Current_Path         As String                                   'カレントフォルダのパス(D:\Shinkaw\ のような)
Public DBX_ora              As Boolean                                  'true=結果をオラクルに書く false=結果をオラクルに書かない
Public Log_Repo             As Long
'******************************************************************************
'
'******************************************************************************
Public Flood_Name           As String                                   '洪水データタイトル
Public Now_Step             As Integer                                  '計算開始時刻から現時刻までのステップ数
Public Const Yosoku_Step = 3
Public All_Step             As Integer                                  '計算開始時刻から予測終了までのステップ数
Public Data_Steps           As Integer                                  '入力データステップ数
Public js(6)                As Integer                                  'データ開始時刻
Public jg(6)                As Integer                                  'データ終了時刻
Public jx(6)                As Integer                                  '時刻ワーク
Public jsd                  As Date                                     'js()の日付型
Public jgd                  As Date                                     'jg()の日付型
Public jxd                  As Date                                     'jx()の日付型
Public NSK_jsd              As Date                                     '不定流計算開始時刻
Public Data_Pich(5)         As Single
'******************************************************************************
'
'******************************************************************************
Public Return_Code_fm       As Boolean                                  '
'******************************************************************************
'
'******************************************************************************
Public Log_Num              As Integer                                  'ログファイルに書き出すファイル番号
Public Log_Time             As Date                                     'ログファイルをクリアするための時刻管理用
'******************************************************************************
'
'******************************************************************************
Public Const nd = 74 '215                                               '断面数
Public Const aksk = -99#                                                '欠測定数
'******************************************************************************
'
'******************************************************************************
Public Y_DS                 As Date                                     'データ開始時刻
Public Y_DE                 As Date                                     'データ終了時刻
Public NDN()                As String                                   '断面記号
Public NT                   As Integer                                  '計算時間数
Public HQ()                 As Single                                   '予測計算結果
Public DX()                 As Single                                   '区間距離
Public sdx()                As Single                                   '累加区間距離
Public MDX                  As Single                                   '河道長
Public MAX_H()              As Single                                   '最大水位
'******************************************************************************
'
'******************************************************************************
Public Input_file           As String                                   '既往洪水データのファイル名
'******************************************************************************
'
'******************************************************************************
Public Initial_HQ(2, nd)    As Single                                   '計算開始水面形
Public Tide(500)            As Single                                   '下流端境界条件
Public Const Hnum = 7                                                   '水位観測所数
Public Name_H(10)           As String                                   '水位観測所名
Public HO(10, 500)          As Single                                   '観測所水位
Public HO_Title             As String                                   '水位タイトル
Public Name_R(10)           As String                                   '雨量観測所名
Public RO(10, 500)          As Single                                   '基準地点上流流域雨量
Public RO_Title             As String                                   '雨量タイトル
Public Const Rnum = 10                                                  '雨量観測所数
Public Wpath                As String                                   'データ出力ディレクトリパス
Public HQA(2)               As Single                                   'H-Q式のA
Public HQB(2)               As Single                                   'H-Q式のB
'******************************************************************************
'
'******************************************************************************
Public V_Sec_Name(10, 2)    As String                                   '検証地点名  1=計算断面名  2=観測所名
Public V_Sec_Num(10)        As Integer                                  '検証地点断面順番号
Public V_Sec_Cnt            As Integer                                  '検証地点数
Public Froude               As Single                                   '不定流計算結果の基準断面の平均フルード数
'******************************************************************************
'
'******************************************************************************
Public H_Scale(5, 3)        As Single                                   '水位目盛り  1=下目盛り 2=上目盛り 3=ピッチ
Public Q_Scale(5, 3)        As Single                                   '流量目盛り  1=下目盛り 2=上目盛り 3=ピッチ
Public H_Stand1(5, 5)       As Single                                   '基準地点、水位＝( 1=ＨＷＬ 2=第三基準  3=第二基準 4=第一基準 5=ゼロ点高 )
Public H_Stand1t(5, 5)      As String                                   '基準地点水位名称
Public H_Stand2(5, 3)       As Single                                   '基準地点、ポンプ水位＝( 1=停止水位 2=再開水位 3=準備水位 )
Public H_Stand2t(5, 3)      As String                                   'ポンプ水位名称
Public H_Standi(5, 2)       As Integer                                  '基準地点、基準水位数＝( 1=基準 2=ポンプ )
'******************************************************************************
'
'******************************************************************************
Public OBS1                 As Boolean                                  '実績データ全部プロットするとき=True
Public CAL1                 As Integer                                  '1=計算値プロット有り  0=計算値プロット無し
'******************************************************************************
'不等流計算等パラメータ
'******************************************************************************
Public Q_kuji               As Single                                   '不流計算用久地野初期流量
Public Q_Haru               As Single                                   '不等流計算用春日初期流量
Public H_Sea                As Single                                   '不等流計算用下流端水位
'******************************************************************************
'
'******************************************************************************
Public Nonuni_H(5, 0 To 3)  As Single                                   '不等流計算による予測結果
Public CO(5, 4)             As Single                                   '不定流計算生値
Public CF(5, 0 To 3)        As Single                                   '不定流計算フィードバック後
'******************************************************************************
'フィードバック補正値
'******************************************************************************
Public Slide1(5)            As Single                                   '現時刻前２時間の１回目スライド量
Public Slide2(5)            As Single                                   '現時刻　　　　の２回目スライド量
Public Delta_H(5)           As Single                                   '１時間当りの水位補正値（２時間後は２倍３時間後は３倍）
Public OBS_Pump             As Boolean                                  '実績ポンプを計算に使う時True
Public Beer                 As Boolean                                  '不定流計算生値ハイドロプロット時TRUE
Public Un_Cal               As Single                                   '水場の水位が [ UN_Cal ] 以下の時は不定流ではなく不等流とする。
Public Category             As Boolean                                  'True=不定流で計算されたとき  False=不等流で計算されたとき
Public Error_Message()      As String                                   'オラクルに出力するエラーメッセージ
Public Error_Message_n      As Long                                     'オラクルに出力するエラーメッセージの数
Public Error_Cal_Type       As String                                   'Error検出時の雨量タイプ

'******************************************************************************
'サブルーチン：Cal_Initial_flow_profile(irc As Boolean)
'処理概要：
'不等流計算により不定流計算用の初期水面形の計算を行う。
'******************************************************************************
Sub Cal_Initial_flow_profile(irc As Boolean)
    Dim i   As Integer
    Dim i1  As Integer
    Dim nf  As Integer
    Dim buf As String
    nf = FreeFile
    Open Wpath & "\Non_Flow.log" For Output As #nf
    Print #nf, " 境界条件　久地野流入= " & Format(Q_kuji, "###0.00") & _
               "   春日流入= " & Format(Q_Haru, "###0.00") & _
               "   下流端水位= " & Format(H_Sea, "##0.000")
    '******************************************************
    '
    '******************************************************
    QU = Q_kuji + Q_Haru
    H_Start = H_Sea
    Start_Sec = "S0.000"
    End_Sec = "S12.40"
    Nonuniform_Flow irc
    If irc = False Then GoTo jump1
    i1 = Start_Num
    '******************************************************
    '
    '******************************************************
    QU = Q_kuji
    H_Start = ch(End_Num)
    Start_Sec = "S12.40"
    End_Sec = "S20.00"
    Nonuniform_Flow irc
    If irc = False Then GoTo jump1
    '******************************************************
    '
    '******************************************************
    H_Start = ch(Start_Num)
    QU = Q_Haru
    Start_Sec = "G0.000"
    End_Sec = "G8.200"
    Nonuniform_Flow irc
    If irc = False Then GoTo jump1
    '******************************************************
    '
    '******************************************************
    Print #nf, "    N   断面       H         A         Q       V       FR     FLAG"
    For i = i1 To End_Num
        buf = Format(Format(i, "####0"), "@@@@@  ") & Sec_Name(i)
        buf = buf & Format(Format(ch(i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CA(1, i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CQ(1, i), "###0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CV(i), "###0.000"), "@@@@@@@@")
        buf = buf & Format(Format(FR(i), "###0.000"), "@@@@@@@@")
        buf = buf & Space(5) & CFLAG(i)
        Print #nf, buf
    Next i
    Close #nf
    '******************************************************
    '
    '******************************************************
    nf = FreeFile
    Open Wpath & "\NSK_初期水面形.Temp" For Output As #nf
    Print #nf, " INITIAL        H          Q"
    For i = i1 To End_Num
        buf = Space(4) & Sec_Name(i)
        buf = buf & Format(Format(ch(i), "#####0.000"), "@@@@@@@@@@")
        buf = buf & Format(Format(CQ(1, i), "######0.00"), "@@@@@@@@@@")
        Print #nf, buf
    Next i
    Close #nf
    Exit Sub
jump1:
    Log_Calc "不定流計算の初期水面形の計算に失敗した。"
    Close #nf
    Exit Sub
End Sub

'******************************************************************************
'サブルーチン：Constant_Read()
'処理概要：
'******************************************************************************
Sub Constant_Read()
    Dim i     As Integer
    Dim j     As Integer
    Dim nf    As Integer
    Dim buf   As String
    Dim msg   As String
    LOG_Out "IN  Constant_Read"
    On Error GoTo ERH
    For i = 1 To 5
        For j = 1 To 5
            H_Stand1(i, j) = 99#
        Next j
        For j = 1 To 3
            H_Stand2(i, j) = 99#
        Next j
    Next i
    msg = "DATAフォルダに基準水位のファイルがない"
    nf = FreeFile
    Open App.Path & "\DATA\基準水位.DAT" For Input As #nf
    Line Input #nf, buf                                                 'タイトル読み飛ばし
    '******************************************************
    '基準水位読み込み
    '******************************************************
    For i = 1 To 5
        Select Case i
            Case 1
            msg = "下之一色："
            Case 2
            msg = "大　治："
            Case 3
            msg = "水場川外水位："
            Case 4
            msg = "久地野："
            Case 5
            msg = "春　日："
        End Select
        Line Input #nf, buf
            H_Standi(i, 1) = CInt(Mid(buf, 1, 5))                       '基準水位の数
            H_Standi(i, 2) = CInt(Mid(buf, 6, 5))                       'ポンプ水位の数
        For j = 1 To H_Standi(i, 1)
            Line Input #nf, buf
            H_Stand1(i, j) = CSng(Mid(buf, 1, 5))
            H_Stand1t(i, j) = Mid(buf, 11, 4)
        Next j
        For j = 1 To H_Standi(i, 2)
            Line Input #nf, buf
            H_Stand2(i, j) = CSng(Mid(buf, 1, 5))
            H_Stand2t(i, j) = Mid(buf, 11, 13)
        Next j
    Next i
    Close #nf
    LOG_Out "OUT Constant_Read Normal Exit"
    On Error GoTo 0
    Exit Sub
ERH:
    If InStr(msg, "：") > 0 Then
        MsgBox msg & "地点の基準水位読み込み中にエラーが発生しました、基準水位を無効とします。" & vbCrLf & _
                     "DATAフォルダの基準水位.DATを修正して下さい。"
        Resume ERH1
    Else
        MsgBox "基準水位読み込み中にエラーが発生しました、DATAフォルダの基準水位.DATが無い可能性があります。" & vbCrLf & _
               "基準水位を無効とします。"
        Resume ERH1
    End If
ERH1:
    For i = 1 To 5
        For j = 1 To 5
            H_Stand1(i, j) = 99#
        Next j
        For j = 1 To 3
            H_Stand2(i, j) = 99#
        Next j
    Next i
    Close #nf
    LOG_Out "OUT Constant_Read ABNormal Exit"
    On Error GoTo 0
    Exit Sub
End Sub

'******************************************************************************
'サブルーチン：Date_dim(d As Date, x() As Integer)
'処理概要：
'******************************************************************************
Sub Date_dim(d As Date, x() As Integer)
    x(1) = Year(d)
    x(2) = Month(d)
    x(3) = Day(d)
    x(4) = Hour(d)
    x(5) = Minute(d)
    x(6) = Second(d)
End Sub

'******************************************************************************
'サブルーチン：C_Date(jx() As Integer) As Date
'処理概要：
'******************************************************************************
Function C_Date(jx() As Integer) As Date
    Dim d As String
    d = Format(jx(1), "####") & "/" & Format(jx(2), "00") & "/" & _
        Format(jx(3), "00") & " " & Format(jx(4), "00") & ":" & _
        Format(jx(5), "00")
    If IsDate(d) Then
        C_Date = CDate(d)
    Else
        MsgBox d & " は日付ではない"
        '******************************************************
        'Ver0.0.0 修正開始 1900/01/01 00:00
        '******************************************************
        'End
        '******************************************************
        'Ver0.0.0 修正終了 1900/01/01 00:00
        '******************************************************
    End If
End Function

'******************************************************************************
'サブルーチン：FeedBack(m As Integer)
'処理概要：
'******************************************************************************
Sub FeedBack(m As Integer)
    Dim i      As Integer
    Dim j      As Integer
    Dim nr     As Integer
    Dim ns     As Integer
    Dim hxj(6) As Single
    Dim hxc(6) As Single
    Dim frd(6) As Single
    Dim H1     As Single
    Dim H2     As Single
    Dim hx     As Single
    Dim hav    As Single
    Dim fd     As Single
    nr = m
    ns = V_Sec_Num(nr)                                                  'V_Sec_Num(nr)は不定流上の断面位置を表す
    hxj(1) = HO(nr + 2, Now_Step - 2)                                   '+2は１に日光川外水位(下流端水位)と2に洗堰流入量が入っている
    hxj(2) = HO(nr + 2, Now_Step - 1)
    hxj(3) = HO(nr + 2, Now_Step - 0)
    hxj(4) = HO(nr + 2, Now_Step + 1)
    hxj(5) = HO(nr + 2, Now_Step + 2)
    hxj(6) = HO(nr + 2, Now_Step + 3)
    hxc(1) = HQ(1, ns, NT - 30)
    hxc(2) = HQ(1, ns, NT - 24)
    hxc(3) = HQ(1, ns, NT - 18)
    hxc(4) = HQ(1, ns, NT - 12)
    hxc(5) = HQ(1, ns, NT - 6)
    hxc(6) = HQ(1, ns, NT - 0)
    frd(1) = HQ(1, ns, NT - 30)
    frd(2) = HQ(1, ns, NT - 24)
    frd(3) = HQ(1, ns, NT - 18)
    frd(4) = HQ(1, ns, NT - 12)
    frd(5) = HQ(1, ns, NT - 6)
    frd(6) = HQ(1, ns, NT - 0)
    CO(nr, 1) = hxc(3)
    CO(nr, 2) = hxc(4)
    CO(nr, 3) = hxc(5)
    CO(nr, 4) = hxc(6)
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'Print #Log_Num, V_Sec_Name(nr, 2)
    'Print #Log_Num, nr; " hxj(1)="; fmt(hxj(1)); " hxc(1)="; fmt(hxc(1)); "  FROUD="; fmt(HQ(3, ns, NT - 30))
    'Print #Log_Num, nr; " hxj(2)="; fmt(hxj(2)); " hxc(2)="; fmt(hxc(2)); "  FROUD="; fmt(HQ(3, ns, NT - 24))
    'Print #Log_Num, nr; " hxj(3)="; fmt(hxj(3)); " hxc(3)="; fmt(hxc(3)); "  FROUD="; fmt(HQ(3, ns, NT - 18))
    'Print #Log_Num, nr; " hxj(4)="; fmt(hxj(4)); " hxc(4)="; fmt(hxc(4)); "  FROUD="; fmt(HQ(3, ns, NT - 12))
    'Print #Log_Num, nr; " hxj(5)="; fmt(hxj(5)); " hxc(5)="; fmt(hxc(5)); "  FROUD="; fmt(HQ(3, ns, NT - 6))
    'Print #Log_Num, nr; " hxj(6)="; fmt(hxj(6)); " hxc(6)="; fmt(hxc(6)); "  FROUD="; fmt(HQ(3, ns, NT - 0))
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    For i = 1 To 3
        If hxj(i) < -20# Then hxj(i) = hxc(i)
    Next i
    hx = hxj(1) - hxc(1)
    Slide1(nr) = hx
    For i = 1 To 6
        hxc(i) = hxc(i) + hx
    Next i
    H1 = (hxj(2) - hxc(2)) / 6#
    H2 = (hxj(3) - hxc(3)) / 12#
    If H1 > H2 Then
        '******************************************************
        'Ver0.0.0 修正開始 1900/01/01 00:00
        '******************************************************
        hav = (H1 + H2) * 0.5
        'hav = h1
        '******************************************************
        'Ver0.0.0 修正終了 1900/01/01 00:00
        '******************************************************
    Else
        hav = H2
    End If
    '******************************************************
    'Ver0.0.0 修正開始 2003/07/14 00:00
    '******************************************************
    '2003/07/14 追加 水場川外水位が3.0m(警戒水位)より小さかったらスライド合わせとする。
    'If HO(5, Now_Step) < 3# Then
    '   hav = 0#
    'End If
    '******************************************************
    'Ver0.0.0 修正終了 2003/07/14 00:00
    '******************************************************
    '******************************************************
    'Ver0.0.0 修正開始 2003/08/12 00:00
    '******************************************************
    If hav < 0# And Froude < 0.1 Then hav = 0#
    'If hav > 0.03 Then hav = 0.03 '2003/0719 修正 これでもだいぶ違う(2003/08/12)
    '******************************************************
    'Ver0.0.0 修正終了 2003/08/12
    '******************************************************
    If hav > 0.06 Then hav = 0.06
    hav = 0#
    Print #Log_Num, nr; " hxj(1)="; fmt(hxj(1)); " hxc(1)="; fmt(hxc(1)); " hxj(1)-hxc(1)="; fmt(hxj(1) - hxc(1))
    Print #Log_Num, nr; " hxj(2)="; fmt(hxj(2)); " hxc(2)="; fmt(hxc(2)); " hxj(2)-hxc(2)="; fmt(hxj(2) - hxc(2))
    Print #Log_Num, nr; " hxj(3)="; fmt(hxj(3)); " hxc(3)="; fmt(hxc(3)); " hxj(3)-hxc(3)="; fmt(hxj(3) - hxc(3))
    Print #Log_Num, nr; " hxj(4)="; fmt(hxj(4)); " hxc(4)="; fmt(hxc(4)); fmt(hxc(4) + hav * 6)
    Print #Log_Num, nr; " hxj(5)="; fmt(hxj(5)); " hxc(5)="; fmt(hxc(5)); fmt(hxc(5) + hav * 12)
    Print #Log_Num, nr; " hxj(6)="; fmt(hxj(6)); " hxc(6)="; fmt(hxc(6)); fmt(hxc(6) + hav * 18)
    Print #Log_Num, nr; "     h1="; fmt(H1); "     h2="; fmt(H2); "           hav="; fmt(hav)
    j = 0
    For i = 18 To 0 Step -1
        HQ(1, ns, NT - i) = HQ(1, ns, NT - i) + hav * j
        j = j + 1
    Next i
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'hav = hav + 1#
    'For i = 0 To 18
    '    HQ(1, ns, NT - i) = HQ(1, ns, NT - i) + hav * (18 - i)
    'Next i
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    If hxj(3) < -10# Then                                               '欠測時は計算値とする 2004/04/26
        hx = 0#
    Else
        hx = hxj(3) - HQ(1, ns, NT - 18)
    End If
    Slide2(nr) = hx
    j = 0
    For i = NT - 18 To NT
        HQ(1, ns, i) = HQ(1, ns, i) + hx
        YHK(nr, j) = HQ(1, ns, i)                                       '県ＤＢ出力用
        Debug.Print "  nr="; nr; "  j="; j; " YHK="; YHK(nr, j)
        j = j + 1
    Next i
    CF(nr, 0) = HQ(1, ns, NT - 18)
    CF(nr, 1) = HQ(1, ns, NT - 12)
    CF(nr, 2) = HQ(1, ns, NT - 6)
    CF(nr, 3) = HQ(1, ns, NT - 0)
    Delta_H(nr) = hav * 6
End Sub

'******************************************************************************
'サブルーチン：FeedBack_Slide_Only(m As Integer)
'処理概要：
'******************************************************************************
Sub FeedBack_Slide_Only(m As Integer)
    Dim i      As Integer
    Dim j      As Integer
    Dim nr     As Integer
    Dim ns     As Integer
    Dim hxj    As Single
    Dim hx     As Single
    nr = m
    ns = V_Sec_Num(nr)
    hxj = HO(nr + 2, Now_Step)
    hx = hxj - HQ(1, ns, NT - 18)
    If hxj < -90# Then Exit Sub
    For i = NT - 18 To NT
        HQ(1, ns, i) = HQ(1, ns, i) + hx
    Next i
End Sub

'******************************************************************************
'サブルーチン：Flood_Data_Write_For_Calc()
'処理概要：
'******************************************************************************
Sub Flood_Data_Write_For_Calc()
    Dim i       As Integer
    Dim j       As Integer
    Dim nf      As Integer
    Dim htw     As Single
    Dim buf     As String
    Dim d       As Date
    Dim ht(500) As Single
    LOG_Out " IN Flood_Data_Write_For_Calc"
    nf = FreeFile
    Open Wpath & "時刻.DAT" For Output As #nf
    '******************************************************
    'データ開始時刻
    '******************************************************
    buf = ""
    For i = 1 To 6
        buf = buf & Format(str(js(i)), "@@@@@")
    Next i
    buf = buf & "      データ開始時刻"
    Print #nf, buf
    '******************************************************
    'データ終了時刻
    '******************************************************
    jxd = DateAdd("h", 3, jgd)
    Date_dim jxd, jx()
    buf = ""
    For i = 1 To 6
        buf = buf & Format(str(jx(i)), "@@@@@")
    Next i
    buf = buf & "      データ終了時刻"
    Print #nf, buf
    '******************************************************
    'データピッチ
    '******************************************************
    buf = ""
    For i = 1 To 5
        Data_Pich(i) = 3600
        buf = buf & Format(str(Data_Pich(i)), "@@@@@")
    Next i
    buf = buf & "           データの時間ピッチ(秒)"
    Print #nf, buf
    Print #nf, Space(4) & Format(JRADAR, "0") & Format(Format(Rsa_Mag, "#0.00"), "@@@@@") '0=テレメータ雨量 1=レーダー雨量
    Close #nf
    '******************************************************
    '実績雨量+予測降雨
    '******************************************************
    d = jsd
    nf = FreeFile
    Open Wpath & "雨.DAT" For Output As #nf
    Print #nf, RO_Title
    For i = 1 To Now_Step + Yosoku_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        For j = 1 To Rnum
            buf = buf & Format(Format(RO(j, i), "#######0.0"), "@@@@@@@@@@")
        Next j
        Print #nf, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #nf
    '******************************************************
    '下流端水位(潮位)
    '******************************************************
    d = jsd
    nf = FreeFile
    Open Wpath & "下流端水位.DAT" For Output As #nf
    Print #nf, "  DATE     TIME      Cal_H      Tide   Suiba_H"
    '******************************************************
    '計算を通すための苦肉の策
    '******************************************************
    For i = 1 To All_Step
        ht(i) = HO(1, i)
    Next i
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'If HO(3, Now_Step) <= 2.5 Then                                     '現時刻水場外水位が2.5ｍ以下のとき
    'htw = -99#
    'For j = Now_Step - 3 To 4
    '    If HO(1, j) > htw Then htw = HO(1, j)
    'Next j
    'For j = Now_Step - 3 To 4
    '    ht(j) = htw
    'Next j
    'End If
    '策 終わり
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    For i = 1 To All_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        buf = buf & Format(Format(ht(i), "######0.00"), "@@@@@@@@@@")
        buf = buf & Format(Format(HO(1, i), "######0.00"), "@@@@@@@@@@")
        buf = buf & Format(Format(HO(3, i), "######0.00"), "@@@@@@@@@@")
        Print #nf, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #nf
    '******************************************************
    '洗堰流入量
    '******************************************************
    d = jsd
    nf = FreeFile
    Open Wpath & "洗堰.DAT" For Output As #nf
    Print #nf, "  DATE     TIME       洗堰"
    For i = 1 To All_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        buf = buf & Format(Format(HO(2, i), "######0.00"), "@@@@@@@@@@")
        Print #nf, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #nf
    LOG_Out " OUT Flood_Data_Write_For_Calc"
End Sub

'******************************************************************************
'サブルーチン：Flood_Data_Write_For_Calc1()
'処理概要：
'******************************************************************************
Sub Flood_Data_Write_For_Calc1()
    Dim i       As Long
    Dim j       As Long
    Dim m       As Long
    Dim fi      As Long
    Dim fo      As Long
    Dim buf     As String
    Dim Titl    As String
    Dim d       As Date
    Dim dd(6)   As Integer
    Dim x(500)  As String
    Dim sx(500) As String
    Dim nf
    Dim ht(500) As Single
    Dim htw     As Single
    Dim Steps   As Long
    Dim ht_max  As Single
    Dim iht_max As Long
    Dim BeforeTime As Long
    Dim Start_h As Single
    Dim d_nsk   As Date                                                 'ＮＳＫ用の開始時刻
    Dim s_nsk   As Integer                                              'ＮＳＫ用のデータ開始ステップ
    LOG_Out " IN Flood_Data_Write_For_Calc1"
    Start_h = MAIN.Text2                                                 '不定流計算をスムーズに行う為出発水位を高いところから計算したいた為の設定値
    For i = 1 To All_Step
        ht(i) = HO(1, i)
    Next i
    '******************************************************
    '計算開始ステップを探す、水位が指定された水位(Main.Text2)以上を探す。
    '******************************************************
    ht_max = -999
    For j = 1 To All_Step - 5
        If ht(j) >= Start_h Then
            m = j
            GoTo jump1
        End If
        If ht_max > ht(j) Then
            ht_max = ht(j)
            iht_max = j
        End If
    Next j
    LOG_Out "こんなことはあってはならない"
    m = iht_max
jump1:
    Steps = m
    d_nsk = DateAdd("h", Steps - 1, jsd)                                '不定流計算開始時刻
    '******************************************************
    '下流端水位(潮位)
    '******************************************************
    LOG_Out "NSK_下流端水位.DAT  For Output ---OPEN"
    fo = FreeFile
    Open Wpath & "NSK_下流端水位.DAT" For Output As #fo
    Print #fo, "  DATE     TIME      Cal_H      Tide   Suiba_H"
    d = d_nsk
    For i = Steps To All_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        buf = buf & Format(Format(ht(i), "######0.00"), "@@@@@@@@@@")
        buf = buf & Format(Format(HO(1, i), "######0.00"), "@@@@@@@@@@")
        buf = buf & Format(Format(HO(3, i), "######0.00"), "@@@@@@@@@@")
        Print #fo, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #fo
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'H_Sea = HO(1, Now_Step - 3)                                        '不等流計算用下流端水位+++++++++++++++++++++++++++初期水面形境界条件
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    H_Sea = HO(1, Steps)                                                '不等流計算用下流端水位+++++++++++++++++++++++++++初期水面形境界条件
    LOG_Out "NSK_下流端水位.DAT  For Output ---CLOSE"
    '******************************************************
    '計算時刻設定
    '******************************************************
    Date_dim d_nsk, dd()
    LOG_Out "NSK_時刻.DAT For Output ---OPEN"
    fo = FreeFile
    Open Wpath & "NSK_時刻.DAT" For Output As #fo
    '******************************************************
    'データ開始時刻
    '******************************************************
    buf = ""
    For i = 1 To 6
        buf = buf & Format(str(dd(i)), "@@@@@")
    Next i
    buf = buf & "      データ開始時刻"
    NSK_jsd = d_nsk
    Print #fo, buf
    '******************************************************
    'データ終了時刻
    '******************************************************
    jxd = DateAdd("h", 3, jgd)
    Date_dim jxd, jx()
    buf = ""
    For i = 1 To 6
        buf = buf & Format(str(jx(i)), "@@@@@")
    Next i
    buf = buf & "      データ終了時刻"
    Print #fo, buf
    '******************************************************
    'データピッチ
    '******************************************************
    buf = ""
    For i = 1 To 5
        buf = buf & " 3600"
    Next i
    buf = buf & "           データの時間ピッチ(秒)"
    Print #fo, buf
    Close #fo
    LOG_Out "NSK_時刻.DAT For Output ---CLOSE"
    '******************************************************
    'SHINK10.U07
    '******************************************************
    LOG_Out "SHINK10.U07  For Input ---OPEN"
    fi = FreeFile
    Open Wpath & "SHINK10.U07" For Input As #fi
    Line Input #fi, Titl
    i = 0
    Do Until EOF(fi)
        i = i + 1
        Line Input #fi, x(i)
    Loop
    Close #fi
    LOG_Out "SHINK10.U07  For Input ---CLOSE"

    LOG_Out "NSK_SHINK10.U07  For Output ---OPEN"
    m = DateDiff("h", d_nsk, jxd) + 1
    Mid(Titl, 1, 10) = " 3600" & Format(str(m), "@@@@@")
    fo = FreeFile
    Open Wpath & "NSK_SHINK10.U07" For Output As #fo
    Print #fo, Titl
    For j = Steps To i
        Print #fo, x(j)
    Next j
    Close #fo
    LOG_Out "NSK_SHINK10.U07  For Output ---CLOSE"
    '******************************************************
    '洗堰流入量
    '******************************************************
    LOG_Out "NSK_洗堰.DAT  For Output ---OPEN"
    d = d_nsk
    fo = FreeFile
    Open Wpath & "NSK_洗堰.DAT" For Output As #fo
    Print #fo, "  DATE     TIME       洗堰"
    For i = Steps To All_Step
        buf = Format(d, "yyyy/mm/dd hh:nn")
        buf = buf & Format(Format(HO(2, i), "######0.00"), "@@@@@@@@@@")
        Print #fo, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #fo
    LOG_Out "NSK_洗堰.DAT  For Output ---CLOSE"
    '******************************************************
    'KUJINO.DAT
    '******************************************************
    LOG_Out "KUJINO.DAT For Input ---OPEN"
    fi = FreeFile
    Open Wpath & "KUJINO.DAT" For Input As #fi
    Line Input #fi, Titl
    i = 0
    Do Until EOF(fi)
        i = i + 1
        Line Input #fi, x(i)
    Loop
    Close #fi
    LOG_Out "KUJINO.DAT For Input ---CLOSE"
    LOG_Out "NSK_KUJINO.DAT  For Output ---OPEN"
    fo = FreeFile
    Open Wpath & "NSK_KUJINO.DAT" For Output As #fo
    Print #fo, Titl
    For j = Steps To i
        Print #fo, x(j)
    Next j
    Close #fo
    Q_kuji = CSng(Mid(x(Steps), 7, 10))                                 '不等流計算用 久地野流入量++++++++++初期水面形境界条件
    LOG_Out "NSK_KUJINO.DAT  For Output ---CLOSE"
    '******************************************************
    'HARUHI.DAT
    '******************************************************
    LOG_Out "HARUHI.DAT  For Input ---OPEN"
    fi = FreeFile
    Open Wpath & "HARUHI.DAT" For Input As #fi
    Line Input #fi, Titl
    i = 0
    Do Until EOF(fi)
        i = i + 1
        Line Input #fi, x(i)
    Loop
    Close #fi
    LOG_Out "HARUHI.DAT  For Input ---CLOSE"

    LOG_Out "NSK_HARUHI.DAT  For Output ---OPEN"
    fo = FreeFile
    Open Wpath & "NSK_HARUHI.DAT" For Output As #fo
    Print #fo, Titl
    For j = Steps To i
        Print #fo, x(j)
    Next j
    Close #fo
    LOG_Out "NSK_HARUHI.DAT  For Output ---CLOSE"
    Q_Haru = CSng(Mid(x(Steps), 7, 10))                                 '不等流計算用 春日流入量+++++++++++初期水面形境界条件
    LOG_Out " OUT Flood_Data_Write_For_Calc1"
End Sub

'******************************************************************************
'サブルーチン：fmt(c As Variant) As String
'処理概要：
'******************************************************************************
Function fmt(c As Variant) As String
   fmt = Format(Format(c, "###0.0000"), "@@@@@@@@@")
End Function

'******************************************************************************
'サブルーチン：Froude_Check(irc As Boolean)
'処理概要：
'******************************************************************************
Sub Froude_Check(irc As Boolean)
    Dim i      As Integer
    Dim j      As Integer
    Dim nr     As Integer
    Dim ns     As Integer
    Dim fds    As Single
    Dim fd     As Single
    ReDim frd(6, V_Sec_Cnt) As Single
    irc = True
    fds = 0#
    For nr = 1 To V_Sec_Cnt
        ns = V_Sec_Num(nr)                                              'V_Sec_Num(nr)は不定流上の断面位置を表す
        frd(1, nr) = HQ(3, ns, NT - 30)
        frd(2, nr) = HQ(3, ns, NT - 24)
        frd(3, nr) = HQ(3, ns, NT - 18)
        frd(4, nr) = HQ(3, ns, NT - 12)
        frd(5, nr) = HQ(3, ns, NT - 6)
        frd(6, nr) = HQ(3, ns, NT - 0)
        CO(nr, 1) = HQ(1, ns, NT - 12)
        CO(nr, 2) = HQ(1, ns, NT - 6)
        CO(nr, 3) = HQ(1, ns, NT - 0)
        Print #Log_Num, V_Sec_Name(nr, 2)
        Print #Log_Num, nr; "  FROUD="; fmt(frd(1, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(2, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(3, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(4, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(5, nr))
        Print #Log_Num, nr; "  FROUD="; fmt(frd(6, nr))
        fd = 0#
        For i = 1 To 6
            fd = fd + frd(i, nr)
        Next i
        fds = fds + fd / 6
        Print #Log_Num, nr; " fd="; fmt(fd / 6)
    Next nr
    Froude = fds / V_Sec_Cnt
    Print #Log_Num, "  Average Froude="; fmt(Froude)
End Sub

'******************************************************************************
'サブルーチン：Hydro_Graph(OBS_Point As Integer)
'処理概要：
'******************************************************************************
Sub Hydro_Graph(OBS_Point As Integer)
    '******************************************************
    '地点プロット
    '******************************************************
    Dim dl     As Single
    Dim dp     As Single
    Dim xs     As Single
    Dim xe     As Single
    Dim ys     As Single
    Dim ye     As Single
    Dim xl     As Single
    Dim yl     As Single
    Dim Hp     As Single
    Dim xw     As Single
    Dim yw     As Single
    Dim fj     As Single
    Dim amu    As Single
    Dim amd    As Single
    Dim amp    As Single
    Dim sc     As Single
    Dim ysm    As Single
    Dim psm    As Single
    Dim msize  As Single
    Dim hqyl   As Single
    Dim ps     As Single
    Dim x      As Single
    Dim y      As Single
    Dim i      As Integer
    Dim j      As Integer
    Dim nday   As Integer
    Dim nbun   As Integer
    Dim niti   As String
    Dim J1     As Integer
    Dim j2     As Integer
    Dim j3     As Integer
    Dim n      As Integer
    Dim mn     As Integer
    Dim xt     As Single
    Dim nr     As Integer
    Dim dw     As Date
    Dim KanG   As Integer
    Dim moji   As String
    Dim w      As Single
    Dim ns     As Integer
    Dim Mp     As Single
    Dim na     As Integer
    Dim rs     As Single
    Dim t1     As String
    Dim t2     As String
    Dim t3     As String
    Dim T4     As String
    Dim Kijun_Name As String
    xl = 215: yl = 155: hqyl = 115
    xs = 35: ys = 30: xe = xs + xl: ye = ys + yl
    KanG = OBS_Point
    Select Case OBS_Point
        Case 1                                                          '下之一色
            Kijun_Name = "下之一色"
            na = 1
        Case 2                                                          '大治
            Kijun_Name = "大　治　"
            na = 2
        Case 3                                                          '水場外水位
            Kijun_Name = "水場外水位"
            na = 3
        Case 4                                                          '久地野
            Kijun_Name = "久地野"
            na = 4
        Case 5                                                          '春日
            Kijun_Name = "春　日　"
            na = 5
    End Select
    nr = KanG
    nday = DateDiff("d", jsd, jgd) + 1
    If nday < 3 Then nday = 3
    VS_Box xs, ys, xe, ye, 0, 0.4, 15, 1
    If isRAIN = "02" Then
        VS_symbol xs, ys - 1#, 12#, "ＦＲＩＣＳ雨量使用", 3
    Else
        VS_symbol xs, ys - 1#, 12#, "気象庁雨量使用", 3
    End If
    dw = jsd
    dl = xl / nday
    '******************************************************
    '時刻目盛り
    '******************************************************
    Hp = dl / 24
    Mp = Hp / 60
    For i = 1 To nday
        x = xs + dl * (i - 1)
        If i <> nday Then
            J1 = 0
        Else
            J1 = 1
        End If
        For j = 0 To 23 + J1
            xw = x + Hp * j
            If (j Mod 6) = 0 Then
                VS_Line xw, ye + 1.5, xw, ys, 0, 0
                If j = 0 Then
                    VS_Line xw, ys, xw, ye, 0, 0
                Else
                    VS_Line xw, ys, xw, ye, 8, 0
                End If
                fj = j Mod 24
                VS_symbol xw, ye + 2#, 8.5, Cvt_2byte(Trim(str(j Mod 24))), 4
            Else
                VS_Line xw, ye + 1#, xw, ye, 0, 0
            End If
            If j = 12 Then
                If i > 1 Then
                    dw = DateAdd("d", 1, dw)
                    j2 = Month(dw)
                    j3 = Day(dw)
                Else
                    j2 = Month(dw)
                    j3 = Day(dw)
                End If
                niti = Cvt_2byte(Format(j2, "##") + "/" + Format(j3, "##"))
                VS_symbol xw, ye + 8#, 12.5, niti, 4
            End If
        Next j
    Next i
    niti = "現時刻 " + Cvt_2byte(Format(jg(2), "##")) & "月" & _
                       Cvt_2byte(Format(jg(3), "##")) & "日" & _
                       Cvt_2byte(Format(jg(4), "#0")) & "時" & _
                       Cvt_2byte(Format(jg(5), "#0")) & "分"
    VS_symbol xe, ys - 1#, 12#, niti, 9
    VS_symbol xs + xl * 0.5, ys - 2.5, 14, Kijun_Name & "観測所", 6
    xt = 20#
    '******************************************************
    '雨量目盛り
    '******************************************************
    rs = 0.5
    For y = 20 To 60 Step 20
        yw = ys + y * 0.5
        VS_Line xs - 1.5, yw, xe, yw, 8, 0
        VS_symbol xs - 2#, yw, 9#, Cvt_2byte(str(y)), 8
    Next y
    VS_symbol xs - xt, ys + 12#, 11.5, "雨", 5
    VS_symbol xs - xt, ys + 20#, 11.5, "量", 5
    moji = Cvt_2byte("(mm)")
    VS_symbol xs - xt, ys + 25#, 9#, moji, 5
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'If Graph.Option2(0) Then
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
        mn = 1
        VS_symbol xs - xt, ye - 80#, 11.5, "水", 5
        VS_symbol xs - xt, ye - 50#, 11.5, "位", 5
        moji = Cvt_2byte("( m )")
        VS_symbol xs - xt, ye - 40#, 10#, moji, 5
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'End If
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'If Graph.Option2(1) Then
    '    mn = 2
    '    Call tv_symbol(xs - xt, ys + 80#, 11.5, "流", 5)
    '    Call tv_symbol(xs - xt, ys + 50#, 11.5, "量", 5)
    '    Call tv_symbol(xs - xt, ys + 40#, 10#, "(m3/s)", 5)
    'End If
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    '******************************************************
    '目盛り
    '******************************************************
     If mn = 1 Then
        amd = H_Scale(KanG, 1)
        amu = H_Scale(KanG, 2)
        amp = H_Scale(KanG, 3)
    End If
    sc = hqyl / (amu - amd)
    For fj = amd To amu + 0.01 Step amp
        y = ye - (fj - amd) * sc
        VS_Line xs - 1.5, y, xe, y, 8, 0
        If mn = 1 Then
            VS_number xs - 2#, y, 9#, fj, 1, 8                          '水位
        Else
            VS_number xs - 2#, y, 9#, fj, -1, 8                         '流量
        End If
    Next fj
    If mn = 1 Then
        '******************************************************
        '基準水位
        '******************************************************
        If H_Stand1(nr, 4) < amu And H_Stand1(nr, 4) > amd Then
            y = ye - (H_Stand1(nr, 4) - amd) * sc
            VS_Line xs, y, xe, y, 2, 0.5
            VS_symbol xe + 1#, y, 9#, H_Stand1t(nr, 4), 3
            VS_number xe + 1#, y, 8#, H_Stand1(nr, 4), 2, 1
        End If
        If H_Stand1(nr, 3) < amu And H_Stand1(nr, 3) > amd Then
            y = ye - (H_Stand1(nr, 3) - amd) * sc
            VS_Line xs, y, xe, y, 5, 0.5
            VS_symbol xe + 1#, y, 9#, H_Stand1t(nr, 3), 3
            VS_number xe + 1#, y, 8#, H_Stand1(nr, 3), 2, 1
        End If
        If H_Stand1(nr, 2) < amu And H_Stand1(nr, 2) > amd Then
            y = ye - (H_Stand1(nr, 2) - amd) * sc
            VS_Line xs, y, xe, y, 5, 0.5
            VS_symbol xe + 1#, y, 9#, H_Stand1t(nr, 2), 3
            VS_number xe + 1#, y, 8#, H_Stand1(nr, 2), 2, 1
        End If
        If H_Stand1(nr, 1) < amu And H_Stand1(nr, 1) > amd Then
            y = ye - (H_Stand1(nr, 1) - amd) * sc
            VS_Line xs, y, xe, y, 12, 0.5
            VS_symbol xe + 1#, y, 9#, H_Stand1t(nr, 1), 3
            VS_number xe + 1#, y, 8#, H_Stand1(nr, 1), 2, 1
        End If
        '******************************************************
        'ポンプ水位
        '******************************************************
        If H_Stand2(nr, 1) < amu And H_Stand2(nr, 1) > amd Then
            y = ye - (H_Stand2(nr, 1) - amd) * sc
            VS_Line xs, y, xe, y, 12, 0.5
            VS_symbol xs + 1#, y, 9#, H_Stand2t(nr, 1), 3
            VS_number xs + 1#, y, 8#, H_Stand2(nr, 1), 2, 1
        End If
        If H_Stand2(nr, 2) < amu And H_Stand2(nr, 2) > amd Then
            y = ye - (H_Stand2(nr, 2) - amd) * sc
            VS_Line xs, y, xe, y, 12, 0.5
            VS_symbol xs + 17#, y, 9#, H_Stand2t(nr, 2), 3
            VS_number xs + 17#, y, 8#, H_Stand2(nr, 2), 2, 1
        End If
        If H_Stand2(nr, 3) < amu And H_Stand2(nr, 3) > amd Then
            y = ye - (H_Stand2(nr, 3) - amd) * sc
            VS_Line xs, y, xe, y, 12, 0.5
            VS_symbol xs + 1#, y, 9#, H_Stand2t(nr, 3), 3
            VS_number xs + 1#, y, 8#, H_Stand2(nr, 3), 2, 1
        End If
    End If
    '******************************************************
    '雨量プロット
    '******************************************************
    xw = xs + js(4) * Hp
    For i = 1 To Now_Step                                               '計算開始から現時刻まで
        If RO(na, i) > 1# Then
            w = RO(na, i) * rs                                          '流域数＋１から各基準地点上流域平均雨量
            '******************************************************
            'Ver0.0.0 修正開始 1900/01/01 00:00
            '******************************************************
            'Debug.Print "  RO="; RO(na, i)
            '******************************************************
            'Ver0.0.0 修正終了 1900/01/01 00:00
            '******************************************************
            ps = xw + (i - 1) * Hp
            VS_Box ps - Hp, ys, ps, ys + w, QBColor(9), 0, QBColor(9), 0
        End If
    Next i
    For i = Now_Step + 1 To All_Step                                    '現時刻＋１から予測時間ステップまで
        If RO(na, i) > 1# Then
            w = RO(na, i) * rs                                          '流域数＋１から各基準地点上流域平均雨量
            ps = xw + (i - 1) * Hp
            '******************************************************
            'Ver0.0.0 修正開始 1900/01/01 00:00
            '******************************************************
            'Debug.Print "  RO="; RO(na, i)
            '******************************************************
            'Ver0.0.0 修正終了 1900/01/01 00:00
            '******************************************************
            VS_Box ps - Hp, ys, ps, ys + w, QBColor(14), 0, QBColor(14), 0
        End If
    Next i
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'iam.DrawWidth = 2
    'If mn = 1 Then
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
        '******************************************************
        '実績水位プロット
        '******************************************************
        xw = xs + js(4) * Hp + js(5) * Mp
        If OBS1 Then
            n = All_Step
        Else
            n = Now_Step
        End If
        For i = 1 To n
            If HO(nr + 2, i) <> aksk Then
                w = ye - (HO(nr + 2, i) - amd) * sc
                ps = xw + (i - 1) * Hp
                VS_Circle ps, w, 0.7, 0, 0, 0, 0
            End If
        Next i
        '******************************************************
        '現時刻まで計算水位プロット
        '******************************************************
        If CAL1 = 1 Then
            ns = V_Sec_Num(nr)
            xw = xs + (js(4) + Now_Step - 4) * Hp + js(5) * Mp          '不定流計算現時刻データ開始から用
            ysm = ye - (HQ(1, ns, 1) - amd) * sc
            psm = xw
            For i = 1 To NT - 17
                amu = ye - (HQ(1, ns, i) - amd) * sc
                ps = xw + i * Mp
                VS_Line psm, ysm, ps, amu, 4, 0.5
                psm = ps
                ysm = amu
            Next i
        End If
        '******************************************************
        '予測計算水位プロット
        '******************************************************
        If Category Then                                                '不定流時
            ns = V_Sec_Num(nr)
            xw = xs + (DateDiff("h", jsd, jgd) + js(4)) * Hp + js(5) * Mp '不定流計算現時刻－４時間用
            If HO(nr + 2, Now_Step) > -90# Then
                ysm = ye - (HO(nr + 2, Now_Step) - amd) * sc
            Else
                ysm = ye - (HQ(1, ns, NT - 18) - amd) * sc
            End If
            psm = xw
            j = 0
            For i = NT - 17 To NT
                j = j + 1
                amu = ye - (HQ(1, ns, i) - amd) * sc
                ps = xw + j * Mp * 10
                VS_Line psm, ysm, ps, amu, 2, 0.5
                psm = ps
                ysm = amu
            Next i
        '******************************************************
        '不等流時
        '******************************************************
        Else
            ysm = ye - (Nonuni_H(nr, 0) - amd) * sc
            xw = xs + (DateDiff("h", jsd, jgd) + js(4)) * Hp + js(5) * Mp
            psm = xw
            For i = 1 To 3
                amu = ye - (Nonuni_H(nr, i) - amd) * sc
                ps = xw + i * Hp
                VS_Line psm, ysm, ps, amu, 2, 0.5
                psm = ps
                ysm = amu
            Next i
        End If
        '******************************************************
        '生
        '******************************************************
        If Beer Then
            xw = xs + js(4) * Hp + js(5) * Mp                           '不定流計算現時刻－４時間用
            ysm = ye - (CO(nr, 1) - amd) * sc
            psm = xw
            For i = 2 To 4
                amu = ye - (CO(nr, i) - amd) * sc
                ps = xw + Hp * (i - 1)
                VS_Line psm, ysm, ps, amu, 4, 0.4
                psm = ps
                ysm = amu
            Next i
        End If
        VS_ShowPage 3
        '******************************************************
        '予測履歴プロット
        '******************************************************
        If History Then
            Dim yp0 As Single, YP1 As Single, YP2 As Single, YP3 As Single
            Dim tp As Date
            Dim XB As Single
            MDB_履歴_Read
            XB = xs + js(4) * Hp
            For i = 1 To Now_Step
                If H_Pred(i, OBS_Point, 1) > -80# Then
                    tp = T_Pred(i)
                    xw = XB + DateDiff("h", jsd, tp) * Hp + Minute(tp) * Mp
                    yp0 = ye - (H_Pred(i, OBS_Point, 1) - amd) * sc
                    YP1 = ye - (H_Pred(i, OBS_Point, 2) - amd) * sc
                    YP2 = ye - (H_Pred(i, OBS_Point, 3) - amd) * sc
                    YP3 = ye - (H_Pred(i, OBS_Point, 4) - amd) * sc
                    VS_Line xw, yp0, xw + Hp, YP1, 1, 0.2
                    VS_Line xw + Hp, YP1, xw + Hp * 2, YP2, 1, 0.2
                    VS_Line xw + Hp * 2, YP2, xw + Hp * 3, YP3, 1, 0.2
                End If
            Next i
        End If
    'End If
End Sub

'******************************************************************************
'サブルーチン：Initial_Constant()
'処理概要：
'******************************************************************************
Sub Initial_Constant()
    '******************************************************
    '下之一色
    '******************************************************
    H_Scale(1, 1) = -1
    H_Scale(1, 2) = 5
    H_Scale(1, 3) = 0.5
    '******************************************************
    '大治
    '******************************************************
    H_Scale(2, 1) = 0
    H_Scale(2, 2) = 5
    H_Scale(2, 3) = 0.5
    '******************************************************
    '水場外
    '******************************************************
    H_Scale(3, 1) = 0
    H_Scale(3, 2) = 7
    H_Scale(3, 3) = 0.5
    '******************************************************
    '久地野
    '******************************************************
    H_Scale(4, 1) = 0
    H_Scale(4, 2) = 8
    H_Scale(4, 3) = 0.5
    '******************************************************
    '春日
    '******************************************************
    H_Scale(5, 1) = 1
    H_Scale(5, 2) = 6
    H_Scale(5, 3) = 0.5
End Sub

'******************************************************************************
'サブルーチン：Initial_Data()
'処理概要：
'******************************************************************************
Sub Initial_Data()
    '******************************************************
    '雨量観測所名
    '******************************************************
    Name_R(1) = "犬　山"
    Name_R(2) = "一ノ宮県"
    Name_R(3) = "一ノ宮気"
    Name_R(4) = "小　牧"
    Name_R(5) = "名古屋空"
    Name_R(6) = "春日井"
    Name_R(7) = "大　里"
    Name_R(8) = "名古屋県"
    Name_R(9) = "蟹　江"
    Name_R(10) = "松　本"
    '******************************************************
    '水位観測所名
    '******************************************************
    Name_H(1) = "日光川外"
    Name_H(2) = "洗　堰"
    Name_H(3) = "下之一色"
    Name_H(4) = "大　治"
    Name_H(5) = "水場川外"
    Name_H(6) = "久地野"
    Name_H(7) = "春　日"
End Sub

'******************************************************************************
'サブルーチン：PDF_Check()
'処理概要：
'******************************************************************************
Sub PDF_Check()
    Dim i        As Long
    Dim ns       As Long
    Dim F        As String
    Dim PDF_Out  As Boolean
    Dim T        As String
    PDF_Out = False
    '******************************************************
    '水場川外水位の実績をチェック
    '******************************************************
    For i = 1 To Now_Step
        If HO(5, i) >= 2# Then
            PDF_Out = True
            Exit For
        End If
    Next i
    If PDF_Out Then
        GoTo PDF_Put
    End If
    '******************************************************
    '水場川外水位の予測をチェック
    '******************************************************
    For i = NT - 17 To NT
        If HQ(1, 3, i) >= 2# Then                                       '水防団待機水位を超える水位があったらPDFを出力
            PDF_Out = True
            Exit For
        End If
    Next i
    If PDF_Out = False Then
        Exit Sub
    End If
PDF_Put:
    T = Format(jgd, "yyyymmddhhnn")
    If isRAIN = 1 Then
        T = "気象庁" & T
    Else
        T = "FRICS" & T
    End If
    Graph3.VSPDF1.ConvertDocument Graph3.VSP, App.Path & "\DATA\PDF\" & T & ".pdf"
End Sub

'******************************************************************************
'サブルーチン：Pump_Full()
'処理概要：
'******************************************************************************
Sub Pump_Full()
    Dim i    As Integer
    Dim n1   As Integer
    Dim n2   As Integer
    Dim buf  As String
    n1 = FreeFile
    Open App.Path & "\DATA\PFULL.DAT" For Input As #n1
    n2 = FreeFile
    Open Wpath & "Pump.dat" For Output As #n2
    Do Until EOF(n1)
        Line Input #n1, buf
        Print #n2, buf
    Loop
    Close #n1
    Close #n2
End Sub

'******************************************************************************
'サブルーチン：Section_Read()
'処理概要：
'******************************************************************************
Sub Section_Read()
    Dim i    As Integer
    Dim j    As Integer
    Dim i1   As Integer
    Dim i2   As Integer
    Dim nf   As Integer
    Dim NDx  As Integer
    Dim DXS  As Single
    Dim buf  As String
    nf = FreeFile
    Open App.Path & "\data\Section.dat" For Input As #nf
    Input #nf, buf
    NDx = CInt(Mid(buf, 1, 5))
    CAL1 = CInt(Mid(buf, 6, 5))                                             'NDx=断面数  CAL1=計算値プロットの有=1無=0
    ReDim DX(NDx), NDN(NDx)
    V_Sec_Cnt = 0
    For i = 1 To NDx
        Input #nf, buf
        NDN(i) = Mid(buf, 1, 6)
        DX(i) = CSng(Mid(buf, 10, 6))
        ZS(i) = CSng(Mid(buf, 21, 10))
        MDX = MDX + DX(i)
        If Mid(buf, 32, 5) <> "" Then
            V_Sec_Cnt = V_Sec_Cnt + 1
            V_Sec_Num(V_Sec_Cnt) = i
            V_Sec_Name(V_Sec_Cnt, 1) = NDN(i)
            V_Sec_Name(V_Sec_Cnt, 2) = Mid(buf, 32, 5)
        End If
    Next i
    MDX = 0
    For i = 1 To 53
        MDX = MDX + DX(i)
    Next i
    ReDim sdx(V_Sec_Cnt)
    i1 = 1
    DXS = 0#
    For j = 1 To V_Sec_Cnt - 1
        i2 = V_Sec_Num(j)
        For i = i1 To i2
            DXS = DXS + DX(i)
            sdx(j) = DXS
        Next i
        i1 = i2 + 1
    Next j
    Close #nf
End Sub

'******************************************************************************
'サブルーチン：Short_Break(S As Long)
'処理概要：
'******************************************************************************
Public Sub Short_Break(S As Long)
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'Sleep s * 1000   'システムが止めるので反応しなくなるからだめ
    '                  ただし、ＣＰＵを使わなくなる。
    '                  下の方法は反応するがＣＰＵを１００％使う事に
    '                  なる。
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    Dim i   As Date
    Dim j   As Date
    Dim k   As Long
    i = Now
    Do
        j = Now
        k = DateDiff("s", i, j)
        If k >= S Then
            Exit Do
        End If
        DoEvents
    Loop
    Exit Sub
End Sub

'******************************************************************************
'サブルーチン：Flood_Data_Read()
'処理概要：
'******************************************************************************
Sub Flood_Data_Read()
    Dim a
    Dim b
    Dim i      As Integer
    Dim j      As Integer
    Dim nf     As Integer
    Dim NG     As Integer
    Dim buf    As String
    LOG_Out " IN Flood_Data_Read"
    nf = FreeFile
    Open Input_file For Input As #nf
    '******************************************************
    '洪水データタイトル
    '******************************************************
    Line Input #nf, Flood_Name
    '******************************************************
    'データ開始時刻
    '******************************************************
    Line Input #nf, buf
    js(1) = CInt(Mid(buf, 1, 5))
    js(2) = CInt(Mid(buf, 6, 5))
    js(3) = CInt(Mid(buf, 11, 5))
    js(4) = CInt(Mid(buf, 16, 5))
    js(5) = CInt(Mid(buf, 21, 5))
    js(6) = 0
    jsd = C_Date(js())
    '******************************************************
    'データ終了時刻
    '******************************************************
    Line Input #nf, buf
    jg(1) = CInt(Mid(buf, 1, 5))
    jg(2) = CInt(Mid(buf, 6, 5))
    jg(3) = CInt(Mid(buf, 11, 5))
    jg(4) = CInt(Mid(buf, 16, 5))
    jg(5) = CInt(Mid(buf, 21, 5))
    jg(6) = 0
    jgd = C_Date(jg())
    '******************************************************
    'データピッチ
    '******************************************************
    Line Input #nf, buf
    Data_Pich(1) = CSng(Mid(buf, 1, 5))
    Data_Pich(2) = CSng(Mid(buf, 6, 5))
    Data_Pich(3) = CSng(Mid(buf, 11, 5))
    Data_Pich(4) = CSng(Mid(buf, 16, 5))
    Data_Pich(5) = CSng(Mid(buf, 21, 5))
    Data_Steps = DateDiff("h", jsd, jgd) + 1
    Now_Step = Data_Steps
    All_Step = Now_Step + Yosoku_Step
    '******************************************************
    'レーダー雨量の有無
    '******************************************************
    Line Input #nf, buf
    IRADAR = CInt(Mid(buf, 1, 5))
    If IRADAR = 1 Then
        Radar_File = Trim(Mid(buf, 6, 30))                              'レーダー雨量ファイル名
        MAIN.Check2.Enabled = True
    Else
        Radar_File = ""
        MAIN.Check2.Enabled = False
    End If
    '******************************************************
    '実績雨量
    '******************************************************
    Line Input #nf, RO_Title
    For i = 1 To Data_Steps
        Line Input #nf, buf
        For j = 1 To Rnum
            RO(j, i) = CSng(Mid(buf, (j - 1) * 10 + 17, 10))
        Next j
    Next i
    '******************************************************
    '実績水位、潮位、洗堰
    '******************************************************
    Line Input #nf, HO_Title
    For i = 1 To Data_Steps
        Line Input #nf, buf
        For j = 1 To Hnum
            HO(j, i) = CSng(Mid(buf, (j - 1) * 10 + 17, 10))
        Next j
    Next i
    '******************************************************
    'ポンプデータ
    '******************************************************
    If EOF(nf) Then
        Pump_Full
    Else
        Input #nf, buf
        a = InStr(buf, "ポンプ")
        If a > 0 Then
            If OBS_Pump Then
                '******************************************************
                'Ver0.0.0 修正開始 1900/01/01 00:00
                '******************************************************
                'B = MsgBox("このデータにはポンプ実績があります、使いますか？", vbYesNo + vbInformation)
                '******************************************************
                'Ver0.0.0 修正終了 1900/01/01 00:00
                '******************************************************
                b = vbYes
            Else
                b = vbNo
            End If
            If b = vbYes Then
                NG = FreeFile
                Open Wpath & "Pump.dat" For Output As #NG
                Do
                    Line Input #nf, buf
                    If InStr(buf, "INIT") > 0 Then
                        Close #NG
                        GoTo EXT
                    End If
                    Print #NG, buf
                Loop
            Else
                Pump_Full
                Do
                    Line Input #nf, buf
                    If InStr(buf, "INIT") > 0 Then
                        GoTo EXT
                    End If
                Loop
            End If
        End If
    End If
    GoTo EXT1
EXT:
    '******************************************************
    '初期水面形データ
    '******************************************************
    NG = FreeFile
    Open Wpath & "初期水面形.dat" For Output As #NG
    Print #NG, buf
    Do Until EOF(nf)
        Line Input #nf, buf
        Print #NG, buf
    Loop
EXT1:
    Close #NG
    Close #nf
    久地野と五条上流端流量
    LOG_Out " OUT Flood_Data_Read"
End Sub

'******************************************************************************
'サブルーチン：Input_Yosoku(irc As Boolean)
'処理概要：
'不定流計算結果を読み込む
'******************************************************************************
Sub Input_Yosoku(irc As Boolean)
    Dim i          As Integer
    Dim j          As Integer
    Dim k          As Integer
    Dim ns         As Integer
    Dim buf        As String
    Dim nf         As Integer
    Dim w          As Single
    Dim Froude_Max As Single
    On Error GoTo ERH1
    irc = True
    NT = 18
    ReDim HQ(3, nd, NT)                                                 '3=(1=H 2=Q 3=flood)
    ReDim MAX_H(nd)
    nf = FreeFile
    Open App.Path & "\WORK\newnskg2.u08" For Input As #nf

    Froude_Max = 0#
    For i = 1 To nd
        Line Input #nf, buf
        NDN(i) = Mid(buf, 25, 6)
        w = CSng(Mid(buf, 31, 10))                                      '水位
        HQ(1, i, 1) = w
        MAX_H(i) = w
        HQ(2, i, 1) = CSng(Mid(buf, 41, 10))                            '流量
        HQ(3, i, 1) = CSng(Mid(buf, 51, 10))                            'フルード数
        If HQ(3, i, 1) > Froude_Max Then Froude_Max = HQ(3, i, 1)
    Next i
    Froude = Froude_Max
    j = 1
    Do Until EOF(nf)
        j = j + 1
        If j > NT Then
            NT = NT + 1
            ReDim Preserve HQ(3, nd, NT)
        End If
        For i = 1 To nd
            Line Input #nf, buf
            NDN(i) = Mid(buf, 25, 6)
            w = CSng(Mid(buf, 31, 10))                                  '水位
            HQ(1, i, j) = w
            HQ(2, i, j) = CSng(Mid(buf, 41, 10))                        '流量
            HQ(3, i, j) = CSng(Mid(buf, 51, 10))                        'フルード数
            If w > MAX_H(i) Then MAX_H(i) = w
        Next i
    Loop
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    '******************************************************
    'ここからＤＢ用の予測値を取ろうとしたが
    'For i = 1 To 5                         '生の計算値なのでやめたしかも不等流時には
    '    ns = V_Sec_Num(i)                  'どうなっているかわからない
    '    k = 0
    '    For j = NT - 18 To NT
    '        YHK(i, k) = HQ(1, ns, j)
    '        k = k + 1
    '    Next j
    'Next i
    'Froude_Check irc
    'If Froude_Max > 1# Then
    '    Froude = 0.4
    'Else
    '    Froude = 0.03
    'End If
    '******************************************************
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    Close #nf
    On Error GoTo 0
    Exit Sub
ERH1:
    MsgBox "不定流計算結果を読み込み中エラーが発生した。" & vbCrLf & _
           "もう一度計算をお願いします、２回目以降もエラーが発生するときは" & vbCrLf & _
           "不定流計算の以上終了が考えられます、入力データをお確かめ下さい。", vbInformation
    On Error GoTo 0
    irc = False
    Close #nf
End Sub

'******************************************************************************
'サブルーチン：LOG_Out(msg As String)
'処理概要：
'******************************************************************************
Sub LOG_Out(msg As String)
    If LOF(Log_Num) > 3000000 Then
        Close #Log_Num
        Open_Log_File
    End If
    Print #Log_Num, Format(Now, "yyyy/mm/dd hh:nn:ss") & Space(2) & msg
End Sub

'******************************************************************************
'サブルーチン：Open_Log_File()
'処理概要：
'******************************************************************************
Sub Open_Log_File()
    Dim File    As String
    Dim L       As Long
    Log_Num = FreeFile
    File = App.Path & "\data\Log_file.dat"
    If Len(Dir(File)) > 0 Then
        L = FileLen(File)
        If L < 3000000 Then
            Open File For Append As #Log_Num
        Else
            Open File For Output As #Log_Num
        End If
    Else
        Open File For Output As #Log_Num
    End If
End Sub

'******************************************************************************
'関数；
'処理概要：TIMEC(dw As Date) As String
'******************************************************************************
Public Function TIMEC(dw As Date) As String
    TIMEC = Format(dw, "yyyy/mm/dd hh:nn")
End Function

'******************************************************************************
'サブルーチン；久地野と五条上流端流量()
'処理概要：TIMEC(dw As Date) As String
'******************************************************************************
Sub 久地野と五条上流端流量()
    Dim i     As Integer
    Dim j     As Integer
    Dim buf   As String
    Dim a(2)  As Single
    Dim b(2)  As Single
    Dim qk    As Single
    Dim qh    As Single
    Dim nf    As Integer
    Dim d     As Date
    LOG_Out " 久地野と五条上流端流量 In"
    '******************************************************
    '
    '******************************************************
    nf = FreeFile
    Open App.Path & "\data\HQ.DAT" For Input As #nf
    Line Input #nf, buf
    Line Input #nf, buf                                                 '久地野Ｈ－Ｑ式
    HQA(1) = CSng(Mid(buf, 1, 10))
    HQB(1) = CSng(Mid(buf, 11, 10))
    Line Input #nf, buf
    Line Input #nf, buf                                                 '春日Ｈ－Ｑ式
    HQA(2) = CSng(Mid(buf, 1, 10))
    HQB(2) = CSng(Mid(buf, 11, 10))
    Close #nf
    '******************************************************
    '
    '******************************************************
    nf = FreeFile
    Open Wpath & "OBSQ.DAT" For Output As #nf
    Print #nf, "    DATE    TIME    久地野      春日"
    '******************************************************
    '2001/11/14 12:10*******.**-------.--
    '******************************************************
    d = jsd
    For i = 1 To Now_Step
        If HO(6, i) > -50# Then
            qk = HQA(1) * (HO(6, i) + HQB(1)) ^ 2
        Else
            qk = -99#
        End If
        If HO(7, i) > -50# Then
            qh = HQA(2) * (HO(7, i) + HQB(2)) ^ 2
        Else
            qh = -99#
        End If
        buf = Format(d, "yyyy/mm/dd hh:nn") & Format(Format(qk, "#####0.000"), "@@@@@@@@@@") & _
                                              Format(Format(qh, "#####0.000"), "@@@@@@@@@@")
        Print #nf, buf
        d = DateAdd("h", 1, d)
    Next i
    Close #nf
    LOG_Out " 久地野と五条上流端流量 Out"
End Sub

'******************************************************************************
'サブルーチン：前回受信時刻チェック(Cat As String, irc As Long)
'仕様
'前回受け取った時刻から15分過ぎてもデータが受信できなかったら
'計算をスキップさせるようにした
'
'入力
'Cat......"KISYO"=気象庁計算時 "FRICS"=FRICS計算時
'
'******************************************************************************
Sub 前回受信時刻チェック(Cat As String, irc As Long)
    Dim da   As String
    Dim Dm   As String
    Dim dw   As Date
    Dim FLw  As String
    Dim nf   As Long
    Dim nn   As Long
    irc = True
    Select Case Cat
        Case "KISYO"
            '******************************************************
            '気象庁ナウキャストデータチェック
            '******************************************************
            FLw = Current_Path & "Oracletest\oraora\Data\F_MESSYU_10MIN_1.DAT"
            nf = FreeFile
            Open FLw For Input As #nf
            Line Input #nf, da
            Line Input #nf, Dm
            Close #nf
            dw = CDate(Dm)
            nn = DateDiff("n", dw, Now)
            If nn > 15 Then
                ADD_ERROR_Message "気象庁ナウキャストデータが入力されませんでした " & Yosoku_Time_K & " の計算をスキップします。"
                Data_Time_Rewrite da, FLw
                irc = False
            End If
            '******************************************************
            '気象庁実績雨量時刻
            '******************************************************
            nf = FreeFile
            FLw = Current_Path & "Oracletest\oraora\Data\P_MESSYU_10MIN.dat"
            Open FLw For Input As #nf
            Line Input #nf, da
            Line Input #nf, Dm
            Close #nf
            dw = CDate(Dm)
            nn = DateDiff("n", dw, Now)
            If nn > 15 Then
                ADD_ERROR_Message "気象庁実績雨量データが入力されませんでした " & Yosoku_Time_K & " の計算をスキップします。"
                Data_Time_Rewrite da, FLw
                irc = False
            End If
        Case "FRICS"
            '******************************************************
            'FRICS 実績雨量時刻
            '******************************************************
            nf = FreeFile
            FLw = Current_Path & "Oracletest\oraora\Data\P_RADAR.dat.dat"
            Open FLw For Input As #nf
            Line Input #nf, da
            Line Input #nf, Dm
            Close #nf
            dw = CDate(Dm)
            nn = DateDiff("n", dw, Now)
            If nn > 15 Then
                ADD_ERROR_Message "FRICS実績雨量データが入力されませんでした " & Yosoku_Time_F & " の計算をスキップします。"
                Data_Time_Rewrite da, FLw
                irc = False
            End If
            '******************************************************
            'FRICS 予測雨量時刻
            '******************************************************
            nf = FreeFile
            FLw = Current_Path & "Oracletest\oraora\Data\F_RADAR.dat.dat"
            Open FLw For Input As #nf
            Line Input #nf, da
            Line Input #nf, Dm
            Close #nf
            dw = CDate(Dm)
            nn = DateDiff("n", dw, Now)
            If nn > 15 Then
                ADD_ERROR_Message "FRICS予測雨量データが入力されませんでした " & Yosoku_Time_F & " の計算をスキップします。"
                Data_Time_Rewrite da, FLw
                irc = False
            End If
    End Select
End Sub
