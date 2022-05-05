Attribute VB_Name = "Make_Pump_Data"
Option Explicit
Option Base 1
Public Pump_Data(21)  As Pump_Stats
Type Pump_Stats
     name        As String 'ポンプ所名
     ability(8)  As Single '各ポンプの能力
     s_num       As Long   'ポンプ順番
     p_num       As Long   'ポンプ台数
     p_base      As Single '最低排水量(取り込みが出来ないポンプ施設を合計した値)
     max         As Single 'ポンプデータが未入力、欠測の時はmaxを使う
End Type
Public P_Hina(183) As String  'ポンプ雛型データ 注意183=データ分だけ
Public P_Ctl(19)   As P_C     '観測ポンプデータからポンプデータを作成する時のコントロールデータ
Type P_C
    op    As Long    '本プログラム上のポンプ順番号
    np    As Long    'ポンプ雛型データのポンプ番号
    pp    As Long    '雛型データ上の順番号
End Type
Public P_Hina_Flag As Boolean  'True=ポンプ雛型が取得できた時
Sub FULL_PUMP_OUT()

    Dim nf      As Long
    Dim i       As Long
    Dim Out_F   As String
    Dim msg(3)  As String
    Dim msgD(3) As Date
    Dim dw      As Date

    msg(1) = "": msg(2) = "": msg(3) = ""

    LOG_Out "In  FULL_PUMP_OUT"

    If HO(3, Now_Step) < H_Stand2(1, 1) And HO(3, Now_Step) > -20# Then
        msg(1) = "ポンプデータが取得できませんでした、下之一色観測局水位が「ポンプ停止水位」以下なので、流域ポンプ場の能力と湛水量の関係からデータを設定します。"
        LOG_Out "In  FULL_PUMP_OUT" & msg(1)
        msgD(1) = DateAdd("s", 21, jgd)
    End If
    If HO(5, Now_Step) < H_Stand2(3, 1) And HO(5, Now_Step) > -20# Then
        msg(2) = "ポンプデータが取得できませんでした、水場川外水位観測局が「ポンプ停止水位」以下なので、流域ポンプ場の能力と湛水量の関係からデータを設定します。"
        LOG_Out "In  FULL_PUMP_OUT" & msg(2)
        msgD(2) = DateAdd("s", 22, jgd)
    End If
    If HO(7, Now_Step) < H_Stand2(5, 1) And HO(7, Now_Step) > -20# Then
        msg(3) = "ポンプデータが取得できませんでした、春日観測局が「ポンプ停止水位」以下なので、流域ポンプ場の能力と湛水量の関係からデータを設定します。"
        LOG_Out "In  FULL_PUMP_OUT" & msg(3)
        msgD(3) = DateAdd("s", 23, jgd)
    End If
    dw = jgd
    For i = 1 To 3
        If msg(i) <> "" Then
            jgd = msgD(i)
            ORA_Message_Out "ポンプデータ受信", msg(i), 1
        End If
    Next i
    jgd = dw
    'ポンプ雛型データをそのまま使う
    nf = FreeFile
    Out_F = App.Path & "\WORK\Pump.dat"
    Open Out_F For Output As #nf
    For i = 1 To 183
        Print #nf, P_Hina(i)
    Next i
    Close #nf

    LOG_Out "Out FULL_PUMP_OUT"

End Sub
'
'
'G_Time ---- ポンプデータを作成する時の現時刻
'
'
'
Sub ポンプデータ作成(G_Time As Date)

    Dim i         As Long
    Dim j         As Long
    Dim k         As Long
    Dim m         As Long
    Dim n         As Long
    Dim buf       As String
    Dim dat(185)  As String
    Dim Out_F     As String
    Dim nf        As Long
    Dim ds        As Date
    Dim dw        As Date
    Dim mn        As Long
    Dim ip        As Long
    Dim Pump_W(19, 16) As Single '19=ポンプ数 16=時間ステップ
    Dim Pump_Obs(16)   As Single
    Dim SQL       As String
    Dim T         As String
    Dim p_name    As String
    Dim w
    Dim pt        As Single
    Dim base      As Single
    Dim maxp      As Single
    Dim d1        As String
    Dim d2        As String

    Dim Pump_Dat(22, 17) As Single

'    On Error GoTo ERHr10

    LOG_Out "In  ポンプデータ作成"

    Const N_Pump = 19  'ポンプ場数

    'ポンプをマックスで初期化
    For i = 1 To N_Pump
        pt = Pump_Data(i).max
        For j = 1 To 16
            Pump_W(i, j) = pt
        Next j
    Next i

    If P_Hina_Flag = False Then
        GoTo FULL_PUMP
    End If

    ds = DateAdd("h", -12, G_Time)
    mn = Minute(G_Time)
    d1 = Format(ds, "yyyy/mm/dd hh:nn")          'これと
    d2 = Format(G_Time, "yyyy/mm/dd hh:nn")      'これの定義が反対だったので修正2008/0/29
    dw = ds
    SQL = "SELECT * FROM ポンプ実績 WHERE " & _
          "TIME BETWEEN '" & d1 & "' AND '" & d2 & "' AND Minute=" & Format(mn, "#0")
    Rec_水文.Open SQL, Con_水文, adOpenDynamic, adLockOptimistic

    LOG_Out "In  ポンプデータ作成 SQL=" & SQL

    If Rec_水文.EOF Then
        LOG_Out " データが取得できなかったのでフルポンプデータを出力。"
        FULL_PUMP_OUT 'データが取得できなかったので雛型データを出力
        Rec_水文.Close
        Exit Sub
    End If

'--------- フィールド名を印刷する -------------------------------
'    n = Rec_水文.Fields.Count
'    For i = 0 To n - 1
'        buf = Rec_水文.Fields(i).name
'        Debug.Print " Number=" & str(i) & " フィールド名=" & buf
'    Next i
'---------------------------------------------------------------

    Do Until Rec_水文.EOF
        T = Rec_水文.Fields("Time").Value
        j = DateDiff("h", ds, T) + 1
        'ポンプ場数回
        For i = 1 To N_Pump
            p_name = Pump_Data(i).name '各ポンプ場の名前
            buf = Rec_水文.Fields(p_name).Value
'buf = "0,1,1,0,0,0,0,0,0,0,0,0,0,"
            base = Pump_Data(i).p_base
            w = Split(buf, ",")
            pt = base
            For m = 1 To Pump_Data(i).p_num '各ポンプ場のポンプ数
                pt = pt + Pump_Data(i).ability(m) * w(m - 1)
            Next m
            Pump_W(i, j) = pt
        Next i
        Rec_水文.MoveNext
    Loop
    Rec_水文.Close
    '現時刻継続で3時間予測
    For i = 1 To N_Pump
        pt = Pump_W(i, 13)
        For j = 14 To 16
            Pump_W(i, j) = pt
        Next j
    Next i

FULL_PUMP:

    'ポンプデータ作成
    For i = 1 To N_Pump
        n = P_Ctl(i).np
        For j = 1 To 16
            k = j + 1
            Pump_Dat(n, k) = Pump_Dat(n, k) + Pump_W(i, j)
        Next j
    Next i

    For i = 1 To N_Pump
        n = P_Ctl(i).np
        m = P_Ctl(i).pp
        k = InStr(P_Hina(m), "99") - 1
        d1 = Mid(P_Hina(m), 1, k) & "9999"
        For j = 2 To 10
            d1 = d1 & Format(Format(Pump_Dat(n, j), "#0.00"), "@@@@@")
        Next j
        d2 = Space(10)
        For j = 11 To 17
            d2 = d2 & Format(Format(Pump_Dat(n, j), "#0.00"), "@@@@@")
        Next j

        P_Hina(m) = d1
        P_Hina(m + 1) = d2

    Next i

    Out_F = App.Path & "\WORK\Pump.dat"
    n = FreeFile
    Open Out_F For Output As #n
    For i = 1 To 183
        Print #n, P_Hina(i)
    Next i
    Close #n

    LOG_Out "Out ポンプデータ作成"
    On Error GoTo 0
    Exit Sub

ERHr10:
    On Error GoTo 0
    LOG_Out " ポンプデータ作成中にエラーが発生したのでフルポンプデータを使う。"
    FULL_PUMP_OUT

End Sub
'
'ポンプ雛型データとコントロールデータを読む
'
'
'
Sub ポンプ雛型データ読み込み()

    Dim i      As Long
    Dim j      As Long
    Dim k
    Dim buf    As String
    Dim nf     As Long
    Dim F      As String

    On Error GoTo ERH3

    LOG_Out "In  ポンプ雛型データ読み込み"

    F = App.Path & "\Data\ポンプ雛型.txt"
    nf = FreeFile
    Open F For Input As #nf
    For i = 1 To 183  '183=データ数
        Line Input #nf, buf
        P_Hina(i) = buf
    Next i

    LOG_Out "データ作成コントロールデータ読み込み"
    Line Input #nf, buf
    Line Input #nf, buf
    For i = 1 To 19
        Line Input #nf, buf
            P_Ctl(i).op = CLng(Mid(buf, 1, 3))    '本プログラム上のポンプ順番号
            P_Ctl(i).np = CLng(Mid(buf, 14, 3))   'ポンプ雛型データのポンプ番号
            P_Ctl(i).pp = CLng(Mid(buf, 29, 3))   '雛型データ上の順番号
    Next i
    Close #nf
'    P_Hina_Flag = True
    P_Hina_Flag = False
    LOG_Out "Out ポンプ雛型データ読み込み大成功"
    On Error GoTo 0
    Exit Sub

ERH3:
    On Error GoTo 0
    ORA_Message_Out "ポンプデータ受信", "ポンプ雛型データの読み込みに失敗しています、フルポンプで計算します。", 1
    P_Hina_Flag = False

End Sub
Sub ポンプ能力表読み込み()

    Dim i        As Long
    Dim j        As Long
    Dim k        As Long
    Dim buf      As String
    Dim File     As String
    Dim nf       As Long
    Dim p        As Single

    LOG_Out "In   ポンプ能力表読み込み"

'    On Error GoTo EHR1

    File = App.Path & "\data\ポンプ能力表.dat"
    nf = FreeFile
    Open File For Input As #nf

    Line Input #nf, buf 'データタイトル1
    Line Input #nf, buf 'データタイトル2

    For i = 1 To 19 '19ポンプ場
        Line Input #nf, buf
        j = CLng(Mid(buf, 1, 2))
        Pump_Data(j).name = Mid(buf, 72, 7)         'ポンプ名
        Pump_Data(j).s_num = CLng(Mid(buf, 1, 2))   'ポンプ順番
        Pump_Data(j).p_num = CLng(Mid(buf, 6, 5))   'ポンプ数
        Pump_Data(j).p_base = CSng(Mid(buf, 56, 5)) 'ポンプベース排水量
        Pump_Data(j).max = CSng(Mid(buf, 61, 5))    'ポンプ最大排水量 欠測等データが得られなかった時に使う
Debug.Print " s_num="; Pump_Data(j).s_num
Debug.Print " Name="; Pump_Data(j).name
Debug.Print " p_num="; Pump_Data(j).p_num
Debug.Print " p_base="; Pump_Data(j).p_base
Debug.Print " max="; Pump_Data(j).max
        '各ポンプ号機毎の能力
        For k = 1 To Pump_Data(j).p_num  'ポンプ数
            p = CSng(Mid(buf, (k - 1) * 5 + 16, 5))
            Pump_Data(i).ability(k) = p
'            Debug.Print " k="; k; " ability="; p
        Next k
    Next i

    Close #nf

    P_Hina_Flag = True
    LOG_Out "Out  ポンプ能力表読み込み"

    On Error GoTo 0
    Exit Sub
EHR1:
    On Error GoTo 0
    ORA_Message_Out "ポンプデータ読み込み", "ポンプ能力データの読み込みに失敗しています、フルポンプで計算します。", 1
    
    P_Hina_Flag = False

End Sub
