Attribute VB_Name = "OraPump"
Option Explicit
Option Base 1
Public Pump()           As String
Public Bit_Num(16)      As Long
Public Pump_Stats(20)   As Pump

Type Pump
    Name    As String  'ポンプ名
    P_Code  As Long    'ポンプ場コード
    S_Num   As Long    'プログラム上順番コード
    P_Num   As Long    'ポンプ台数
    sv_Num  As Long    'sv番号
End Type

Sub Bit_Intial()
    Dim i As Long
    Dim j As Long
    Bit_Num(1) = 32768
    For j = 2 To 16
        i = j - 1
        Bit_Num(j) = Bit_Num(i) / 2
    Next j
End Sub

'ポンプのビット状態を調べる
'入力
'     n ------- 16ビットのポンプ状態
'出力
'    na() ----- 16個のアレーでポンプがオンの所に1が入る。
Sub Check_Bit(n As Long, na() As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    k = n
    m = 1
    For i = 1 To 16
        j = k And m
        If j > 0 Then na(i) = 1
        m = m * 2
    Next i
End Sub

'ポンプ実績最新取得時刻チェック
Sub Check_OWARI_PUMP(rc As Boolean)
    Exit Sub
    Dim nf  As Integer
    Dim n   As Long
    Dim d1  As Date
    Dim d2  As Date
    Dim d3  As Date
    Dim ans As Long
    Dim buf As String
    Dim irc As Boolean
    Dim ic  As Boolean
    nf = FreeFile
'    Debug.Print " Freefile="; nf
    Open App.Path & "\data\OWARI_PUMP.DAT" For Input As #nf
    Line Input #nf, buf
    d1 = CDate(buf)
    Close #nf
'    d3 = Format(Now, "yyyy/mm/dd hh") & ":00"
'    n = DateDiff("h", d1, d3)
'    If n > 25 Then
'        d1 = DateAdd("h", -25, d3) '前回終了が２５時間より前だったので取り込み開始時刻を変更
'    End If
    ORA_KANSOKU_JIKOKU_GET "OWARI_SV", d2, irc
    If irc = False Then
        rc = irc
        GoTo JUMP
    End If
    If d2 > d1 Then
'        n = DateDiff("h", d1, d2) + 1
'        If n > 100 Then '100は適当に決めた値、要するにif文に引っかからないようにした。2002/08/07 in YOKOHAMA
'            ans = MsgBox("追加で取得しようとしているポンプデータステップが２４ｈｒの" & vbCrLf & _
'                         "間隔があります。作業を継続しますか？" & vbCrLf & _
'                         "新規の洪水計算ではじめることをお進めします。" & vbCrLf & _
'                         "[はい]でこのジョブは終了します、[いいえ]で継続します。", vbInformation + vbYesNo)
'            If ans = vbYes Then
'                Close
'                End
'            End If
'        End If
        d1 = DateAdd("n", 10, d1)
        ORA_LOG "ポンプデータ取り込み開始 " & d1 & " から " & d2 & "まで"
        ORA_P_PUMP d1, d2, ic
        If Not ic Then
            ORA_LOG "オラクルデータベースよりポンプデータを取得しようとした時に" & vbCrLf & _
                    "エラーが発生しています。"
            GoTo JUMP
        Else
            ORA_LOG "ポンプデータ取り込み正常終了"
            ORA_LOG "ポンプデータ時刻書き込み開始 " & d2
            nf = FreeFile
            Open App.Path & "\data\OWARI_PUMP.DAT" For Output As #nf
            Print #nf, Format(d2, "yyyy/mm/dd hh:nn")
'            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
            Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
'            d1 = CDate(buf)
            Close #nf
            ORA_LOG "ポンプデータ時刻書き込み終了"
        End If
    End If
JUMP:
End Sub

'ポンプデータ取得
'入力
'     Name ----------- ポンプ場名
'       np ----------- このプログラム上のポンプ順番
'     Code ----------- ポンプ場コード
'       sv ----------- sv番号
'      N_P ----------- ポンプ数
'       d1 ----------- データ取得開始時刻
'       d2 ----------- データ取得終了時刻
'出力
'      Pump() -------- ポンプ状態が入る グローバル変数
'          rc -------- 完了状態
Sub Ora_OWARI_PUMP(Name As String, Code As Long, sv As Long, _
                   N_P As Long, np As Long, d1 As Date, d2 As Date, rc As Boolean)
    Dim i       As Long
    Dim j       As Long
    Dim n       As Long
    Dim m
    Dim F       As Long
    Dim P_MAX   As Single
    Dim P(16)   As Long
    Dim SQL     As String
    Dim ODATS   As String
    Dim ODATE   As String
    Dim dw      As Date
    Dim t       As Date
    Dim pw      As String
    Dim flag    As Long
    Dim jj      As Long
    Dim svc     As Long
    Dim sta     As Long
    ORA_LOG "IN    Ora_OWARI_PUMP"
    Const w = "1,1,1,1,1,1,1,1,1,1,"
    OracleDB.Label3 = "オラクルより" & Name & "データ取得中"
    OracleDB.Label3.Refresh
    'とりあえず初期値として全ポンプオンにする
    n = DateDiff("n", d1, d2) / 10 + 1
    pw = Mid(w, 1, N_P * 2)
    For j = 1 To n
        Pump(np, j) = pw
    Next j
    ODATS = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    ODATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
'    SQL = "SELECT * FROM oracle.OWARI_SV WHERE jikoku " & _
          "BETWEEN TO_DATE(" & ODATS & ") AND TO_DATE(" & ODATE & ")" & _
          " AND Station=" & Str(Code) & " AND sv_no=" & Str(sv)
    SQL = "SELECT * FROM oracle.OWARI_SV WHERE jikoku= TO_DATE(" & ODATS & ")"
    ORA_LOG " SQL=" & SQL
    ' SQLステートメントを指定してダイナセットを取得する
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG Name & "ポンプデータがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        ORA_LOG "SQL=" & SQL
        rc = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "オラクルより" & Name & "ポンプデータ取得失敗"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    dw = d1
    Do Until dynOra.EOF
        t = CDate(dynOra.Fields("jikoku").Value)
        n = DateDiff("n", d1, t) / 10 + 1
        m = AscB(dynOra.Fields("sv_data").Value)
        svc = dynOra.Fields("sv_no").Value
        sta = dynOra.Fields("Station").Value
        flag = dynOra.Fields("flag").Value
        ORA_LOG " Name=" & Name
        ORA_LOG " t=" & Format(t, "yyyy/mm/dd hh:nn")
        ORA_LOG " n=" & Str(n)
        ORA_LOG " m=" & Str(m)
        ORA_LOG " svc=" & Str(svc)
        ORA_LOG " sta=" & Str(sta)
        ORA_LOG " flag=" & Str(flag)
        Debug.Print " Name="; Name
        Debug.Print " t="; t
        Debug.Print " n="; n
        Debug.Print " m="; m
        Debug.Print " svc="; svc
        Debug.Print " sta="; sta
        Debug.Print " flag="; flag
        If flag = 0 Then
            pw = ""
            For i = 1 To N_P
                jj = m And Bit_Num(i)
                If jj > 0 Then
                    pw = pw & "1,"  'ポンプ稼動中
                Else
                    pw = pw & "0,"  'ポンプ停止中
                End If
            Next i
            Pump(np, n) = pw
        Else
            pw = ""
            For i = 1 To N_P
                pw = pw & "1,"
            Next i
            Pump(np, n) = pw
        End If
        dynOra.MoveNext
    Loop
    dynOra.Close
    ORA_LOG "OUT   Ora_OWARI_PUMP"
End Sub

'水場川ポンプ場用に作成したサブルーチンです
'
'
Sub Ora_OWARI_SUIBA_PUMP(d1 As Date, d2 As Date, rc As Boolean)
    Dim i       As Long
    Dim j       As Long
    Dim n       As Long
    Dim nn      As Long
    Dim m
    Dim F       As Long
    Dim P_MAX   As Single
    Dim P(16)   As Long
    Dim SQL     As String
    Dim ODATS   As String
    Dim ODATE   As String
    Dim dw      As Date
    Dim t       As Date
    Dim pw      As String
    Dim flag    As Long
    Dim jj      As Long
    Dim svc     As Long
    Dim sta     As Long
    Dim sv      As Long

    Const Name = "水場川"
    Const Code = 1017
    Const np = 18

    ORA_LOG "IN    Ora_OWARI_SUIBA_PUMP"

    OracleDB.Label3 = "オラクルより" & Name & "データ取得中"
    OracleDB.Label3.Refresh

    ODATS = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    ODATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"

    SQL = "SELECT * FROM oracle.OWARI_SV WHERE jikoku " & _
          "BETWEEN TO_DATE(" & ODATS & ") AND TO_DATE(" & ODATE & ")" & _
          " AND Station=" & Str(Code) & " AND sv_no BETWEEN 1 AND 4"

    SQL = "SELECT * FROM oracle.OWARI_SV WHERE jikoku= TO_DATE(" & ODATS & ")"

    Debug.Print " SQL="; SQL

    nn = DateDiff("n", d1, d2) / 10 + 1
    ReDim s_p(nn, 4) As Variant 'nn=時刻ステップ数  4=sv番号ごとのデータ
    '初期化
    For n = 1 To nn
        s_p(n, 1) = 1025
        s_p(n, 2) = 64
        s_p(n, 3) = 4100
        s_p(n, 4) = 256
    Next n

    ' SQLステートメントを指定してダイナセットを取得する
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)

    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG Name & "データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        ORA_LOG "SQL=" & SQL
        rc = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "オラクルより" & Name & "ポンプデータ取失敗"
        OracleDB.Label3.Refresh
        GoTo JUMP1
    End If

    dw = d1
    Do Until dynOra.EOF
        t = CDate(dynOra.Fields("jikoku").Value)
        n = DateDiff("n", d1, t) / 10 + 1
        m = AscB(dynOra.Fields("sv_data").Value)
        svc = dynOra.Fields("sv_no").Value
        sta = dynOra.Fields("Station").Value
        flag = dynOra.Fields("flag").Value
        ORA_LOG " Name=" & Name
        ORA_LOG " t=" & Format(t, "yyyy/mm/dd hh:nn")
        ORA_LOG " n=" & Str(n)
        ORA_LOG " m=" & Str(m)
        ORA_LOG " svc=" & Str(svc)
        ORA_LOG " sta=" & Str(sta)
        ORA_LOG " flag=" & Str(flag)
        Debug.Print " Name="; Name
        Debug.Print " t="; t
        Debug.Print " n="; n
        Debug.Print " m="; m
        Debug.Print " svc="; svc
        Debug.Print " sta="; sta
        Debug.Print " flag="; flag
        If flag = 0 And svc < 5 Then
            s_p(n, svc) = m
        End If
        dynOra.MoveNext
    Loop
    dynOra.Close

JUMP1:
    For n = 1 To nn '時刻数
        pw = ""
        For i = 1 To 4 'sv番号数
            Select Case i
                Case 1
                    jj = s_p(n, i) And 1024
                    If jj > 0 Then
                        pw = "1,"
                    Else
                        pw = "0,"
                    End If
                    jj = s_p(n, i) And 1
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
                Case 2
                    jj = s_p(n, i) And 64
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
                Case 3
                    jj = s_p(n, i) And 4096
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
                    jj = s_p(n, i) And 4
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
                Case 4
                    jj = s_p(n, i) And 256
                    If jj > 0 Then
                        pw = pw & "1,"
                    Else
                        pw = pw & "0,"
                    End If
            End Select
        Next i
        Pump(np, n) = pw
    Next n

    ORA_LOG "OUT   Ora_OWARI_SUIBA_PUMP"

End Sub
'
'ポンプデータを取得する
'
'
Sub ORA_P_PUMP(d1 As Date, d2 As Date, rc As Boolean)

    Dim i       As Long
    Dim j       As Long
    Dim n       As Long
    Dim m       As Long
    Dim Name    As String   'ポンプ所名
    Dim Code    As Long     'ポンプ所コード
    Dim sv      As Long     'sv番号
    Dim N_P     As Long     'ポンプ数
    Dim np      As Long     'ポンプ場の通し番号
    Dim ic      As Boolean

    n = DateDiff("n", d1, d2) / 10 + 1 '10分データの個数

    ReDim Pump(19, n)  '17=ポンプ場数  n=時刻ステップ数


    Name = "福田ポンプ所"
    Code = 2501
    sv = 1
    N_P = 4
    np = 1
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "下之一色ポンプ所"
    Code = 2502
    sv = 2
    N_P = 5
    np = 2
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "助光ポンプ所"
    Code = 2503
    sv = 3
    N_P = 4
    np = 3
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "富田ポンプ所"
    Code = 2504
    sv = 1
    N_P = 5
    np = 4
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "伏屋ポンプ所"
    Code = 2505
    sv = 1
    N_P = 3
    np = 5
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "万場ポンプ所"
    Code = 2506
    sv = 3
    N_P = 2
    np = 6
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "中小田井ポンプ所"
    Code = 2507
    sv = 1
    N_P = 4
    np = 7
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "第一平田ポンプ所"
    Code = 2508
    sv = 2
    N_P = 2
    np = 8
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "平田処理場ポンプ場"
    Code = 2509
    sv = 3
    N_P = 5
    np = 9
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "第二平田ポンプ所"
    Code = 2610
    sv = 1
    N_P = 2
    np = 10
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "上小田井ポンプ所"
    Code = 2611
    sv = 2
    N_P = 8
    np = 11
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "小場塚ポンプ場"
    Code = 2601
    sv = 1
    N_P = 4
    np = 12
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "二ツ杁ポンプ場"
    Code = 2602
    sv = 1
    N_P = 5
    np = 13
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "豊田川ポンプ場"
    Code = 2603
    sv = 1
    N_P = 5
    np = 14
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "堀江ポンプ場"
    Code = 2604
    sv = 1
    N_P = 6
    np = 15
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "土器野ポンプ場"
    Code = 2605
    sv = 1
    N_P = 4
    np = 16
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "電車川ポンプ場"
    Code = 2606
    sv = 1
    N_P = 3
    np = 17
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    Name = "水場川ポンプ場"
    Ora_OWARI_SUIBA_PUMP d1, d2, ic

    Name = "鴨田川ポンプ場"
    Code = 1082
    sv = 1
    N_P = 4
    np = 19
    Ora_OWARI_PUMP Name, Code, sv, N_P, np, d1, d2, ic

    rc = True
    Pump_TO_mdb d1, d2

End Sub
Sub Pump_Inital()

    Pump_Stats(1).Name = "福田ポンプ所"
    Pump_Stats(1).P_Code = 2501
    Pump_Stats(1).sv_Num = 1
    Pump_Stats(1).P_Num = 4
    Pump_Stats(1).S_Num = 1

    Pump_Stats(2).Name = "下之一色ポンプ所"
    Pump_Stats(2).P_Code = 2502
    Pump_Stats(2).sv_Num = 2
    Pump_Stats(2).P_Num = 5
    Pump_Stats(2).S_Num = 2

    Pump_Stats(3).Name = "助光ポンプ所"
    Pump_Stats(3).P_Code = 2503
    Pump_Stats(3).sv_Num = 3
    Pump_Stats(3).P_Num = 4
    Pump_Stats(3).S_Num = 3

    Pump_Stats(4).Name = "富田ポンプ所"
    Pump_Stats(4).P_Code = 2504
    Pump_Stats(4).sv_Num = 1
    Pump_Stats(4).P_Num = 5
    Pump_Stats(4).S_Num = 4

    Pump_Stats(5).Name = "伏屋ポンプ所"
    Pump_Stats(5).P_Code = 2505
    Pump_Stats(5).sv_Num = 1
    Pump_Stats(5).P_Num = 3  '1,3,4 2は最後に追加予定
    Pump_Stats(5).S_Num = 5

    Pump_Stats(6).Name = "万場ポンプ所"
    Pump_Stats(6).P_Code = 2506
    Pump_Stats(6).sv_Num = 3
    Pump_Stats(6).P_Num = 2
    Pump_Stats(6).S_Num = 6

    Pump_Stats(7).Name = "中小田井ポンプ所"
    Pump_Stats(7).P_Code = 2507
    Pump_Stats(7).sv_Num = 1
    Pump_Stats(7).P_Num = 5
    Pump_Stats(7).S_Num = 7

    Pump_Stats(8).Name = "第一平田ポンプ所"
    Pump_Stats(8).P_Code = 2508
    Pump_Stats(8).sv_Num = 2
    Pump_Stats(8).P_Num = 2
    Pump_Stats(8).S_Num = 8

    Pump_Stats(9).Name = "平田処理場内ポンプ場"
    Pump_Stats(9).P_Code = 2509
    Pump_Stats(9).sv_Num = 3
    Pump_Stats(9).P_Num = 5
    Pump_Stats(9).S_Num = 9

    Pump_Stats(10).Name = "第二平田ポンプ所"
    Pump_Stats(10).P_Code = 2610
    Pump_Stats(10).sv_Num = 1
    Pump_Stats(10).P_Num = 2
    Pump_Stats(10).S_Num = 10

    Pump_Stats(11).Name = "上小田井ポンプ所"
    Pump_Stats(11).P_Code = 2611
    Pump_Stats(11).sv_Num = 2
    Pump_Stats(11).P_Num = 8
    Pump_Stats(11).S_Num = 11

    Pump_Stats(12).Name = "小場塚ポンプ場"
    Pump_Stats(12).P_Code = 2601
    Pump_Stats(12).sv_Num = 1
    Pump_Stats(12).P_Num = 4
    Pump_Stats(12).S_Num = 12

    Pump_Stats(13).Name = "二ツ杁ポンプ場"
    Pump_Stats(13).P_Code = 2602
    Pump_Stats(13).sv_Num = 1
    Pump_Stats(13).P_Num = 5
    Pump_Stats(13).S_Num = 13

    Pump_Stats(14).Name = "豊田川ポンプ場"
    Pump_Stats(14).P_Code = 2603
    Pump_Stats(14).sv_Num = 1
    Pump_Stats(14).P_Num = 5
    Pump_Stats(14).S_Num = 14

    Pump_Stats(15).Name = "堀江ポンプ場"
    Pump_Stats(15).P_Code = 2604
    Pump_Stats(15).sv_Num = 1
    Pump_Stats(15).P_Num = 6
    Pump_Stats(15).S_Num = 15

    Pump_Stats(16).Name = "土器野ポンプ場"
    Pump_Stats(16).P_Code = 2605
    Pump_Stats(16).sv_Num = 1
    Pump_Stats(16).P_Num = 4
    Pump_Stats(16).S_Num = 16

    Pump_Stats(17).Name = "電車川ポンプ場"
    Pump_Stats(17).P_Code = 2606
    Pump_Stats(17).sv_Num = 1
    Pump_Stats(17).P_Num = 3
    Pump_Stats(17).S_Num = 17

    Pump_Stats(18).Name = "水場川ポンプ場"
    Pump_Stats(18).P_Code = 1017
    Pump_Stats(18).sv_Num = 1
    Pump_Stats(18).P_Num = 6
    Pump_Stats(18).S_Num = 18

    Pump_Stats(19).Name = "鴨田川ポンプ場"
    Pump_Stats(19).P_Code = 1082
    Pump_Stats(19).sv_Num = 1
    Pump_Stats(19).P_Num = 4
    Pump_Stats(19).S_Num = 19

End Sub
