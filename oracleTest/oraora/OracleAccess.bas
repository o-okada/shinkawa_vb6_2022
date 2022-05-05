Attribute VB_Name = "OracleAccess"
'******************************************************************************
'モジュール名：OracleAccess
'
'修正履歴：
'2015/08/04：オラクルデータベースのテーブル削除に対応し、下記テーブルへの
'アクセス処理部分をコメントアウトする。
'ARAIZEKI           ：洗堰水位／越流量
'F_MESSYU_10MIN_1   ：１０分間降水量予測メッシュ値（予測１時間）
'F_MESSYU_10MIN_2   ：１０分間降水量予測メッシュ値（予測６時間）
'F_RADAR            ：FRICS予測レーダ雨量情報
'P_MESSYU_10MIN     ：１０分間降水量実況メッシュ値
'P_RADAR            ：FRICS実況レーダー雨量情報
'YOHOU_TARGET_RIVER ：予報文対象河川
'YOHOUBUN           ：予報文作成履歴
'YOHOUBUNAN         ：予報文案作成履歴
'
'OracleAccess.ORA_Araizeki()を修正した。【01】
'※OracleDB.Check_Araizeki_Time()を修正すること。【01-01】
'
'OracleAccess.Dump_F_MESSYU_10MIN_1()を修正した。【02】
'※OracleAccess.ORA_F_MESSYU_10MIN_2()を修正すること。【02-01】
'※OracleAccess.ORA_F_MESSYU_10MIN_1()を修正すること。【02-02】
'
'OracleAccess.ORA_F_MESSYU_10MIN_1()を修正した。【03】
'※OracleAccess.ORA_F_MESSYU_10MIN_1()を修正すること。【03-01】
'※OracleDB.Check_F_MESSYU_10MIN_1_Time()を修正すること。【03-02】
'
'OracleAccess.ORA_F_MESSYU_10MIN_2()を修正した。【04】
'※OracleDB.Check_F_MESSYU_10MIN_2_Time()を修正すること。【04-01】
'
'OracleAccess.ORA_F_MESSYU_10MIN_20()を修正した。【05】
'※この修正に伴う影響はない。
'
'OracleAccess.ORA_F_RADAR()を修正した。【06】
'※OracleDB.Check_F_RADAR_Time()を修正すること【06-01】。
'
'OracleAccess.ORA_P_MESSYU_10MIN()を修正した。【07】
'※OracleDB.Check_P_MESSYU_10MIN_Time()を修正すること。【07-01】
'
'OracleAccess.ORA_P_MESSYU_1Hour()を修正した。【08】
'※OracleDB.Check_P_MESSYU_1HOUR_Time()を修正すること。【08-01】
'
'OracleAccess.ORA_P_RADAR()を修正した。【09】
'※OracleDB.Check_P_RADAR_Time()を修正すること。【09-01】
'
'OracleAccess.ORA_YOHOUBUNAN()を修正した。【10】
'※予報文テスト送信.Command1_Click()を修正すること。【10-01】
'
'******************************************************************************
Option Explicit
Option Base 1

'******************************************************************************
'その他のグローバル変数をセットする。
'******************************************************************************
Global jgd                As Date           '本番ではコメントにすること
Global LOG_N              As Integer        'ログ出力用番号
Global LOG_File           As String         'ログ出力用ファイル名
Global Dmp_N              As Integer        'データダンプファイル番号
'******************************************************************************
'OO4O関連のグローバル変数をセットする。
'******************************************************************************
'Global ssOra              As Object         '
Global dbOra              As Object ' OraDatabase    '
Global dynOra             As Object ' OraDynaset     '
Public gAdoCon As ADODB.Connection
Public gAdoRst As ADODB.Recordset
Global gbool_Start_Set    As Boolean        'データ取り込み中はTrue
Global gdate_oraTims      As Date           'オラクルデータ取り込み開始時刻
Global gdate_oraTime      As Date           'オラクルデータ取り込み終了時刻
'******************************************************************************
'予報文関係のグローバル変数をセットする。
'******************************************************************************
Global B1                 As String         '主文
Global B2                 As String         '
Global C1                 As String         'WRITE_TIME
Global C2                 As String         'FORECAST_KIND
Global C3                 As String         'FORECAST_KIND_CODE
Global C4                 As String         'ESTIMATE_TIME
Global C5                 As String         'ANNOUNCE_TIME
Global YHK(6, 18)         As Single         '
'******************************************************************************
'構造体をセットする。
'******************************************************************************
Type FRC
     ir As Long                             '
     ic As Long                             '
     m  As Long                             '
End Type

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public gDebugMode As String

Public Function GetConfigData(ByVal aSection As String, ByVal aKey As String, ByVal aFilename As String) As String

    Const intMaxSize As Integer = 255
    Dim strBuffer As String

    strBuffer = Space(intMaxSize)
    If GetPrivateProfileString(aSection, aKey, vbNullString, strBuffer, intMaxSize, aFilename) Then
        GetConfigData = SetNullCharCut(strBuffer)
    Else
        GetConfigData = vbNullString
    End If

End Function

Public Function SetNullCharCut(ByVal aChar As String) As String

    Dim intNullPos As Integer

    intNullPos = InStr(aChar, vbNullChar)
    If intNullPos > 0 Then
        SetNullCharCut = Left(aChar, intNullPos - 1)
    Else
        SetNullCharCut = aChar
    End If

End Function

'******************************************************************************
'サブルーチン：Bin2Int()
'処理概要：
'******************************************************************************
Sub Bin2Int(b As Variant, i As Integer, rc As Boolean)
    Dim nf    As Long
    Dim rec   As Long
    Dim F     As String
    On Error GoTo BinErorr
    F = App.Path & "\Pump.Bin"
    nf = FreeFile
    Open F For Binary As #nf
    rec = 1
    Put #nf, rec, b
    rec = 1
    Get #nf, rec, i
    Close #nf
    On Error GoTo 0
    rc = True
    Exit Sub
BinErorr:
    On Error GoTo 0
    ORA_LOG "バイナリ変換でエラーが発生しました。"
End Sub

'******************************************************************************
'サブルーチン：Dump_F_MESSYU_10MIN_1()
'処理概要：
'******************************************************************************
Sub Dump_F_MESSYU_10MIN_1(m As String, d As Date, data() As String)
    Dim i    As Long
    Dim j    As Long
    Dim k    As Long
    Dim bf   As String
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【02】
    '※OracleAccess.ORA_F_MESSYU_10MIN_2()を修正すること。【02-01】
    '※OracleAccess.ORA_F_MESSYU_10MIN_1()を修正すること。【02-02】
    '※既にコメントアウト済みであるが、コメントアウトを戻さないようにコメントを追加すること。
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【02】
    '******************************************************
    Print #Dmp_N, m; "  "; d
    k = 0
    For i = 1 To 5
        bf = " "
        For j = 1 To 5
        k = k + 1
           bf = bf & data(k) & " "
        Next j
        Print #Dmp_N, bf
    Next i
End Sub

'******************************************************************************
'サブルーチン：ORA_P_WATER()
'処理概要：
'水位データをデータベースより取得する
'観測所番号
'station IN( 1002,1015,1016,1017,1019,1020 )
'1002=日光川外水位
'1015=新川下之一色
'1016=大治
'1017=水場川外水位
'1019=久地野
'1020=春日
'副水位
'1076=下之一色
'1077=春日
'1079=水場川外水位
'"テレメータ水位受信"
'副水位補填のメッセージは長時間取得の場合最初の30件までメッセージを
'出力します、以降は補填はするがメッセージは出ません。
'******************************************************************************
Sub ORA_P_WATER(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
'    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w(6)         As Single              '主水位計データ
    Dim s(3)         As Single              '副水位計データ
    Dim dw           As Date
    Dim dt           As String
    Dim A1
    Dim A2
    Dim A3
    Dim A4
    Dim f1
    Dim nf           As Integer
    Dim buf          As String
    Dim msgD(100)    As Date
    Dim msg(100)     As String
    Dim msg_num      As Long
    Dim hw           As Single
    Dim rc           As Boolean
    Const Ksk = -99#
    On Error GoTo ORA_P_WATER_Error
    ic = False
    ORA_LOG "水位データ取得開始"
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'" '," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'" '," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "愛知県河川情報システムデータベース装置より水位データ取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
'    sql_SELECT = "SELECT * FROM oracle.P_WATER "
    sql_SELECT = "SELECT"
    sql_SELECT = sql_SELECT & "  obs_time"
    sql_SELECT = sql_SELECT & ", obs_sta_id"
    sql_SELECT = sql_SELECT & ", flag10"
    sql_SELECT = sql_SELECT & ", data10"
    sql_SELECT = sql_SELECT & "  FROM t_water_level_data"
    '******************************************************
    'WHERE
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
'    sql_WHERE = "WHERE station IN( 1002,1015,1016,1017,1019,1020,1076,1077,1079 ) AND jikoku BETWEEN TO_DATE(" & _
'                SDATE & ") AND TO_DATE(" & EDATE & ") ORDER BY jikoku"
    'sql_WHERE = "WHERE station IN( 2,16,17,18,20,21 ) and JIKOKU = TO_DATE(" & Sdate & ")"
    'Ver0.0.0 修正終了 1900/01/01 00:00
    sql_WHERE = " WHERE obs_sta_id IN(1012, 81, 201, 91, 71, 131, 80, 130, 240)"
    sql_WHERE = sql_WHERE & " AND obs_time BETWEEN " & SDATE
    sql_WHERE = sql_WHERE & " AND " & EDATE
    sql_WHERE = sql_WHERE & " ORDER BY obs_time"
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
'    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    Set gAdoRst = New ADODB.Recordset
    gAdoRst.CursorType = adOpenStatic
    gAdoRst.LockType = adLockReadOnly
    gAdoRst.Open SQL, gAdoCon, , , adCmdText
    If gAdoRst.EOF And gAdoRst.BOF Then
        ORA_LOG "水位観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        ORA_LOG "SQL=" & SQL
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'MsgBox "水位観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        'Ver0.0.0 修正終了 1900/01/01 00:00
        ic = False
        Call SQLdbsDeleteRecordset(gAdoRst)
        OracleDB.Label3 = "愛知県河川情報システムデータベース装置より水位データ取得失敗"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    nf = FreeFile
    Open App.Path & "\Data\DB_H.DAT" For Output As #nf
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Print #nf, " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk: w(6) = Ksk
    s(1) = Ksk: s(2) = Ksk: s(3) = Ksk
    gAdoRst.MoveFirst
    i = 0
    Timew = ""
    msg_num = 0
    Do
        buf = ""
        If Not gAdoRst.EOF Then A1 = Format(gAdoRst.Fields("obs_time").Value, "yyyy/mm/dd hh:nn")
        If Timew <> A1 And i > 0 Or gAdoRst.EOF Then
            '******************************************************
            '下之一色水位データが欠測かをチェックする。
            '******************************************************
            If w(2) = Ksk Then
                If s(1) <> Ksk Then
                    w(2) = s(1)
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "下之一色水位データが欠測しました。副水位計データで補填しました。"
                        msgD(msg_num) = DateAdd("s", 5, jgd)
                    End If
                Else
                    dw = CDate(A1)
                    光水位取得 dw, hw, "下之一色", rc
                    If rc Then
                        w(2) = hw
                        If msg_num < 100 Then
                            msg_num = msg_num + 1
                            msg(msg_num) = "下之一色水位観測局データの無線経由データが欠測しました。光回線経由の主水位計データで補填しました。"
                            msgD(msg_num) = DateAdd("s", 5, jgd)
                        End If
                    End If
                End If
            End If
            '******************************************************
            '春日水位データが欠測かをチェックする。
            '******************************************************
            If w(6) = Ksk Then
                If s(2) <> Ksk Then
                    w(6) = s(2)
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "春日水位データが欠測しました。副水位計データで補填しました。"
                        msgD(msg_num) = DateAdd("S", 6, jgd)
                    End If
                Else
                    dw = CDate(A1)
                    光水位取得 dw, hw, "春日", rc
                    If rc Then
                        w(6) = hw
                        If msg_num < 100 Then
                            msg_num = msg_num + 1
                            msg(msg_num) = "春日水位観測局データの無線経由データが欠測しました。光回線経由の主水位計データで補填しました。"
                            msgD(msg_num) = DateAdd("s", 6, jgd)
                        End If
                    End If
                End If
            End If
            '******************************************************
            '水場川外水位データが欠測かをチェックする。
            '******************************************************
            If w(4) = Ksk Then
                If s(3) <> Ksk Then
                    w(4) = s(3)
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "水場川外水位データが欠測しました。副水位計データで補填しました。"
                        msgD(msg_num) = DateAdd("S", 7, jgd)
                    End If
                Else
                    dw = CDate(A1)
                    光水位取得 dw, hw, "水場川外", rc
                    If rc Then
                        w(4) = hw
                        If msg_num < 100 Then
                            msg_num = msg_num + 1
                            msg(msg_num) = "水場川外水位観測局データの無線経由データが欠測しました。光回線経由の主水位計データで補填しました。"
                            msgD(msg_num) = DateAdd("s", 7, jgd)
                        End If
                    End If
                End If
            End If
            '******************************************************
            '副水位計のない観測所を光水位で補填する。
            '大治
            '******************************************************
            If w(3) = Ksk Then
                dw = CDate(A1)
                光水位取得 dw, hw, "大治", rc
                If rc Then
                    w(3) = hw
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "大治水位観測局データの無線経由データが欠測しました。光回線経由の主水位計データで補填しました。"
                        msgD(msg_num) = DateAdd("s", 8, jgd)
                    End If
                End If
            End If
            '******************************************************
            '副水位計のない観測所を光水位で補填する。
            '久地野
            '******************************************************
            If w(5) = Ksk Then
                dw = CDate(A1)
                光水位取得 dw, hw, "久地野", rc
                If rc Then
                    w(5) = hw
                    If msg_num < 100 Then
                        msg_num = msg_num + 1
                        msg(msg_num) = "久地野水位観測局データの無線経由データが欠測しました。光回線経由の主水位計データで補填しました。"
                        msgD(msg_num) = DateAdd("s", 9, jgd)
                    End If
                End If
            End If
            '******************************************************
            'MDBに書き込む。
            '******************************************************
            dw = CDate(Timew)
            dt = Format(dw, "yyyy/mm/dd hh:nn")
            MDB_Rst_H.Open "select * from .水位 where Time = '" & dt & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
            If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
                MDB_Rst_H.AddNew            '水位データを追加する。
            End If
            MDB_Rst_H.Fields("Time").Value = dt
            MDB_Rst_H.Fields("Minute").Value = Minute(dw)
            MDB_Rst_H.Fields("Tide").Value = w(1)
            MDB_Rst_H.Fields("下之一色").Value = w(2)
            MDB_Rst_H.Fields("大治").Value = w(3)
            MDB_Rst_H.Fields("水場川外").Value = w(4)
            MDB_Rst_H.Fields("久地野").Value = w(5)
            MDB_Rst_H.Fields("春日").Value = w(6)
            MDB_Rst_H.Update
            MDB_Rst_H.Close
            w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk: w(6) = Ksk
            s(1) = Ksk: s(2) = Ksk: s(3) = Ksk
            Timew = A1
            '******************************************************
            '関数をコールする。
            '******************************************************
            Pump_Check dt, dw, w()          'ポンプ停止水位のチェックを行う。
            If gAdoRst.EOF Then Exit Do
        End If
        If i = 0 Then Timew = A1
        i = i + 1
        A2 = gAdoRst.Fields("obs_sta_id").Value
        A3 = gAdoRst.Fields("obs_time").Value
        A4 = gAdoRst.Fields("data10").Value
        f1 = gAdoRst.Fields("flag10").Value
        buf = buf & Format(A1, "@@@@@@@@@@@@@@@@@@@@,")
        buf = buf & Format(Str(A2), "@@@@@@@@@@,")
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'buf = buf & Format(Str(A3), "@@@@@@@@@@")
        'Ver0.0.0 修正終了 1900/01/01 00:00
        buf = buf & Format(Str(A4), "@@@@@@@@@@@@@@@,")
        buf = buf & Format(Str(f1), "@@@@@")
        Print #nf, buf
        jgd = CDate(A3)
        ORA_LOG Format(A1, "yyyy/mm/dd hh:nn") & "  " & A2 & "  H=" & A4 & " f=" & Str(f1)
        If f1 = 0 Or f1 = 10 Or f1 = 40 Or f1 = 50 Or f1 = 60 Or f1 = 70 Then
            Select Case CInt(A2)
                 Case 1012                  '日光川外水位
                    w(1) = CSng(A4) * 0.01
                 Case 81                    '新川下之一色
                    w(2) = CSng(A4) * 0.01
                 Case 201                   '大治
                    w(3) = CSng(A4) * 0.01
                 Case 91                    '水場川外水位
                    w(4) = CSng(A4) * 0.01
                 Case 71                    '久地野
                    w(5) = CSng(A4) * 0.01
                 Case 131                   '春日
                    w(6) = CSng(A4) * 0.01
                 Case 80                    '新川下之一色副水位計
                    s(1) = CSng(A4) * 0.01
                 Case 130                   '春日副水位計
                    s(2) = CSng(A4) * 0.01
                 Case 240                   '水場川外水位副水位計
                    s(3) = CSng(A4) * 0.01
            End Select
        End If
        gAdoRst.MoveNext
        DoEvents
    Loop
    If msg_num > 0 Then
        For i = 1 To msg_num
            jgd = msgD(i)
            ORA_LOG msg(i)
            ORA_Message_Out "テレメータ水位受信", msg(i), 1
        Next i
    End If
    ic = True
    Close #nf
    Call SQLdbsDeleteRecordset(gAdoRst)
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "愛知県河川情報システムデータベース装置より水位データ取得終了"
    OracleDB.Label3.Refresh
    Exit Sub
ORA_P_WATER_Error:
    Dim strMessage As String
    strMessage = Err.Description
    ORA_LOG "愛知県河川情報システムデータベース装置より水位データ取得失敗"
    ORA_LOG strMessage
    On Error GoTo 0
    Call SQLdbsDeleteRecordset(gAdoRst)
    OracleDB.Label3 = "愛知県河川情報システムデータベース装置より水位データ取得失敗"
    OracleDB.Label3.Refresh
End Sub

Public Sub SQLdbsDeleteRecordset(aobjAdorst As ADODB.Recordset)
    On Error Resume Next
    If Not (aobjAdorst Is Nothing) Then
        If aobjAdorst.State = adStateOpen Then aobjAdorst.Close
    End If
    Set aobjAdorst = Nothing
End Sub

'******************************************************************************
'サブルーチン：ORA_DataBase_Close()
'処理概要：
'後始末する。
'******************************************************************************
Sub ORA_DataBase_Close()
    '******************************************************
    'oo4oの接続を解除する。
    '******************************************************
'    Set ssOra = Nothing
'    Set dbOra = Nothing
    On Error Resume Next
    If Not (gAdoCon Is Nothing) Then
        If gAdoCon.State = adStateOpen Then gAdoCon.Close
    End If
    Set gAdoCon = Nothing
    gDebugMode = vbNullString
End Sub

'******************************************************************************
'サブルーチン：ORA_DataBase_Connection()
'処理概要：
'******************************************************************************
Sub ORA_DataBase_Connection(ic As Boolean)
    '******************************************************
    'OO4O で Oracle に接続する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'On Error GoTo ERRHAND
    On Error Resume Next
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'セッションを作成する。
    '******************************************************
'    Set ssOra = CreateObject("OracleInProcServer.XOraSession")
'    If Err <> 0 Then
'        'Ver0.0.0 修正開始 1900/01/01 00:00
'        'MsgBox "愛知県オラクルデータベースに接続出来ません。" & Chr(10) & _
'               "CreateObject - Oracle oo4o エラー"
'        'Ver0.0.0 修正終了 1900/01/01 00:00
'        ORA_LOG "愛知県オラクルデータベースに接続出来ません。" & Chr(10) & _
'                "CreateObject - Oracle oo4o エラー" & Chr(10) & _
'                "10秒休憩します " & Now
'        Short_Break 10
'        GoTo ERRHAND
'    End If
    '******************************************************
    'サービス名（サーバ名）と ユーザ名/パスワード を指定する。
    '******************************************************
    
    Dim strConfigFile As String
    Dim strProvider As String
    Dim strServer As String
    Dim strDBS As String
    Dim strUID As String
    Dim strPWD As String
    Dim strConn As String
    gDebugMode = vbNullString
    strConfigFile = App.Path
    If Right(strConfigFile, 1) <> "\" Then strConfigFile = strConfigFile & "\"
    strConfigFile = strConfigFile & "dbsinfo.cfg"
    If Len(Dir(strConfigFile, vbNormal)) < 1 Then
        ORA_LOG "愛知県河川情報システムデータベース装置の接続情報ファイルがありません。" & Chr(10) & _
                "10秒休憩します " & Now
        Short_Break 10
        GoTo ERRHAND
    End If
    
    strProvider = GetConfigData("databases", "provider", strConfigFile)
    strServer = GetConfigData("databases", "server", strConfigFile)
    strDBS = GetConfigData("databases", "dbs", strConfigFile)
    strUID = GetConfigData("databases", "uid", strConfigFile)
    strPWD = GetConfigData("databases", "pwd", strConfigFile)
    gDebugMode = GetConfigData("databases", "debug", strConfigFile)
    If Len(strServer) < 1 Or Len(strDBS) < 1 Then
        ORA_LOG "愛知県河川情報システムデータベース装置の接続情報がありません。" & Chr(10) & _
                "10秒休憩します " & Now
        Short_Break 10
        GoTo ERRHAND
    End If
    
    strConn = vbNullString
    If Len(strProvider) > 0 Then
        strConn = strConn & "Provider="
        strConn = strConn & strProvider
        strConn = strConn & ";"
    End If
    strConn = strConn & "Data Source="
    strConn = strConn & strServer
    strConn = strConn & ";"
    strConn = strConn & "Initial Catalog="
    strConn = strConn & strDBS
    strConn = strConn & ";"
    strConn = strConn & "User ID="
    strConn = strConn & strUID
    strConn = strConn & ";"
    strConn = strConn & "Password="
    strConn = strConn & strPWD
    strConn = strConn & ";"
    
'    Set dbOra = ssOra.OpenDatabase("ORACLE", "oracle/oracle", 0&)
    Set gAdoCon = New ADODB.Connection
    gAdoCon.ConnectionTimeout = 60
    gAdoCon.CommandTimeout = 60
    gAdoCon.Open strConn
    If Err.Number <> 0 Then
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'MsgBox "愛知県オラクルデータベースに接続出来ません。" & vbCrLf & _
              Err & ": " & Error
        'Ver0.0.0 修正終了 1900/01/01 00:00
        ORA_LOG "愛知県河川情報システムデータベース装置と接続できません。" & Chr(10) & _
                "10秒休憩します " & Now
        Short_Break 10
        GoTo ERRHAND
    End If
   On Local Error GoTo 0
   ic = True
   Exit Sub
ERRHAND:
    Dim strMessage As String
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'If dbOra.LastServerErr <> 0 Then
'    If dbOra.LastServerErr <> 0 Then
'    'Ver0.0.0 修正終了 1900/01/01 00:00
'        strMessage = dbOra.LastServerErrText    'DB処理でエラーが発生したときの処理。
'    Else
        strMessage = Err.Description            'DB処理以外でエラーが発生したときの処理。
'    End If
    ORA_LOG strMessage
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'MsgBox strMessage, vbExclamation
    'Ver0.0.0 修正終了 1900/01/01 00:00
    ic = False
    On Local Error GoTo 0
    Call ORA_DataBase_Close
End Sub

'******************************************************************************
'サブルーチン：
'処理概要：
'気象庁10分予測メッシュデータ(10分雨量）
'６個足し算しないと時間雨量にならないので足し算をして
'ＭＤＢには時間流域平均雨量として格納する
'指定されたd1時刻に計算をした今後６０分の１０分ピッチの予測雨量
'合計するとd1時刻の１時間後の予測雨量になる。
'******************************************************************************
Sub ORA_F_MESSYU_10MIN_1(d1 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1kmメッシュ値
    Dim w2(250)      As Single              '2kmメッシュ値
    Dim dw           As String
    Dim df           As Date
    Dim dm           As Date
    Dim MM           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2次メッシュ番号
    Dim FM(25)       As String              'ＤＢ上2次メッシュ番号
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim rr(135)      As Single
    Dim Ytime        As String
    Dim tm(10)       As String
    Dim c            As Single
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【03】
    '※OracleAccess.ORA_F_MESSYU_10MIN_1()を修正すること。【03-01】
    '※OracleDB.Check_F_MESSYU_10MIN_1_Time()を修正すること。【03-02】
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【03】
    '******************************************************
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    OracleDB.Label3 = "オラクルより気象庁１０分予測レーダデータ雨量取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.F_MESSYU_10MIN_1 "
    '******************************************************
    'WHERE
    '******************************************************
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ver0.0.0 修正開始 1900/01/01 00:00
    sql_WHERE1 = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                               "'523666','523667','523760','523656','523646' ) AND " & _
                 "jikoku = TO_DATE(" & SDATE & ") "
    'sql_WHERE2 = " AND \yosoku_time = TO_DATE(" & EDATE & ") "
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'フィールド内容を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw, sql_WHERE, d2
    'd1 = "2002/06/14 21:10"
    'd2 = "2002/06/14 21:10"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
    '    " AND code IN( '533606','533607','523676','523677','523770','523666','523667','523760','523656','523646' )"
    'SQL = sql_SELECT & sql_WHERE & " order by jikoku"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = dynOra.Fields("jikoku").Value
    '    Tw = Tw & "  " & dynOra.Fields("code").Value
    '    Tw = Tw & "  " & dynOra.Fields("start_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("end_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("yosoku_time").Value
    '    Debug.Print Tw
    '    dynOra.MoveNext
    '    DoEvents
    'Loop
    'Ver0.0.0 修正終了 1900/01/01 00:00
    ic = True
    SQL = sql_SELECT & sql_WHERE1           '& sql_WHERE2
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    ORA_LOG "SQL=" & SQL
    If dynOra.EOF And dynOra.BOF Then
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'MsgBox "観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        'Ver0.0.0 修正終了 1900/01/01 00:00
        dynOra.Close
        ORA_LOG "気象庁レーダー予測 スキップ" & d1
        ORA_LOG "気象庁１０分予測雨量データスキップ時刻書き込み開始 " & d1
        dm = DateAdd("h", 1, d1)
        nf = FreeFile
        Open App.Path & "\data\F_MESSYU_10MIN_1.DAT" For Output As #nf
        Print #nf, Format(dm, "yyyy/mm/dd hh:nn")
        Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
        Close #nf
        ORA_LOG "気象庁１０分予測雨量データキップ時刻書き込み終了"
        GoTo SKIP
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'Set dynOra = Nothing
        'OracleDB.Label3 = "ＦＲＩＣＳ１０分予測レーダデータ雨量取失得敗"
        'OracleDB.Label3.Refresh
        'Exit Sub
        'Ver0.0.0 修正終了 1900/01/01 00:00
    End If
    For i = 1 To 10
        tm(i) = "000000"
    Next i
    Do Until dynOra.EOF
        Ytime = dynOra.Fields("yosoku_time").Value
        dw = dynOra.Fields("jikoku").Value
        df = CDate(dw)
        dw = Format(df, "yyyy/mm/dd hh:nn")
        MM = dynOra.Fields("code").Value
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'Debug.Print " dw="; dw; "  code="; MM; "  m="; m; "  Ytime="; Ytime
        'Ver0.0.0 修正終了 1900/01/01 00:00
        Select Case MM
            Case mes(1)
                ks = 1
                m = 1
            Case mes(2)
                ks = 26
                m = 2
            Case mes(3)
                ks = 51
                m = 3
            Case mes(4)
                ks = 76
                m = 4
            Case mes(5)
                ks = 101
                m = 5
            Case mes(6)
                ks = 128
                m = 6
            Case mes(7)
                ks = 151
                m = 7
            Case mes(8)
                ks = 176
                m = 8
            Case mes(9)
                ks = 201
                m = 9
            Case mes(10)
                ks = 226
                m = 10
        End Select
        Select Case Ytime
            Case "010"
                dm = DateAdd("n", 1, df)
                Mid(tm(m), 1, 1) = "1"
            Case "020"
                dm = DateAdd("n", 2, df)
                Mid(tm(m), 2, 1) = "2"
            Case "030"
                dm = DateAdd("n", 3, df)
                Mid(tm(m), 3, 1) = "3"
            Case "040"
                dm = DateAdd("n", 4, df)
                Mid(tm(m), 4, 1) = "4"
            Case "050"
                dm = DateAdd("n", 5, df)
                Mid(tm(m), 5, 1) = "5"
            Case "060"
                dm = DateAdd("n", 6, df)
                Mid(tm(m), 6, 1) = "6"
            Case Else
                ORA_LOG "IN ORA_F_MESSYU_10MIN_1  どうして個々にくるの？"
                'Ver0.0.0 修正開始 1900/01/01 00:00
                'MsgBox " どうして個々にくるの？"
                'Ver0.0.0 修正終了 1900/01/01 00:00
        End Select
        For i = ks To ks + 24
            j = i - ks + 1
            c = CSng(dynOra.Fields(FM(j)).Value)
            If c < 0# Then c = 0#           '気象庁が-1を送ってくる可能性がある。
            w2(i) = w2(i) + c
        Next i
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'Dump_F_MESSYU_10MIN_1 MES(m), df, dmp()
        'Ver0.0.0 修正終了 1900/01/01 00:00
        Mesh_2km_to_1km_cvt w2(), w1()
        dynOra.MoveNext
    Loop
    Mesh_To_Ryuiki w1(), rr(), irc
    dm = DateAdd("h", 1, d1)
    ORA_LOG "気象庁レーダー予測10分 " & dm
    For i = 1 To 10
        ORA_LOG Format(Str(i), "@@@") & " " & mes(i) & " " & tm(i)
    Next i
    '******************************************************
    'MDBをOPENする。
    '******************************************************
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    '******************************************************
    'MDBに書き込む。
    '******************************************************
    MDB_Rst_H.Open "select * from .気象庁レーダー予測_1 where Time = '" & Format(dm, "yyyy/mm/dd hh:nn") & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
    If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
        MDB_Rst_H.AddNew                    '存在していたらデータを追加する。
    End If
    MDB_Rst_H.Fields("Time").Value = Format(dm, "yyyy/mm/dd hh:nn")
    MDB_Rst_H.Fields("Minute").Value = Minute(dm)
    For i = 1 To 135
        i1 = Format(i, "###")
        MDB_Rst_H.Fields(i1).Value = rr(i)
    Next i
    MDB_Rst_H.Update
    MDB_Rst_H.Close
    ORA_LOG "気象庁１０分予測雨量データ時刻書き込み開始 " & d1
    nf = FreeFile
    Open App.Path & "\data\F_MESSYU_10MIN_1.DAT" For Output As #nf
    Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
    Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
    Close #nf
    ORA_LOG "気象庁１０分予測雨量データ時刻書き込み終了"
SKIP:
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルより気象庁レーダー予測終了"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'サブルーチン：ORA_F_MESSYU_10MIN_2()
'処理概要：
'気象庁正時予測メッシュデータ(10分雨量）６時間分
'６個足し算しないと時間雨量にならない
'ＭＤＢには１０分流域平均雨量として格納する
'******************************************************************************
Sub ORA_F_MESSYU_10MIN_2(d1 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1kmメッシュ値
    Dim w2(250)      As Single              '2kmメッシュ値
    Dim dw           As String
    Dim df           As Date
    Dim dm           As Date
    Dim dmc          As String
    Dim MM           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2次メッシュ番号
    Dim FM(25)       As String              'ＤＢ上2次メッシュ番号
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim rr(135)      As Single
    Dim sr(135)      As Single
    Dim rrr(36, 135) As Single
    Dim Ytime        As String
    Dim tm(36)       As String
    Dim m1           As Long
    Dim m2           As Long
    Dim m3           As Long
    Dim m4           As Long
    Dim dmp(25)      As String
    Dim c            As Single
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【04】
    '※OracleDB.Check_F_MESSYU_10MIN_2_Time()を修正すること。【04-01】
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【04】
    '******************************************************
    ic = True
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    OracleDB.Label3 = "オラクルより気象庁正時予測レーダデータ雨量取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.F_MESSYU_10MIN_2 "
    '******************************************************
    'WHERE
    '******************************************************
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ver0.0.0 修正開始 1900/01/01 00:00
    sql_WHERE1 = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                               "'523666','523667','523760','523656','523646' ) AND " & _
                 "jikoku = TO_DATE(" & SDATE & ") ORDER BY jikoku"
    'sql_WHERE2 = " AND \yosoku_time = TO_DATE(" & EDATE & ") "
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'フィールド内容を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw, sql_WHERE, d2
    'd1 = "2002/06/11 15:00"
    'd2 = "2002/06/11 15:00"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
    '    " AND code IN( '533606','533607','523676','523677','523770','523666','523667','523760','523656','523646' )"
    'SQL = sql_SELECT & sql_WHERE & " order by jikoku"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = dynOra.Fields("jikoku").Value
    '    Tw = Tw & "  " & dynOra.Fields("code").Value
    '    Tw = Tw & "  " & dynOra.Fields("start_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("end_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("yosoku_time").Value
    '    Debug.Print Tw
    '    dynOra.MoveNext
    '    DoEvents
    'Loop
    'Ver0.0.0 修正終了 1900/01/01 00:00
    SQL = sql_SELECT & sql_WHERE1 '& sql_WHERE2
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'MsgBox "観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        'Ver0.0.0 修正終了 1900/01/01 00:00
        ic = False
        dynOra.Close
        ORA_LOG "気象庁レーダー予測 スキップ" & d1
        ORA_LOG "気象庁正時予測雨量データスキップ時刻書き込み開始 " & d1
        nf = FreeFile
        Open App.Path & "\data\F_MESSYU_10MIN_2.DAT" For Output As #nf
        Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
        Close #nf
        ORA_LOG "気象庁正時予測雨量データキップ時刻書き込み終了"
        GoTo SKIP
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'Set dynOra = Nothing
        'OracleDB.Label3 = "気象庁１０分予測レーダデータ雨量取失得敗"
        'OracleDB.Label3.Refresh
        'Exit Sub
        'Ver0.0.0 修正終了 1900/01/01 00:00
    End If
    For j = 1 To 36
        tm(j) = "0000000000"
    Next j
    Do Until dynOra.EOF
        Erase w1
        For m4 = 1 To 10
            If m4 = 1 Then
                Ytime = dynOra.Fields("yosoku_time").Value
            Else
                If dynOra.Fields("yosoku_time").Value <> Ytime Then
                    ORA_LOG "気象庁正時予測データが不完全"
                    GoTo SKIP
                End If
            End If
            dw = dynOra.Fields("jikoku").Value
            df = CDate(dw)
            dw = Format(df, "yyyy/mm/dd hh")
            MM = dynOra.Fields("code").Value
            'Ver0.0.0 修正開始 1900/01/01 00:00
            'Debug.Print " dw="; dw; "  code="; MM; "  m="; m; "  Ytime="; Ytime
            'Ver0.0.0 修正終了 1900/01/01 00:00
            Select Case MM
                Case mes(1)
                    ks = 1
                    n = 1
                Case mes(2)
                    ks = 26
                    n = 2
                Case mes(3)
                    ks = 51
                    n = 3
                Case mes(4)
                    ks = 76
                    n = 4
                Case mes(5)
                    ks = 101
                    n = 5
                Case mes(6)
                    ks = 126
                    n = 6
                Case mes(7)
                    ks = 151
                    n = 7
                Case mes(8)
                    ks = 176
                    n = 8
                Case mes(9)
                    ks = 201
                    n = 9
                Case mes(10)
                    ks = 226
                    n = 10
            End Select
            For i = ks To ks + 24
                j = i - ks + 1
                c = CSng(dynOra.Fields(FM(j)).Value)
                If c < 0# Then c = 0#       '気象庁が-1を送ってくる可能性があるのでおまじないをする。
                w2(i) = c
            Next i
            'Ver0.0.0 修正開始 1900/01/01 00:00
            'Dump_F_MESSYU_10MIN_1 MM & " " & Ytime, df, dmp()
            'Ver0.0.0 修正終了 1900/01/01 00:00
            dynOra.MoveNext
            DoEvents
            Mesh_2km_to_1km_cvt w2(), w1()
            If IsNumeric(Ytime) Then
                dw = dw & ":" & Mid(Ytime, 1, 2)
                m = CInt(Mid(Ytime, 1, 2))
            Else
                ORA_LOG "IN ORA_F_MESSYU_10MIN_2  どうして個々にくるの？"
                'Ver0.0.0 修正開始 1900/01/01 00:00
                'MsgBox " どうしてここにくるの？"
                'Ver0.0.0 修正終了 1900/01/01 00:00
            End If
            If n < 10 Then
                Mid(tm(m), n, 1) = Trim(Str(n))
            Else
                Mid(tm(m), n, 1) = "A"
            End If
        Next m4
        ORA_LOG "  Date=" & d1 & "  Ytime=" & Ytime
        Mesh_To_Ryuiki w1(), rr(), irc
        For i = 1 To 135
            rrr(m, i) = rr(i)
        Next i
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'GoTo SKIP
        'Ver0.0.0 修正終了 1900/01/01 00:00
    Loop
    '******************************************************
    '
    '******************************************************
    For i = 1 To 36
        If tm(i) <> "123456789A" Then
            ORA_LOG " tm(" & Format(Trim(Str(i)), "@@") & ")= " & tm(i)
        End If
    Next i
    dm = DateAdd("h", 1, d1)
    For m1 = 1 To 31
        Debug.Print " 10MIN_2" & " d1="; Format(d1, "yyyy/mm/dd hh:nn") & " dm=" & Format(dm, "yyyy/mm/dd hh:nn")
        Erase sr
        For m2 = m1 To m1 + 5
            For m3 = 1 To 135
                sr(m3) = sr(m3) + rrr(m2, m3)
            Next m3
        Next m2
        '******************************************************
        'MDBをOPENする。
        '******************************************************
        Set MDB_Rst_H.ActiveConnection = MDB_Con
        '******************************************************
        'MDBに書き込む。
        '******************************************************
        dmc = Format(dm, "yyyy/mm/dd hh:nn")
        MDB_Rst_H.Open "select * from .気象庁レーダー予測_2 where Time = '" & dmc & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '存在していたらデータを追加する。
        End If
        MDB_Rst_H.Fields("Time").Value = dmc
        MDB_Rst_H.Fields("Minute").Value = Minute(dm)
        For i = 1 To 135
            i1 = Format(i, "###")
            MDB_Rst_H.Fields(i1).Value = sr(i)
        Next i
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        dm = DateAdd("n", 10, dm)
    Next m1
    ic = True
SKIP:
    dynOra.Close
    Set dynOra = Nothing
    OracleDB.Label3 = "オラクルより気象庁レーダー正時予測終了"
    OracleDB.Label3.Refresh
    Set MDB_Rst_H = Nothing
End Sub

'******************************************************************************
'サブルーチン：ORA_F_MESSYU_10MIN_20
'処理概要：
'気象庁正時予測メッシュデータ(10分雨量）６時間分
'６個足し算しないと時間雨量にならない
'ＭＤＢには１０分流域平均雨量として格納する
'******************************************************************************
Sub ORA_F_MESSYU_10MIN_20(d1 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1kmメッシュ値
    Dim w2(250)      As Single              '2kmメッシュ値
    Dim dw           As String
    Dim df           As Date
    Dim dm           As Date
    Dim dmc          As String
    Dim MM           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2次メッシュ番号
    Dim FM(25)       As String              'ＤＢ上2次メッシュ番号
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim rr(135)      As Single
    Dim sr(135)      As Single
    Dim rrr(36, 135) As Single
    Dim Ytime        As String
    Dim tm(36)       As String
    Dim m1           As Long
    Dim m2           As Long
    Dim m3           As Long
    Dim jj           As Long
    Dim jjj          As Long
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【05】
    '※この修正に伴う影響はない。
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【05】
    '******************************************************
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    OracleDB.Label3 = "オラクルより気象庁正時予測レーダデータ雨量取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.F_MESSYU_10MIN_2 "
    '******************************************************
    'WHERE
    '******************************************************
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ver0.0.0 修正開始 1900/01/01 00:00
    sql_WHERE1 = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                               "'523666','523667','523760','523656','523646' ) AND " & _
                 "jikoku = TO_DATE(" & SDATE & ") AND DETAIL = 2"
    'sql_WHERE2 = " AND \yosoku_time = TO_DATE(" & EDATE & ") "
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'フィールド内容を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw, sql_WHERE, d2
    'd1 = "2002/06/11 15:00"
    'd2 = "2002/06/11 15:00"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
    '    " AND code IN( '533606','533607','523676','523677','523770','523666','523667','523760','523656','523646' )"
    'SQL = sql_SELECT & sql_WHERE & " order by jikoku"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = dynOra.Fields("jikoku").Value
    '    Tw = Tw & "  " & dynOra.Fields("code").Value
    '    Tw = Tw & "  " & dynOra.Fields("start_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("end_time").Value
    '    Tw = Tw & "  " & dynOra.Fields("yosoku_time").Value
    '    Debug.Print Tw
    '    dynOra.MoveNext
    '    DoEvents
    'Loop
    'Ver0.0.0 修正終了 1900/01/01 00:00
    SQL = sql_SELECT & sql_WHERE1 '& sql_WHERE2
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'If dynOra.EOF And dynOra.BOF Then
    ''    MsgBox "観測データがデータベースに登録されていません。" & vbCrLf & _
    '             "時刻を確かめてください。"
    '    ic = False
    '    dynOra.Close
    '    ORA_LOG "気象庁レーダー予測 スキップ" & d1
    '    ORA_LOG "気象庁正時予測雨量データスキップ時刻書き込み開始 " & d1
    '    nf = FreeFile
    '    Open App.Path & "\data\F_MESSYU_10MIN_2.DAT" For Output As #nf
    '    Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
    '    Close #nf
    '    ORA_LOG "気象庁正時予測雨量データキップ時刻書き込み終了"
    '    GoTo SKIP
    ''    Set dynOra = Nothing
    ''    OracleDB.Label3 = "気象庁１０分予測レーダデータ雨量取失得敗"
    ''    OracleDB.Label3.Refresh
    ''    Exit Sub
    'End If
    'Ver0.0.0 修正終了 1900/01/01 00:00
    For j = 1 To 36
        tm(j) = "0000000000"
    Next j
    For jjj = 1 To 36                       'Do Until dynOra.EOF
    For jj = 1 To 10                        'Do Until dynOra.EOF
        Ytime = Format(jjj, "00") & "0"     'Ytime = dynOra.Fields("yosoku_time").Value
        dw = d1                             'dynOra.Fields("jikoku").Value
        df = CDate(dw)
        dw = Format(df, "yyyy/mm/dd hh")
        MM = mes(jj)                        'MM = dynOra.Fields("code").Value
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'Debug.Print " dw="; dw; "  code="; MM; "  m="; m; "  Ytime="; Ytime
        'Ver0.0.0 修正終了 1900/01/01 00:00
        Select Case MM
            Case mes(1)
                ks = 1
                n = 1
            Case mes(2)
                ks = 26
                n = 2
            Case mes(3)
                ks = 51
                n = 3
            Case mes(4)
                ks = 76
                n = 4
            Case mes(5)
                ks = 101
                n = 5
            Case mes(6)
                ks = 128
                n = 6
            Case mes(7)
                ks = 151
                n = 7
            Case mes(8)
                ks = 176
                n = 8
            Case mes(9)
                ks = 201
                n = 9
            Case mes(10)
                ks = 226
                n = 10
        End Select
        For i = ks To ks + 24
            j = i - ks + 1
            w2(i) = 1                       'dynOra.Fields(FM(j)).Value * 0.5
        Next i
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'dynOra.MoveNext
        'Ver0.0.0 修正終了 1900/01/01 00:00
        DoEvents
        Mesh_2km_to_1km_cvt w2(), w1()
        ORA_LOG "  Date=" & d1 & "  Ytime=" & Ytime
        Mesh_To_Ryuiki w1(), rr(), irc
        If IsNumeric(Ytime) Then
            dw = dw & ":" & Mid(Ytime, 1, 2)
            m = CInt(Mid(Ytime, 1, 2))
        Else
            ORA_LOG "IN ORA_F_MESSYU_10MIN_2  どうして個々にくるの？"
            'Ver0.0.0 修正開始 1900/01/01 00:00
            'MsgBox " どうしてここにくるの？"
            'Ver0.0.0 修正終了 1900/01/01 00:00
        End If
        If n < 10 Then
            Mid(tm(m), n, 1) = Trim(Str(n))
        Else
            Mid(tm(m), n, 1) = "A"
        End If
        For i = 1 To 135
            rrr(m, i) = rr(i)
        Next i
    Next jj                                 'Loop
    Next jjj                                'Loop
    '******************************************************
    '
    '******************************************************
    For i = 1 To 36
        If tm(i) <> "123456789A" Then
            ORA_LOG " tm(" & Format(Trim(Str(i)), "@@") & ")= " & tm(i)
        End If
    Next i
    dm = DateAdd("h", 1, d1)
    For m1 = 1 To 31
        Debug.Print " 10MIN_2" & " d1="; Format(d1, "yyyy/mm/dd hh:nn") & " dm=" & Format(dm, "yyyy/mm/dd hh:nn")
        Erase sr
        For m2 = m1 To m1 + 5
            For m3 = 1 To 135
                sr(m3) = sr(m3) + rrr(m2, m3)
            Next m3
            '******************************************************
            'MDBをOPENする。
            '******************************************************
            Set MDB_Rst_H.ActiveConnection = MDB_Con
            '******************************************************
            'MDBに書き込む。
            '******************************************************
            dmc = Format(dm, "yyyy/mm/dd hh:nn")
            MDB_Rst_H.Open "select * from .気象庁レーダー予測_2 where Time = '" & dmc & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
            If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
                MDB_Rst_H.AddNew            '存在していたらデータを追加する
            End If
            MDB_Rst_H.Fields("Time").Value = dmc
            MDB_Rst_H.Fields("Minute").Value = Minute(dm)
            For i = 1 To 135
                i1 = Format(i, "###")
                MDB_Rst_H.Fields(i1).Value = sr(i)
            Next i
            MDB_Rst_H.Update
            MDB_Rst_H.Close
        Next m2
        dm = DateAdd("n", 10, dm)
    Next m1
SKIP:
'    dynOra.Close
'    Set dynOra = Nothing
'    OracleDB.Label3 = "オラクルより気象庁レーダー正時予測終了"
'    OracleDB.Label3.Refresh
    Set MDB_Rst_H = Nothing
End Sub

'******************************************************************************
'サブルーチン：ORA_F_RADAR()
'処理概要：
'FRICSレーダー予測雨量
'FRICSレーダー予測雨量はデータ量が多いので多時間取得を止めて
'単時間取得のルーティンとした。
'******************************************************************************
Sub ORA_F_RADAR(d1 As Date, irc As Boolean)
    Dim SQL          As String
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim MDB_SQL      As String
    Dim SDATE        As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim ir           As Long
    Dim im           As Long
    Dim Timew        As String
    Dim w1(18, 315)  As Single              '1kmメッシュ値
    Dim w2(315)      As Single              '1kmメッシュ値
    Dim rr
    Dim MS           As Long
    Dim ruika        As Long
    Dim Dim2         As String
    Dim buf          As String
    Dim rrr(135)     As Single
    Dim dw           As Date
    Dim DC           As String
    Dim i1           As Long
    Dim nf           As Long
    Dim Minutew      As Long
    Dim Mesh         As String
    Dim MMS          As Long
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【06】
    '※OracleDB.Check_F_RADAR_Time()を修正すること。【06-01】
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【06】
    '******************************************************
    '使用する2次メッシュ番号
    '533607
    '533606
    '523770
    '523677
    '523676
    '523760
    '523667
    '523666
    '523656
    '523646
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'On Error GoTo ERR1
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Const Ksk = -99#
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "オラクルよりFRICS予測レーダデータ雨量取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.F_RADAR WHERE "
    '******************************************************
    'WHERE1
    '******************************************************
    sql_WHERE1 = "jikoku = TO_DATE(" & SDATE & ") AND "
    SQL = sql_SELECT & sql_WHERE1 & Dim2_WHERE2
    ORA_LOG " SQL= " & SQL
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG " SQL = " & SQL & " このデータはORACLE-DBにない"
        ORA_LOG "FRICSレーダー予測雨量 スキップ"
        OracleDB.Label3 = "FRICS予測レーダデータ雨量取得失敗"
        OracleDB.Label3.Refresh
        GoTo SKIP
    End If
    Erase w1
    MMS = 0
    Do Until dynOra.EOF
        ORA_LOG " jikoku     =" & dynOra.Fields("JIKOKU").Value
        ORA_LOG " DATA_STATUS=" & dynOra.Fields("DATA_STATUS").Value
        Mesh = dynOra.Fields("LATITUDE").Value & dynOra.Fields("LONGITUDE").Value & Format(dynOra.Fields("CODE").Value, "00")
        ORA_LOG " Mesh       =" & Mesh
        Minutew = dynOra.Fields("RUIKA_DATE").Value / 10
        buf = dynOra.Fields("RADAR").Value
        rr = Split(buf, ",")
        Select Case Mesh
            Case "533607"
                 MS = 1
            Case "533606"
                 MS = 2
            Case "523770"
                 MS = 3
            Case "523677"
                 MS = 4
            Case "523676"
                 MS = 5
            Case "523760"
                 MS = 6
            Case "523667"
                 MS = 7
            Case "523666"
                 MS = 8
            Case "523656"
                 MS = 9
            Case "523646"
                 MS = 10
            Case Else
                 GoTo NOP
        End Select
        MMS = MMS + MS
        For i = 1 To Dim2_mesh_Number(MS)
            ir = Dim2_To_315(MS, i).Rn
            im = Dim2_To_315(MS, i).Mn - 1
            If rr(im) > 250 Then
                w1(Minutew, ir) = 0
            Else
                w1(Minutew, ir) = CSng(rr(im))
            End If
        Next i
NOP:
        dynOra.MoveNext
    Loop
SKIP:
    dynOra.Close
    ORA_LOG " MMS        =" & Format(MMS, "#0") & "  = 990 でないとおかしい。"
    DC = Format(d1, "yyyy/mm/dd hh:nn")
    For ruika = 1 To 18
        For i = 1 To 315
            w2(i) = w1(ruika, i)
        Next i
        Mesh_To_Ryuiki w2(), rrr(), irc
        '**************************************************
        'MDBをOPENする。
        '**************************************************
        
        '**************************************************
        'MDBに書き込む。
        '**************************************************
        MDB_SQL = "select * from FRICSレーダー予測 where Time = '" & DC & "' AND Prediction_Minute =" & Str(ruika * 10)
        MDB_Rst_H.Open MDB_SQL, MDB_Con, adOpenDynamic, adLockOptimistic
        ORA_LOG "MDB FRICS予測レーダ予測オープン SQL=" & MDB_SQL
        ORA_LOG "MDB_Rst_H.BOF=" & MDB_Rst_H.BOF
        ORA_LOG "MDB_Rst_H.EOF=" & MDB_Rst_H.EOF
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                'なかったらデータを追加する。
            MDB_Rst_H.Fields("Time").Value = DC
            MDB_Rst_H.Fields("Prediction_Minute").Value = ruika * 10
        End If
        For i = 1 To 135
            MDB_Rst_H.Fields(i + 1).Value = rrr(i)
        Next
        ORA_LOG "MDB FRICS予測レーダ予値書き込み"
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        ORA_LOG "MDB FRICS予測レーダ予値書き込み終了 "
    Next ruika
    ORA_LOG "FRICS予測レーダデータ時刻書き込み開始 " & d1
    nf = FreeFile
    Open App.Path & "\data\F_RADAR.DAT" For Output As #nf
    Print #nf, Format(d1, "yyyy/mm/dd hh:nn")
    Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
    Close #nf
    ORA_LOG "FRICS予測レーダデータ時刻書き込み終了"
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルよりFRICSレーダー予測雨量データ取終了"
    OracleDB.Label3.Refresh
    irc = True
    On Error GoTo 0
    Exit Sub
ERR1:
    On Error GoTo 0
    On Error Resume Next
    If MDB_Rst_H.State = adStateOpen Then
        MDB_Rst_H.Close
    End If
    irc = False
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルよりFRICSレーダー予測雨量データ取り込み異常終了"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'サブルーチン：ORA_KANSOKU_JIKOKU_GET()
'処理概要：
'テーブル KANSOKU_JIKOU から最新情報を取り込む。
'******************************************************************************
Sub ORA_KANSOKU_JIKOKU_GET(TBL As String, dw As Date, ic As Boolean)
    Dim cDw   As String
    Dim SQL   As String
    Dim buf   As String
    Dim n     As Long
    If TBL <> "F_MESSYU_10MIN_2" Then
        SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU WHERE TABLE_NAME='" & TBL & "'"
    Else
        SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU WHERE TABLE_NAME='" & TBL & "' AND DETAIL = 2"
    End If
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw, i
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
    '    Debug.Print " Value=" & dynOra.Fields(i).Value
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'ＤＢ内容を取得する。
    '******************************************************
    Dim w1, w2, w3
    n = FreeFile
    Open App.Path & "\KANSOKU_JIKOKU.dat" For Output As #n
    w1 = dynOra.Fields(0).Value
    w2 = dynOra.Fields(1).Value
    w3 = dynOra.Fields(2).Value
    If IsNull(w1) Then
        ORA_LOG "Error IN  ORA_KANSOKU_JIKOKU_GET Field=(" & TBL & ")のテーブル参照時にNULLが帰ってきた"
        ORA_LOG "SQL= (" & SQL & ")"
        ic = False
        GoTo JUMP1
    End If
    buf = Format(w1, "yyyy/mm/dd hh:nn") & "  "
    buf = buf & Format(w2, "@@@@@@@@@@@@@@@") & "  "
    buf = buf & Format(w3, "yyyy/mm/dd hh:nn") & "  "
    Print #n, buf
    '******************************************************
    '
    '******************************************************
    ic = True
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'w3 = "2005/08/05 18:20"
    'Ver0.0.0 修正終了 1900/01/01 00:00
    dw = CDate(w3)
JUMP1:
    DoEvents
    Close #n
    dynOra.Close
    Set dynOra = Nothing
End Sub

'******************************************************************************
'サブルーチン：ORA_KANSOKU_JIKOKU_PUT()
'処理概要：
'テーブル KANSOKU_JIKOU に最新情報を書き込む。
'******************************************************************************
Sub ORA_KANSOKU_JIKOKU_PUT(TBL As String, dw As Date)
    Dim cDw   As String
    Dim SQL   As String
    Dim buf   As String
    Dim n     As Long
    SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU"     'WHERE TABLE_NAME=" & TBL
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw, i, n
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'ＤＢの内容を取得する。
    '******************************************************
    Dim w1, w2, w3
    n = FreeFile
    Open App.Path & "\KANSOKU_JIKOKU.dat" For Output As #n
    Do Until dynOra.EOF
        w1 = dynOra.Fields(0).Value
        w2 = dynOra.Fields(1).Value
        w3 = dynOra.Fields(2).Value
        buf = Format(w1, "yyyy/mm/dd hh:nn") & "  "
        buf = buf & Format(w2, "@@@@@@@@@@@@@@@") & "  "
        buf = buf & Format(w3, "yyyy/mm/dd hh:nn") & "  "
        Print #n, buf
        dynOra.MoveNext
    Loop
    '******************************************************
    '
    '******************************************************
    If dynOra.EOF Then
        dynOra.AddNew
    Else
        dynOra.Edit
    End If
    cDw = Format(Now, "yyyy/mm/dd hh:nn")
    dynOra.Fields("write_time").Value = cDw
    dynOra.Fields("table_name").Value = TBL
    cDw = Format(jgd, "yyyy/mm/dd hh:nn")
    dynOra.Fields("last_date_time").Value = cDw
    dynOra.Update
    dynOra.Close
    DoEvents
    Close #n
    dynOra.Close
    Set dynOra = Nothing
End Sub

'******************************************************************************
'サブルーチン：ORA_LOG()
'処理概要：
'******************************************************************************
Sub ORA_LOG(msg As String)
    OracleDB.List1.AddItem Format$(Now, "MM:DD:HH:NN:SS") & " " & msg
    OracleDB.List1.ListIndex = OracleDB.List1.NewIndex
    If OracleDB.List1.ListIndex > 30000 Then
        Close #LOG_N
        LOG_N = FreeFile
        Open App.Path & LOG_File For Output As #LOG_N
        OracleDB.List1.Clear
    End If
    Print #LOG_N, Format(Now, "yyyy/mm/dd hh:nn:ss") & "  " & msg
    OracleDB.Time_Disp = OracleDB.List1.ListIndex
    OracleDB.Time_Disp.Refresh
End Sub

'******************************************************************************
'サブルーチン：ORA_Message_Out
'処理概要：
'計算状況をＤＢに書き込む。
'******************************************************************************
Sub ORA_Message_Out(Place As String, msg As String, Lebel As Long)
    Exit Sub
    Dim i        As Long
    Dim SQL      As String
    Dim WHERE    As String
    Dim Code(2)  As String
    Dim Ndate    As String
    Dim dw       As Date
    Dim rc       As Boolean
    Dim Obs_Time As Long
    Code(1) = "1"                           '仮計算値
    Code(2) = "2"                           '計算不可
    ORA_LOG "IN  ORA_Message_Out"
    ORA_LOG "    msg=" & msg
    If msg = "" Then
        Exit Sub
    End If
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'If DBX_ora = False Then
    '    'オラクルサーバーにアップしない
    '    Exit Sub
    'End If
    'Exit Sub  '応急処置
    'For Obs_Time = 1 To 10
    '    ORA_DataBase_Connection rc
    '    If rc Then GoTo JUMP1
    'Next Obs_Time
    'ORA_LOG "    オラクルにつながらないのでメッセージ出力をあきらめます"
    'Exit Sub
'JUMP1:
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Obs_Time = 0
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'SQL = "SELECT * FROM oracle.CAL_MESSAGE"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw, n
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).name
    '    Debug.Print " Number=" & Format(str(i), "@@@") & " フィールド名="; Tw
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    On Error GoTo ErrOracle
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'dw = DateAdd("s", Obs_Time, jgd)
    dw = jgd
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Ndate = "'" & Format(dw, "yyyy/mm/dd hh:nn:ss") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Message.Show
    'Message.Label1 = "メッセージをDBアップ中"
    'Message.ZOrder 0
    'Message.Label1.Refresh
    'Ver0.0.0 修正終了 1900/01/01 00:00
    SQL = "SELECT * FROM oracle.CAL_MESSAGE WHERE jikoku= TO_DATE(" & Ndate & ") "
    ORA_LOG "    SQL=" & SQL
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'If dynOra.EOF Then
    'Ver0.0.0 修正終了 1900/01/01 00:00
        dynOra.AddNew
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Else
    '    dynOra.Edit
    'End If
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'データ書き込む。
    '******************************************************
    dynOra.Fields("WRITE_TIME").Value = Format(Now, "yyyy/mm/dd hh:nn")     '書き込み時刻
    dynOra.Fields("jikoku").Value = Format(dw, "yyyy/mm/dd hh:nn:ss")
    dynOra.Fields("river_no").Value = "85053002"
    dynOra.Fields("RAIN_KIND").Value = "01"
    dynOra.Fields("error_area").Value = Place                               '障害個所
    dynOra.Fields("message").Value = msg
    dynOra.Fields("cal_level").Value = 1                                    'Lebel
    dynOra.Update
    On Error GoTo 0
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'ORA_DataBase_Close
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Exit Sub
ErrOracle:
    '******************************************************
    'ここからエラー処理部分
    '******************************************************
    Dim strMessage As String
    If dbOra.LastServerErr <> 0 Then
        strMessage = dbOra.LastServerErrText                                'DB処理におけるエラーの処理。
    Else
        strMessage = Err.Description                                        'DB処理以外のエラーの処理。
    End If
    ORA_LOG "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    ORA_LOG "IN ORA_Message_Out " & strMessage
    ORA_LOG "     SQL=" & SQL
    ORA_LOG "###### ERROR ERROR ERROR ERROR ERROR ERROR ERROR ######"
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'ORA_DataBase_Close
    'Ver0.0.0 修正終了 1900/01/01 00:00
    On Error GoTo 0
End Sub

'******************************************************************************
'テーブル KANSOKU_JIKOU の内容を出力する
'******************************************************************************
Sub ORA_NEW_DATA_TIME()
    Dim cDw   As String
    Dim SQL   As String
    Dim buf   As String
    Dim n     As Long
    SQL = "SELECT * FROM oracle.KANSOKU_JIKOKU"
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim tw, i
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; tw
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'ＤＢの内容を取得する。
    '******************************************************
    Dim w1, w2, w3
    Do Until dynOra.EOF
        w1 = dynOra.Fields(0).Value
        w2 = dynOra.Fields(1).Value
        w3 = dynOra.Fields(2).Value
        buf = Format(w1, "yyyy/mm/dd hh:nn") & "  "
        buf = buf & Format(w2, "@@@@@@@@@@@@@@@") & "  "
        buf = buf & Format(w3, "yyyy/mm/dd hh:nn") & "  "
        Debug.Print buf
        dynOra.MoveNext
    Loop
    '******************************************************
    '
    '******************************************************
    dynOra.Close
    DoEvents
    dynOra.Close
    Set dynOra = Nothing
End Sub

'******************************************************************************
'サブルーチン：ORA_OWARI_WATER()
'処理概要：
'2005/05/16  光ケーブル経由の入力
'水位データをデータベースより取得する
'観測所番号
'station IN( 1015,1016,1017,1019,1020 )
'1015=新川下之一色
'1016=大治
'1017=水場川外水位
'1019=久地野
'1020=春日
'******************************************************************************
Sub ORA_OWARI_WATER(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w(5)         As Single
    Dim dw           As Date
    Dim dt           As String
    Dim A1
    Dim A2
    Dim A3
    Dim A4
    Dim f1
    Dim nf           As Integer
    Dim buf          As String
    Const Ksk = -99#
    ORA_LOG "光水位データ取得開始"
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "オラクルより光水位データ取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.OWARI_WATER "
    '******************************************************
    'WHERE
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    sql_WHERE = "WHERE station IN( 1015,1016,1017,1019,1020 ) AND jikoku BETWEEN TO_DATE(" & _
                SDATE & ") AND TO_DATE(" & EDATE & ") ORDER BY jikoku"
    'sql_WHERE = "WHERE station IN( 2,16,17,18,20,21 ) and JIKOKU = TO_DATE(" & Sdate & ")"
    'Ver0.0.0 修正終了 1900/01/01 00:00
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG "水位観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        ORA_LOG "SQL=" & SQL
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'MsgBox "水位観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        'Ver0.0.0 修正終了 1900/01/01 00:00
        ic = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "オラクルより水位データ取失敗"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    '******************************************************
    'MDBをOPENする。
    '******************************************************
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    nf = FreeFile
    Open App.Path & "\Data\DB_H.DAT" For Output As #nf
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk
    dynOra.MoveFirst
    i = 0
    Timew = ""
    Do
        buf = ""
        If Not dynOra.EOF Then A1 = Str(dynOra.Fields("jikoku").Value)
        If Timew <> A1 And i > 0 Or dynOra.EOF Then
            '******************************************************
            'MDBに書き込む。
            '******************************************************
            dw = CDate(Timew)
            dt = Format(dw, "yyyy/mm/dd hh:nn")
            MDB_Rst_H.Open "select * from .光水位 where Time = '" & dt & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
            If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
                MDB_Rst_H.AddNew            '水位データを追加する。
            End If
            MDB_Rst_H.Fields("Time").Value = dt
            MDB_Rst_H.Fields("Minute").Value = Minute(dw)
            MDB_Rst_H.Fields("下之一色").Value = w(1)
            MDB_Rst_H.Fields("大治").Value = w(2)
            MDB_Rst_H.Fields("水場川外").Value = w(3)
            MDB_Rst_H.Fields("久地野").Value = w(4)
            MDB_Rst_H.Fields("春日").Value = w(5)
            MDB_Rst_H.Update
            MDB_Rst_H.Close
            w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk
            Timew = A1
            'Ver0.0.0 修正開始 1900/01/01 00:00
            'Pump_Check dt, dw, w()
            'Ver0.0.0 修正終了 1900/01/01 00:00
            If dynOra.EOF Then Exit Do
        End If
        If i = 0 Then Timew = A1
        i = i + 1
        A2 = dynOra.Fields("station").Value
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'A3 = dynOra.Fields("water_flag").Value
        'Ver0.0.0 修正終了 1900/01/01 00:00
        A4 = dynOra.Fields("water_data").Value
        f1 = dynOra.Fields("flag").Value
        buf = buf & Format(A1, "@@@@@@@@@@@@@@@@@@@@,")
        buf = buf & Format(Str(A2), "@@@@@@@@@@,")
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'buf = buf & Format(Str(A3), "@@@@@@@@@@")
        'Ver0.0.0 修正終了 1900/01/01 00:00
        buf = buf & Format(Str(A4), "@@@@@@@@@@@@@@@,")
        buf = buf & Format(Str(f1), "@@@@@")
        Print #nf, buf
        ORA_LOG Format(A1, "yyyy/mm/dd hh:nn") & "  " & A2 & " H(cm)=" & A4
        If f1 = 0 Then
            Select Case CInt(A2)
                 Case 1015                  '新川下之一色
                    w(1) = CSng(A4) * 0.01
                 Case 1016                  '大治
                    w(2) = CSng(A4) * 0.01
                 Case 1017                  '水場川外水位
                    w(3) = CSng(A4) * 0.01
                 Case 1019                  '久地野
                    w(4) = CSng(A4) * 0.01
                 Case 1020                  '春日
                    w(5) = CSng(A4) * 0.01
            End Select
        End If
        dynOra.MoveNext
        DoEvents
    Loop
    ic = True
    Close #nf
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルより光水位データ取得終了"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'サブルーチン：ORA_P_RADAR()
'処理概要：
'FRICSレーダー実績雨量
'******************************************************************************
Sub ORA_P_RADAR(d1 As Date, d2 As Date, irc As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim ir           As Long
    Dim ic           As Long
    Dim MM           As Long
    Dim SDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1kmメッシュ値
    Dim dw           As Date
    Dim wk           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim buf
    Dim m            As Integer
    Dim w(5)
    Dim i1           As String
    Dim im           As String
    Dim rc           As Boolean
    Dim rr
    Dim rrr(135)     As Single
    Dim Times        As Long
    Dim Tim          As Long
    Dim MS           As Long
    Dim MMS          As Long
    Dim ds           As String
    Dim Mesh         As String
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【09】
    '※OracleDB.Check_P_RADAR_Time()を修正すること。【09-01】
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【09】
    '******************************************************
    '使用する2次メッシュ番号
    '533607
    '533606
    '523770
    '523677
    '523676
    '523760
    '523667
    '523666
    '523656
    '523646
    Const Ksk = -99#
    Times = DateDiff("n", d1, d2) / 10 + 1
    dw = d1
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'On Error GoTo ERR1
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    OracleDB.Label3 = "オラクルよりFRICS実績レーダデータ雨量取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT JIKOKU,DATA_STATUS,LATITUDE,LONGITUDE,CODE,RADAR FROM oracle.P_RADAR "
    For Tim = 1 To Times
        SDATE = "'" & Format(dw, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
        sql_WHERE1 = "WHERE JIKOKU= TO_DATE(" & SDATE & ") AND "
        Erase w1, rrr
        SQL = sql_SELECT & sql_WHERE1 & Dim2_WHERE2
        ORA_LOG " SQL= " & SQL
        '******************************************************
        'SQLステートメントを指定してダイナセットを取得する。
        '******************************************************
        Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
        '******************************************************
        'フィールド名を取得する。
        '******************************************************
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'Dim Tw
        'n = dynOra.Fields.Count
        'For i = 0 To n - 1
        '    Tw = dynOra.Fields(i).Name
        '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
        'Next i
        'Ver0.0.0 修正終了 1900/01/01 00:00
        If dynOra.EOF And dynOra.BOF Then
            ORA_LOG " SQL = " & SQL & " このデータはORACLE-DBにない"
            ORA_LOG "FRICSレーダー実績雨量 スキップ"
            OracleDB.Label3 = "FRICSレーダデータ実績雨量取得失敗"
            OracleDB.Label3.Refresh
            GoTo SKIP
        End If
        MMS = 0
        Do Until dynOra.EOF
            ORA_LOG " jikoku     =" & dynOra.Fields("JIKOKU").Value
            ORA_LOG " DATA_STATUS=" & dynOra.Fields("DATA_STATUS").Value
            Mesh = dynOra.Fields("LATITUDE").Value & dynOra.Fields("LONGITUDE").Value & Format(dynOra.Fields("CODE").Value, "00")
            ORA_LOG " Mesh       =" & Mesh
            buf = dynOra.Fields("RADAR").Value
            rr = Split(buf, ",")
            Select Case Mesh
                Case "533607"
                     MS = 1
                Case "533606"
                     MS = 2
                Case "523770"
                     MS = 3
                Case "523677"
                     MS = 4
                Case "523676"
                     MS = 5
                Case "523760"
                     MS = 6
                Case "523667"
                     MS = 7
                Case "523666"
                     MS = 8
                Case "523656"
                     MS = 9
                Case "523646"
                     MS = 10
                Case Else
                     GoTo NOP
            End Select
            MMS = MMS + MS
            For i = 1 To Dim2_mesh_Number(MS)
                ir = Dim2_To_315(MS, i).Rn
                im = Dim2_To_315(MS, i).Mn - 1
                If rr(im) > 250 Then
                    w1(ir) = 0
                Else
                    w1(ir) = rr(im)
                End If
            Next i
NOP:
            dynOra.MoveNext
        Loop
SKIP:
        dynOra.Close
        ORA_LOG " MMS        =" & Format(MMS, "#0")
        Mesh_To_Ryuiki w1(), rrr(), irc
        '**************************************************
        'MDBをOPENする。
        '**************************************************
        
        '**************************************************
        'MDBに書き込む。
        '**************************************************
        ds = Format(dw, "yyyy/mm/dd hh:nn")
        MDB_Rst_H.Open "select * from .FRICSレーダー実績 where Time = '" & ds & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '存在していたらデータを追加する。
            MDB_Rst_H.Fields("Time").Value = ds
            MDB_Rst_H.Fields("Minute").Value = Minute(ds)
        End If
        For i = 1 To 135
            i1 = Format(i, "###")
            MDB_Rst_H.Fields(i1).Value = rrr(i)
        Next
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        Erase rrr
        ORA_LOG "FRICS実績レーダデータ時刻書き込み開始 " & dw
        nf = FreeFile
        Open App.Path & "\data\P_RADAR.DAT" For Output As #nf
        Print #nf, Format(dw, "yyyy/mm/dd hh:nn")
        Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
        Close #nf
        ORA_LOG "FRICS実績レーダデータ実績時刻書き込み終了"
        dw = DateAdd("n", 10, dw)
    Next Tim
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルよFRICS実績レーダデータ実績取り込み終了"
    OracleDB.Label3.Refresh
    On Error GoTo 0
    Exit Sub
ERR1:
    If MDB_Rst_H.State = 1 Then
        MDB_Rst_H.Close
    End If
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルよりFRICS実績レーダデータ実績異常終了"
    OracleDB.Label3.Refresh
    On Error GoTo 0
End Sub

'******************************************************************************
'サブルーチン：ORA_Araizeki()
'処理概要：
'洗堰データをデータベースより取得する
' Number=  1 フィールド名=JIKOKU
' Number=  2 フィールド名=SUII
' Number=  3 フィールド名=ETURYU_NOW
' Number=  4 フィールド名=ETURYU_010
' Number=  5 フィールド名=ETURYU_020
' Number=  6 フィールド名=ETURYU_030
' Number=  7 フィールド名=ETURYU_040
' Number=  8 フィールド名=ETURYU_050
' Number=  9 フィールド名=ETURYU_060
' Number= 10 フィールド名=ETURYU_070
' Number= 11 フィールド名=ETURYU_080
' Number= 12 フィールド名=ETURYU_090
' Number= 13 フィールド名=ETURYU_100
' Number= 14 フィールド名=ETURYU_110
' Number= 15 フィールド名=ETURYU_120
' Number= 16 フィールド名=ETURYU_130
' Number= 17 フィールド名=ETURYU_140
' Number= 18 フィールド名=ETURYU_150
' Number= 19 フィールド名=ETURYU_160
' Number= 20 フィールド名=ETURYU_170
' Number= 21 フィールド名=ETURYU_180
' Number= 22 フィールド名=STATION
'******************************************************************************
Sub ORA_Araizeki(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w(6)         As Single
    Dim dw           As Date
    Dim dt           As String
    Dim f1
    Dim nf           As Integer
    Dim buf          As String
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【01】
    '※OracleDB.frm Check_Araizeki_Time()を修正すること。【01-01】
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【01】
    '******************************************************
    Const Ksk = -99#
    ORA_LOG "洗堰データ取得開始"
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "オラクルより洗堰データ取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.ARAIZEKI "
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'SQL = sql_SELECT & sql_WHERE
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'Dim Tw
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'フィールド内容を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw
    'd1 = "2002/06/24 11:10"
    'd2 = "2002/06/24 16:50"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") "
    'SQL = sql_SELECT & sql_WHERE & " order by jikoku"
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = ""
    '    For i = 0 To 22
    '        Tw = Tw & "  " & dynOra.Fields(i).Value
    '        DoEvents
    '    Next i
    '    Debug.Print Tw
    '    dynOra.MoveNext
    'Loop
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'WHERE
    '******************************************************
    sql_WHERE = "WHERE  jikoku BETWEEN TO_DATE(" & _
                SDATE & ") AND TO_DATE(" & EDATE & ") ORDER BY jikoku"
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        ORA_LOG "洗堰観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        ORA_LOG "SQL=" & SQL
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'MsgBox "洗堰観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        'Ver0.0.0 修正終了 1900/01/01 00:00
        ic = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "オラクルより洗堰データ取得失敗"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    '******************************************************
    'MDBをOPENする。
    '******************************************************
    Set MDB_Rst_H.ActiveConnection = MDB_Con
    w(1) = Ksk: w(2) = Ksk: w(3) = Ksk: w(4) = Ksk: w(5) = Ksk: w(6) = Ksk
    dynOra.MoveFirst
    Do Until dynOra.EOF
        If Not IsNull(dynOra.Fields("jikoku").Value) Then
            buf = ""
            dw = CDate(dynOra.Fields("jikoku").Value)
            Timew = Format(dw, "yyyy/mm/dd hh:nn")
            If IsNumeric(dynOra.Fields("eturyu_now").Value) Then
                w(1) = CSng(dynOra.Fields("eturyu_now").Value) * 0.001 '単位の確認 2002/06/24 18:00 Frics
            Else
                w(1) = 0#
            End If
            If IsNumeric(dynOra.Fields("eturyu_060").Value) Then
                w(2) = CSng(dynOra.Fields("eturyu_060").Value) * 0.001
            Else
                w(2) = 0#
            End If
            If IsNumeric(dynOra.Fields("eturyu_120").Value) Then
                w(3) = CSng(dynOra.Fields("eturyu_120").Value) * 0.001
            Else
                w(3) = 0#
            End If
            If IsNumeric(dynOra.Fields("eturyu_180").Value) Then
                w(4) = CSng(dynOra.Fields("eturyu_180").Value) * 0.001
            Else
                w(4) = 0#
            End If
            '******************************************************
            'MDBに書き込む。
            '******************************************************
            MDB_Rst_H.Open "select * from .洗堰 where Time = '" & Timew & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
            If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
                MDB_Rst_H.AddNew            '洗堰データを追加する。
            End If
            MDB_Rst_H.Fields("Time").Value = Timew
            MDB_Rst_H.Fields("Minute").Value = Minute(dw)
            MDB_Rst_H.Fields("Q0").Value = w(1)
            MDB_Rst_H.Fields("Q1").Value = w(2)
            MDB_Rst_H.Fields("Q2").Value = w(3)
            MDB_Rst_H.Fields("Q3").Value = w(4)
            MDB_Rst_H.Update
            MDB_Rst_H.Close
        End If
        dynOra.MoveNext
        DoEvents
    Loop
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルより洗堰データ取終了"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'サブルーチン：ORA_P_MESSYU_10MIN
'処理概要：
'気象庁10分実況メッシュデータ(10分雨量）
'６個足し算しないと時間雨量にならない
'ＭＤＢには１０分流域平均雨量として格納する
' Number=  0 フィールド名=WRITE_TIME
' Number=  1 フィールド名=JIKOKU
' Number=  2 フィールド名=COUNT
' Number=  3 フィールド名=SIZE_LN
' Number=  4 フィールド名=SIZE_LE
' Number=  5 フィールド名=SEKISAN_TIME
' Number=  6 フィールド名=TANI
' Number=  7 フィールド名=START_TIME
' Number=  9 フィールド名=TIME_SPAN
' Number= 10 フィールド名=YOSOKU_TIME
' Number= 11 フィールド名=CODE
' Number= 12 フィールド名=DATA_11
' Number= 13 フィールド名=DATA_12
' Number= 15 フィールド名=DATA_14
' Number= 16 フィールド名=DATA_15
' Number= 17 フィールド名=DATA_21
' Number= 18 フィールド名=DATA_22
' Number= 19 フィールド名=DATA_23
' Number= 20 フィールド名=DATA_24
' Number= 21 フィールド名=DATA_25
' Number= 22 フィールド名=DATA_31
' Number= 23 フィールド名=DATA_32
' Number= 25 フィールド名=DATA_34
' Number= 26 フィールド名=DATA_35
' Number= 27 フィールド名=DATA_41
' Number= 28 フィールド名=DATA_42
' Number= 29 フィールド名=DATA_43
' Number= 30 フィールド名=DATA_44
' Number= 31 フィールド名=DATA_45
' Number= 32 フィールド名=DATA_51
' Number= 33 フィールド名=DATA_52
' Number= 34 フィールド名=DATA_53
' Number= 35 フィールド名=DATA_54
' Number= 36 フィールド名=DATA_55
'******************************************************************************
Sub ORA_P_MESSYU_10MIN(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE1   As String
    Dim sql_WHERE2   As String
    Dim SQL          As String
    Dim SDATE        As String
    Dim EDATE        As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim Wdate        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1kmメッシュ値
    Dim w2(250)      As Single              '2kmメッシュ値
    Dim dw           As Date
    Dim dt           As String
    Dim Ntime        As Long
    Dim MM           As Long
    Dim nn           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2次メッシュ番号
    Dim FM(25)       As String              'ＤＢ上2次メッシュ番号
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim cr           As String
    Dim rr           As Single
    Dim rrr(135)     As Single
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【07】
    '※OracleDB.Check_P_MESSYU_10MIN_Time()を修正すること。【07-01】
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【07】
    '******************************************************
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    Const Ksk = -99#
    OracleDB.Label3 = "オラクルより気象庁１０分実況レーダデータ雨量取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.P_MESSYU_10MIN "
    '******************************************************
    'WHERE1
    '******************************************************
    sql_WHERE1 = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                                "'523666','523667','523760','523656','523646' ) AND "
    '******************************************************
    'フィールド名を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw
    'SQL = sql_SELECT
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'For i = 0 To n - 1
    '    Tw = dynOra.Fields(i).Name
    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
    'Next i
    'Ver0.0.0 修正終了 1900/01/01 00:00
    '******************************************************
    'フィールド内容を取得する。
    '******************************************************
    'Ver0.0.0 修正開始 1900/01/01 00:00
    'Dim Tw, sql_WHERE
    'd1 = "2002/06/10 19:00"
    'd2 = "2002/06/10 19:30"
    'SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'EDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    'sql_WHERE = " WHERE  jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
    '    "ORDER BY code,jikoku"
    'SQL = sql_SELECT & sql_WHERE
    'Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    'n = dynOra.Fields.Count
    'Do Until dynOra.EOF
    '    Tw = dynOra.Fields("jikoku").Value
    '    Tw = Tw & "  " & dynOra.Fields("code").Value
    '    Debug.Print Tw
    '    dynOra.MoveNext
    '    DoEvents
    'Loop
    'Ver0.0.0 修正終了 1900/01/01 00:00
    Ntime = DateDiff("n", d1, d2) / 10 + 1
    dw = d1
    For nn = 1 To Ntime
        '******************************************************
        'WHERE2
        '******************************************************
        Wdate = "'" & Format(dw, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
        sql_WHERE2 = "jikoku = TO_DATE(" & Wdate & ") "
        SQL = sql_SELECT & sql_WHERE1 & sql_WHERE2
        '******************************************************
        'SQLステートメントを指定してダイナセットを取得する。
        '******************************************************
        Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
        If dynOra.EOF And dynOra.BOF Then
            'Ver0.0.0 修正開始 1900/01/01 00:00
            'MsgBox "観測データがデータベースに登録されていません。" & vbCrLf & _
                     "時刻を確かめてください。" & vbCrLf & dw
            'Ver0.0.0 修正終了 1900/01/01 00:00
            ORA_LOG "気象庁レーダー実績データスキップ時刻書き込み開始 " & dt
            nf = FreeFile
            Open App.Path & "\data\P_MESSYU_10MIN.DAT" For Output As #nf
            Print #nf, dt
            Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
            Close #nf
            ORA_LOG "気象庁レーダー実績データスキップ時刻書き込み終了"
            ic = False
            dynOra.Close
            'Ver0.0.0 修正開始 1900/01/01 00:00
            'Set dynOra = Nothing
            'Ver0.0.0 修正終了 1900/01/01 00:00
            OracleDB.Label3 = "気象庁１０分実況レーダデータ雨量取失得敗"
            OracleDB.Label3.Refresh
            GoTo SKIP
        End If
        Erase w2
        m = 0
        Do Until dynOra.EOF
            m = m + 1
            'Ver0.0.0 修正開始 1900/01/01 00:00
            'dm = dynOra.Fields("jikoku").Value
            'Ver0.0.0 修正終了 1900/01/01 00:00
            MM = dynOra.Fields("code").Value
            Select Case MM
                Case mes(1)
                    ks = 1
                Case mes(2)
                    ks = 26
                Case mes(3)
                    ks = 51
                Case mes(4)
                    ks = 76
                Case mes(5)
                    ks = 101
                Case mes(6)
                    ks = 128
                Case mes(7)
                    ks = 151
                Case mes(8)
                    ks = 176
                Case mes(9)
                    ks = 201
                Case mes(10)
                    ks = 226
            End Select
            For i = ks To ks + 24
                j = i - ks + 1
                cr = dynOra.Fields(FM(j)).Value
                If IsNumeric(cr) Then
                    rr = CSng(cr)
                Else
                    rr = 0#
                End If
                If rr < 0# Then rr = 0#
                w2(i) = rr
            Next i
            dynOra.MoveNext
        Loop
        If m < 10 Then
            ORA_LOG "2次メッシュのどれかが取得できていない。" & dw
        End If
        Mesh_2km_to_1km_cvt w2(), w1()
        Mesh_To_Ryuiki w1(), rrr(), irc
        '**************************************************
        'MDBをOPENする。
        '**************************************************
        dt = Format(dw, "yyyy/mm/dd hh:nn")
        ORA_LOG " 気象庁レーダー実績 MDBに書き込み" & dw
        Set MDB_Rst_H.ActiveConnection = MDB_Con
        '**************************************************
        'MDBに書き込む。
        '**************************************************
        MDB_Rst_H.Open "select * from .気象庁レーダー実績 where Time = '" & dt & "' ; ", MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '存在していたらデータを追加する。
        End If
        MDB_Rst_H.Fields("Time").Value = dt
        MDB_Rst_H.Fields("Minute").Value = Minute(dw)
        For i = 1 To 135
            i1 = Format(i, "###")
            MDB_Rst_H.Fields(i1).Value = rrr(i)
            'Ver0.0.0 修正開始 1900/01/01 00:00
            'Debug.Print dt; " rrr="; rrr(i)
            'Ver0.0.0 修正終了 1900/01/01 00:00
        Next i
        MDB_Rst_H.Update
        MDB_Rst_H.Close
        ORA_LOG " 気象庁レーダー実績 MDBに書き込み終了" & dw
        ORA_LOG "気象庁レーダー実績データ時刻書き込み開始 " & dt
        nf = FreeFile
        Open App.Path & "\data\P_MESSYU_10MIN.DAT" For Output As #nf
        Print #nf, dt
        Print #nf, Format(Now, "yyyy/mm/dd hh:nn")
        Close #nf
        ORA_LOG "気象庁レーダー実績データ時刻書き込み終了"
SKIP:
        dw = DateAdd("n", 10, dw)
    Next nn
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルより気象庁レーダー実績取得終了"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'サブルーチン：ORA_P_MESSYU_1Hour
'処理概要：
'気象庁10分実況メッシュデータ(10分雨量）
'６個足し算しないと時間雨量にならない
'ＭＤＢには１０分流域平均雨量として格納する
'******************************************************************************
Sub ORA_P_MESSYU_1Hour(d1 As Date, d2 As Date, ic As Boolean)
    Dim sql_SELECT   As String
    Dim sql_WHERE    As String
    Dim SQL          As String
    Dim n            As Integer
    Dim i            As Long
    Dim j            As Long
    Dim k            As Long
    Dim SDATE        As String
    Dim EDATE        As String
    Dim Timew        As String
    Dim w1(315)      As Single              '1kmメッシュ値
    Dim w2(250)      As Single              '2kmメッシュ値
    Dim dw           As String
    Dim dm           As Date
    Dim MM           As Long
    Dim nf           As Integer
    Dim ks           As Integer
    Dim mes(10)      As String              '2次メッシュ番号
    Dim FM(25)       As String              'ＤＢ上2次メッシュ番号
    Dim i1           As String
    Dim j1           As String
    Dim buf          As String
    Dim irc          As Boolean
    Dim m            As Integer
    Dim rr(135)     As Single
    '******************************************************
    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【08】
    '※OracleDB.Check_P_MESSYU_1HOUR_Time()を修正すること。【08-01】
    '******************************************************
    Exit Sub
    '******************************************************
    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【08】
    '******************************************************
    mes(1) = "533606"
    mes(2) = "533607"
    mes(3) = "523676"
    mes(4) = "523677"
    mes(5) = "523770"
    mes(6) = "523666"
    mes(7) = "523667"
    mes(8) = "523760"
    mes(9) = "523656"
    mes(10) = "523646"
    n = 0
    For i = 1 To 5
        i1 = Format(i, "0")
        For j = 1 To 5
            j1 = Format(j, "0")
            n = n + 1
            FM(n) = "data_" & i1 & j1
        Next j
    Next i
    Const Ksk = -99#
    SDATE = "'" & Format(d1, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    EDATE = "'" & Format(d2, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
    OracleDB.Label3 = "オラクルより気象庁１０分実況レーダデータ雨量取得中"
    OracleDB.Label3.Refresh
    '******************************************************
    'SELECT
    '******************************************************
    sql_SELECT = "SELECT * FROM oracle.P_MESSYU_10MIN "
    '******************************************************
    'WHERE
    '******************************************************
    sql_WHERE = "WHERE code IN( '533606','533607','523676','523677','523770'," & _
                               "'523666','523667','523760','523656','523646' ) AND " & _
        "jikoku BETWEEN TO_DATE(" & SDATE & ") AND TO_DATE(" & EDATE & ") " & _
        "ORDER BY jikoku,code"
    SQL = sql_SELECT & sql_WHERE
    '******************************************************
    'SQLステートメントを指定してダイナセットを取得する。
    '******************************************************
    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
    If dynOra.EOF And dynOra.BOF Then
        'Ver0.0.0 修正開始 1900/01/01 00:00
        'MsgBox "気象庁１０分観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        'Ver0.0.0 修正終了 1900/01/01 00:00
        ORA_LOG "気象庁１０分観測データがデータベースに登録されていません。" & vbCrLf & _
                 "時刻を確かめてください。"
        ic = False
        dynOra.Close
        Set dynOra = Nothing
        OracleDB.Label3 = "気象庁１時間実況レーダデータ雨量取失得敗"
        OracleDB.Label3.Refresh
        Exit Sub
    End If
    Do Until dynOra.EOF
        For m = 1 To 10
            dw = dynOra.Fields("jikoku").Value
            MM = dynOra.Fields("code").Value
            Select Case MM
                Case mes(1)
                    ks = 1
                Case mes(2)
                    ks = 26
                Case mes(3)
                    ks = 51
                Case mes(4)
                    ks = 76
                Case mes(5)
                    ks = 101
                Case mes(6)
                    ks = 128
                Case mes(7)
                    ks = 151
                Case mes(8)
                    ks = 176
                Case mes(9)
                    ks = 201
                Case mes(10)
                    ks = 226
            End Select
            For i = ks To ks + 24
                j = k - ks + 1
                w2(i) = dynOra.Fields(FM(j)).Value * 0.1
            Next i
        Next m
        dynOra.MoveNext
        Mesh_2km_to_1km_cvt w2(), w1()
        Mesh_To_Ryuiki w1(), rr(), irc
        '**************************************************
        'MDBをOPENする。
        '**************************************************
        Set MDB_Rst_H.ActiveConnection = MDB_Con
        '**************************************************
        'MDBに書き込む。
        '**************************************************
        MDB_Rst_H.Open "select * from .気象庁レーダー実績 where Time = #" & dw & "# ; ", MDB_Con, adOpenDynamic, adLockOptimistic
        If MDB_Rst_H.EOF Or MDB_Rst_H.BOF Then
            MDB_Rst_H.AddNew                '存在していたらデータを追加する。
        End If
        dm = CDate(dw)
        MDB_Rst_H.Fields("Time").Value = dm
        MDB_Rst_H.Fields("Minute").Value = Minute(dm)
        For i = 1 To 135
            i1 = Format(i, "###")
            MDB_Rst_H.Fields(i1).Value = rr(i)
        Next i
        MDB_Rst_H.Update
        MDB_Rst_H.Close
    Loop
    dynOra.Close
    Set dynOra = Nothing
    Set MDB_Rst_H = Nothing
    OracleDB.Label3 = "オラクルより気象庁１時間実績データ取終了"
    OracleDB.Label3.Refresh
End Sub

'******************************************************************************
'サブルーチン：ORA_YOHOUBUNAN()
'処理概要：
'愛知県サーバーに予報文を書き込む。
'******************************************************************************
'Sub ORA_YOHOUBUNAN(Return_Code As Boolean)
'    Dim sql_SELECT   As String
'    Dim sql_WHERE    As String
'    Dim SQL          As String
'    Dim N_rec        As Long
'    Dim n            As Integer
'    Dim i            As Long
'    Dim SDATE        As String
'    Dim EDATE        As String
'    Dim jssd         As Date
'    Dim jeed         As Date
'    Dim Timew        As String
'    Dim f1           As String
'    Dim f2           As String
'    Dim B11          As String
'    Dim B12          As String
'    '******************************************************
'    'Ver1.0.0 修正開始 2015/08/04 O.OKADA【10】
'    '※予報文テスト送信.Command1_Click()を修正すること。【10-01】
'    '******************************************************
'    Exit Sub
'    '******************************************************
'    'Ver1.0.0 修正終了 2015/08/04 O.OKADA【10】
'    '******************************************************
'    jssd = CDate(C4)
'    jeed = DateAdd("n", 30, jssd)
'    SDATE = "'" & Format(jssd, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
'    EDATE = "'" & Format(jeed, "yyyy/mm/dd hh:nn") & "'," & "'" & "yyyy/mm/dd hh24:mi:ss" & "'"
'    '******************************************************
'    'SELECT
'    '******************************************************
'    sql_SELECT = "SELECT * FROM oracle.YOHOUBUNAN"
'    '******************************************************
'    'WHERE
'    '******************************************************
'    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ")"
'    'Ver0.0.0 修正開始 1900/01/01 00:00
'    SQL = sql_SELECT & sql_WHERE
'    'SQL = sql_SELECT
'    'Ver0.0.0 修正終了 1900/01/01 00:00
'    '******************************************************
'    'フィールド名を取得する。
'    '******************************************************
'    'Ver0.0.0 修正開始 1900/01/01 00:00
'    'Dim Tw
'    'n = RST_YB.Fields.Count
'    'For i = 0 To n - 1
'    '    Tw = RST_YB.Fields(i).Name
'    '    Debug.Print " Number=" & Format(Str(i), "@@@") & " フィールド名="; Tw
'    'Next i
'    'Ver0.0.0 修正終了 1900/01/01 00:00
'    '******************************************************
'    'SQLステートメントを指定してダイナセットを取得する。
'    '******************************************************
'    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
'    Dim nf As Integer
'    Dim buf As String
'    nf = FreeFile
'    Open App.Path & "\Data\DB_YB.DAT" For Output As #nf
'    If dynOra.EOF Then
'        dynOra.AddNew
'    Else
'        dynOra.Edit
'    End If
'    f1 = Format(CDate(C4), "d日h時m分")
'    f2 = Format(CDate(C4), "d日h時m分")
'    B12 = "　　新川の水位は" & f2 & "頃には､次のように見込まれます｡"
'    B11 = "　　新川の水位は" & f1 & "現在､次のとおりとなっています｡" & vbLf & _
'          "　　水場川外水位水位観測所［新川町大字阿原地内］で" & vbLf & _
'          "  　　　　　4.52m(急上昇中)" & vbLf & _
'          B12 & vbLf & _
'          "　　水場川外水位水位観測所［新川町大字阿原地内］で" & vbLf & _
'          "　　　　　　5.30m程度"
'    dynOra.Fields("WRITE_TIME").Value = C1                      '書き込み時刻
'    dynOra.Fields("DATA_KIND_CODE").Value = "フケンコウズイアン01"
'    dynOra.Fields("DATA_KIND").Value = "予報文案（水位部分）"
'    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
'    dynOra.Fields("SENDING_STATION").Value = "愛知県尾張建設事務所"
'    dynOra.Fields("APPOINTED_CODE").Value = ""
'    dynOra.Fields("ESTIMATE_TIME").Value = C4
'    dynOra.Fields("PRACTICE_FLG_CODE").Value = "40"             '"40"=予報  "99"=演習
'    dynOra.Fields("PRACTICE_FLG").Value = "予報"                '"演習"
'    dynOra.Fields("SEQ_NO").Value = ""
'    dynOra.Fields("ANNOUNCE_TIME").Value = C5
'    dynOra.Fields("RIVER_NAME").Value = "愛知県庄内川水系　新川"
'    dynOra.Fields("RIVER_NO_CODE").Value = "85053002"
'    dynOra.Fields("RIVER_NO").Value = "新川"
'    dynOra.Fields("RIVER_DIV_CODE").Value = "00"
'    dynOra.Fields("RIVER_DIV").Value = ""
'    dynOra.Fields("ANNOUNCE_NO").Value = ""
'    dynOra.Fields("FORECAST_KIND").Value = C2
'    dynOra.Fields("FORECAST_KIND_CODE").Value = C3
'    dynOra.Fields("BUNSHO1").Value = B1
'    dynOra.Fields("BUNSHO2").Value = B2
'    dynOra.Fields("BUNSHO3").Value = ""
'    dynOra.Fields("RAIN_KIND").Value = "01"
'    dynOra.Update
'    dynOra.Close
'    '******************************************************
'    '予報文対象河川
'    'SELECT
'    '******************************************************
'    sql_SELECT = "SELECT * FROM oracle.YOHOU_TARGET_RIVER"
'    '******************************************************
'    'WHERE
'    '******************************************************
'    sql_WHERE = " WHERE  ESTIMATE_TIME = TO_DATE(" & SDATE & ")"
'    SQL = sql_SELECT & sql_WHERE
'    Set dynOra = dbOra.DbCreateDynaset(SQL, 0&)
'    If dynOra.EOF Then
'        dynOra.AddNew
'    Else
'        dynOra.Edit
'    End If
'    dynOra.Fields("WRITE_TIME").Value = C1                      '書き込み時刻
'    dynOra.Fields("BUNAN_CODE").Value = "01"
'    dynOra.Fields("DATA_KIND_CODE").Value = "フケンコウズイアン01"
'    dynOra.Fields("SENDING_STATION_CODE").Value = "23001"
'    dynOra.Fields("ESTIMATE_TIME").Value = C4
'    dynOra.Fields("TRIVER_NAME").Value = "新川"
'    dynOra.Fields("TRIVER_NO_CODE").Value = "85053002"
'    dynOra.Fields("TRIVER_NO").Value = "新川"
'    dynOra.Fields("TRIVER_DIV_CODE").Value = "00"
'    dynOra.Fields("FORECAST_KIND").Value = C2
'    dynOra.Fields("FORECAST_KIND_CODE").Value = C3
'    dynOra.Fields("RAIN_KIND").Value = "01"
'    dynOra.Fields("OUT_NO").Value = 1
'    dynOra.Update
'    dynOra.Close
'    DoEvents
'    Close #nf
'    Set dynOra = Nothing
'End Sub

'******************************************************************************
'サブルーチン：WAIT_Minute()
'処理概要：
'******************************************************************************
Sub WAIT_Minute(m As Integer)
    Dim Start_Time, End_Time
    Start_Time = Timer
    End_Time = Start_Time + m
    Do While Timer < End_Time
        DoEvents
        If Timer - Start_Time > 1 Then
            OracleDB.Label3.Caption = "Oracle DBに接続できません、２分間休憩中 あと" & Format(End_Time - Timer, "###0") & "秒"
            Start_Time = Timer
            OracleDB.Time_Disp.Caption = " " + Format(Now, "yyyy年mm月dd日　hh時nn分ss秒")
            OracleDB.Time_Disp.Refresh
        End If
    Loop
    OracleDB.Label3.Caption = "取り込み待ち"
End Sub

Public Sub WaterDataNewTime(aObsTime As Date, aFlag As Boolean)

    Dim strSQL As String
    Dim strGetMinTime As String
    Dim strGetMaxTime As String
    Dim strNowTime As String
    Dim intMinute As Integer
    
    On Error GoTo WaterDataNewTime_Error
    
    aFlag = False
    strGetMinTime = vbNullString
    strGetMaxTime = vbNullString
    
    strSQL = "SELECT"
    strSQL = strSQL & "  MIN(latest_obs_time) AS min_time"
    strSQL = strSQL & ", MAX(latest_obs_time) AS max_time"
    strSQL = strSQL & "  FROM t_water_level_obs_sta_status"
    strSQL = strSQL & " WHERE obs_sta_id IN(1012, 81, 201, 91, 71, 131, 80, 130, 240)"
    
    Set gAdoRst = New ADODB.Recordset
    gAdoRst.CursorType = adOpenStatic
    gAdoRst.LockType = adLockReadOnly
    gAdoRst.Open strSQL, gAdoCon, , , adCmdText
    If Not gAdoRst.EOF Then
        If IsDate(gAdoRst!min_time) Then strGetMinTime = Format(gAdoRst!min_time, "yyyy/mm/dd hh:nn")
        If IsDate(gAdoRst!max_time) Then strGetMaxTime = Format(gAdoRst!max_time, "yyyy/mm/dd hh:nn")
    End If
    Call SQLdbsDeleteRecordset(gAdoRst)
    
    If Not IsDate(strGetMinTime) And Not IsDate(strGetMaxTime) Then
        ORA_LOG "Error IN  水位観測データ、最新観測時刻情報なし"
        ORA_LOG "SQL= (" & strSQL & ")"
    Else
        If DateDiff("n", strGetMinTime, strGetMaxTime) = 0 Then
            aObsTime = CDate(strGetMaxTime)
        Else
            strNowTime = Format(Now, "yyyy/mm/dd hh:nn")
            intMinute = DatePart("n", strNowTime) Mod 10
            strNowTime = Format(DateAdd("n", -(intMinute), strNowTime), "yyyy/mm/dd hh:nn")
            If intMinute >= 6 Or DateDiff("n", strGetMaxTime, strNowTime) > 0 Then
                aObsTime = CDate(strGetMaxTime)
            Else
                aObsTime = CDate(strGetMinTime)
            End If
        End If
        aFlag = True
    End If
    
    Exit Sub
WaterDataNewTime_Error:
    Dim strMessage As String
    strMessage = Err.Description
    ORA_LOG strMessage
    On Error GoTo 0

End Sub

Public Sub RadarMeshuDataNewTime(ByVal aTableName As String, aObsTime As Date, aFlag As Boolean)

    Dim strTableName As String
    Dim strSQL As String
    Dim strGetTime As String
    Const intJSTAddHour9 As Long = 540
    
    On Error GoTo RadarMeshuDataNewTime_Error
    
    aFlag = False
    strTableName = vbNullString
    strGetTime = vbNullString
    
    Select Case aTableName
        Case "VDXA70"
            strTableName = "t_excg_vdxa70"
        Case "VDXB70"
            strTableName = "t_excg_vdxb70"
        Case "VCXB70"
            strTableName = "t_excg_vcxb70"
        Case "VCXB71"
            strTableName = "t_excg_vcxb71"
        Case "VCXB75"
            strTableName = "t_excg_vcxb75"
        Case "VCXB76"
            strTableName = "t_excg_vcxb76"
        Case Else
            ORA_LOG "Error IN  メッシュデータのデータベーステーブルなし"
            ORA_LOG "テーブル名= (" & aTableName & ")"
            Exit Sub
    End Select
    
    strSQL = "SELECT last_data_time"
    strSQL = strSQL & " FROM t_excg_kansoku_jikoku"
    strSQL = strSQL & " WHERE table_name='" & strTableName & "'"
    
    Set gAdoRst = New ADODB.Recordset
    gAdoRst.CursorType = adOpenStatic
    gAdoRst.LockType = adLockReadOnly
    gAdoRst.Open strSQL, gAdoCon, , , adCmdText
    If Not gAdoRst.EOF Then
        If IsDate(gAdoRst!last_data_time) Then strGetTime = Format(DateAdd("n", intJSTAddHour9, gAdoRst!last_data_time), "yyyy/mm/dd hh:nn")
    End If
    Call SQLdbsDeleteRecordset(gAdoRst)
    
    If Not IsDate(strGetTime) Then
        ORA_LOG "Error IN  メッシュデータ、最新観測時刻情報なし"
        ORA_LOG "SQL= (" & strSQL & ")"
    Else
        aObsTime = CDate(strGetTime)
        aFlag = True
    End If
    
    Exit Sub
RadarMeshuDataNewTime_Error:
    Dim strMessage As String
    strMessage = Err.Description
    ORA_LOG strMessage
    On Error GoTo 0

End Sub
