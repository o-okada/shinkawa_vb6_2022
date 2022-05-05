Attribute VB_Name = "Radar_Rain"
Option Explicit
Option Base 1

Public Const RRYU = 135                     '流域数
Public Radar_YRain_File        As String    'RSHINK用レーダー予測流域平均雨量
Public Radar_Rain_File         As String    'RSHINK用レーダー実績流域平均雨量
Public IRADAR                  As Integer   'レーダー雨量がある時=1 ない時=0
Public JRADAR                  As Integer   'レーダー雨量を使う時=1 使わない時=0
Public Radar_File              As String    'レーダー雨量ファイル名
Public R_Thissen(20, 140)      As Single    'レーダー雨量用ティーセン係数(１流域最大２０メッシュ)
Public R_Meshu(20, 140)        As Integer   '流域雨量用メッシュ番号(１流域最大２０メッシュ)
Public R_T_Name(140)           As String    '流域名記号
Public rr()                    As Single    'レーダー実績流域平均雨量
Public RY()                    As Single    'レーダー予測流域平均雨量
Public RhY(5, 18)              As Single    'レーダー予測流域平均雨量HANS用

Public R_Ave(5, 500)           As Single    '基準地店上流流域平均雨量
Public R_Ave_N(135)            As Long      '基準地店上流流域平均雨量作成コントロール
Public R_Ave_Num               As Integer   '流域数

Public JMA_Num                 As Long
'
'バイナリサーチ
'
'
Sub Find_Rname(Xname As String, num As Long)

    Dim i1   As Long
    Dim i2   As Long
    Dim i3   As Long
    Dim j    As Long

    i1 = 1
    i2 = RRYU

f1:
    i3 = Int(i1 + i2) / 2
    If Xname > R_T_Name(i3) Then
        i1 = i3
    Else
        i2 = i3
    End If
    If Xname = R_T_Name(i3) Then
        num = i3
        Exit Sub
    End If
    If i2 - i1 <= 1 Then
        If Xname = R_T_Name(i1) Then
            num = i1
            Exit Sub
        Else
            num = i2
            Exit Sub
        End If
    End If
    GoTo f1

End Sub
'
'気象庁雨量取得データチェック
'
'
Sub JMA_File_Open()

    Dim File    As String
    Dim L       As Long

    JMA_Num = FreeFile
    File = App.Path & "\data\気象庁雨量データ取得状態チェック.dat"
    If Len(Dir(File)) > 0 Then
        L = FileLen(File)
        If L < 3000000 Then
            Open File For Append As #JMA_Num
        Else
            Open File For Output As #JMA_Num
        End If
    Else
        Open File For Output As #JMA_Num
    End If

End Sub
Sub JMA_OUT(msg As String)

    If LOF(JMA_Num) > 3000000 Then
        Close #JMA_Num
        JMA_File_Open
    End If

    Print #JMA_Num, Format(Now, "yyyy/mm/dd hh:nn:ss") & "|" & msg

End Sub
'**************************************************************
'洗堰越流データ取得
'
'ds=希望開始時刻
'de=希望終了時刻
'
'
'dw=取得最終時刻
'
'*************************************************************
Sub MDB_洗堰(ds As Date, de As Date, irc As Long)

    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim b      As String
    Dim SQL    As String
    Dim mi     As String
    Dim nd1    As String
    Dim dw     As String
    Dim dew    As Date
    Dim w      As Single

    Dim ConR   As New ADODB.Recordset

    LOG_Out "IN  MDB_洗堰"

    k = DateDiff("h", ds, de) + 1

    For i = 1 To k + 3
        HO(2, i) = 0#
    Next i

    mi = Format(Minute(de), "00")

    SQL = "select * from 洗堰 where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' and Minute = " & mi & " order by Time"

    ConR.Open SQL, Con_水文, adOpenKeyset, adLockReadOnly

    If ConR.EOF Then
        LOG_Out "IN MB_洗堰データ取得できず"
        LOG_Out Format(ds, "yyyy/mm/dd hh:nn") & " ～ " & Format(de, "yyyy/mm/dd hh:nn")
        ConR.Close
        irc = False
        Exit Sub
    End If

    nd1 = Format(jgd, "yyyy/mm/dd hh:nn")
    j = 0
    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        i = DateDiff("h", ds, dw) + 1
        w = ConR.Fields("Q0").Value    '実績越流量
        HO(2, i) = w * 0.01
        If dw = nd1 Then
            w = ConR.Fields("Q1").Value  '1時間後予測越流量
            HO(2, i + 1) = w * 0.01
            w = ConR.Fields("Q2").Value  '2時間後予測越流量
            HO(2, i + 2) = w * 0.01
            w = ConR.Fields("Q3").Value  '3時間後予測越流量
            HO(2, i + 3) = w * 0.01
            j = 1
            Exit Do
        End If
        ConR.MoveNext
    Loop
    ConR.Close

    If j = 1 Then
        irc = 0
    Else
        '現時刻のデータが無かったので10分前を取りに行く 2006/03/31 15:31 In FRICS YOKOHAMA DC
        dew = DateAdd("n", -10, de)
        SQL = "select * from 洗堰 where Time ='" & Format(dew, "yyyy/mm/dd hh:nn") & "'"
        ConR.Open SQL, Con_水文, adOpenKeyset, adLockReadOnly
        i = Now_Step
        If ConR.EOF Then
            '10分前も無かった
            irc = 2
            HO(2, i + 1) = 0#  '予測越流量
            HO(2, i + 2) = 0#  '予測越流量
            HO(2, i + 3) = 0#  '予測越流量
            ConR.Close
            LOG_Out "Out MDB_洗堰 挫折"
            Exit Sub
        Else
            '10分前があった
            i = DateDiff("h", ds, de) + 1
            HO(2, i + 1) = ConR.Fields("Q1").Value  '予測越流量
            HO(2, i + 2) = ConR.Fields("Q2").Value  '予測越流量
            HO(2, i + 3) = ConR.Fields("Q3").Value  '予測越流量
            irc = 1
            ConR.Close
        End If
    End If

    LOG_Out "Out MDB_洗堰"

End Sub
'**************************************************************
'FRICS実績雨量取得
'
'ds=希望開始時刻
'de=希望終了時刻
'
'
'dw=取得最終時刻
'
'*************************************************************
Sub MDB_FRICSレーダー実績(ds As Date, de As Date, dw As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Long
    Dim b      As String
    Dim SQL    As String
    Dim mi     As String
    Dim ConR   As New ADODB.Recordset

    LOG_Out "IN MB_FRICSレーダー実績 " & ds & "～ " & de

    ReDim rr(500, 140)

    mi = Format(Minute(de), "00")

    SQL = "select * from FRICSレーダー実績 where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' and Minute = " & mi & " order by Time"

    ConR.Open SQL, Con_水文, adOpenKeyset, adLockReadOnly

    If ConR.EOF Then
        LOG_Out "IN MB_FRICSレーダー実績指定データ取得できず"
        LOG_Out Format(ds, "yyyy/mm/dd hh:nn") & " ～ " & Format(de, "yyyy/mm/dd hh:nn")
        ConR.Close
        irc = False
        Exit Sub
    End If

    i = 0
    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        i = DateDiff("h", ds, dw) + 1
        Debug.Print "FRICS i=" & Format(i, "000") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
        For j = 1 To 135
            b = Format(j, "###")
            rr(i, j) = ConR.Fields(b).Value    '流域雨量
        Next j
        ConR.MoveNext
    Loop
    irc = True
    ConR.Close


End Sub
'**************************************************************
'FRICS予測降雨を取得
'
'ds=希望開始時刻
'
'
' RY(3, 140)
'dw=取得最終時刻
'
'*************************************************************
Sub MDB_FRICSレーダー予測(ds As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim m      As Long
    Dim n      As Long
    Dim dur    As Date
    Dim dwr    As Date
    Dim dw     As Date
    Dim b      As String
    Dim SQL    As String
    Dim ConR   As New ADODB.Recordset
    Dim r
    Dim rw     As Single

    ReDim RY(3, 140)

    dur = ds
    SQL = "select * from FRICSレーダー予測 where Time ='" & Format(dur, "yyyy/mm/dd hh:nn") & "' and " & _
          " Prediction_Minute IN( 60, 120, 180)"

    ConR.Open SQL, Con_水文, adOpenKeyset, adLockReadOnly

    If ConR.EOF Then
        LOG_Out "IN MB_FRICSレーダー予測指定データ取得できず"
        LOG_Out "SQL=" & SQL
        ORA_Message_Out "FRICSレーダ雨量受信", "FRICS降雨予測が計算時に洪水予測システムに取り込まれませんでした。計算をスキップします。", 1
        irc = False
        ConR.Close
        Exit Sub
    End If

    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        k = DateDiff("h", jgd, dw)
        m = CLng(ConR.Fields("Prediction_Minute").Value / 60 + 0.4)
    Debug.Print "  m=" & Format(m, "##0") & " k=" & Format(k, "000") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
    Debug.Print "FRICS i=" & Format(i, "000") & " Now_Step=" & Format(Now_Step, "##0")
        For j = 1 To 135
            b = Format(j, "###")
            rw = CSng(ConR.Fields(b).Value)
            If rw > 250 Then
                RY(m, j) = 0                       '流域雨量
            Else
                RY(m, j) = rw                      '流域雨量
            End If
        Next j
NOP:
       ConR.MoveNext
    Loop

    irc = True
    ConR.Close


End Sub
'**************************************************************
'HANS画面用FRICS予測降雨を取得
'
'ds=希望開始時刻
'
'
' RY(18, 140)  '10分ピッチで3時間後まで
'dw=取得最終時刻
'
'*************************************************************
Sub MDB_FRICSレーダー予測_For_HANS(ds As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim m      As Long
    Dim n      As Long
    Dim dur    As Date
    Dim dwr    As Date
    Dim dw     As Date
    Dim b      As String
    Dim SQL    As String
    Dim ConR   As New ADODB.Recordset
    Dim r
    Dim rw     As Single

    Dim RYg(18, 140)  As Single

    LOG_Out "IN MDB_FRICSレーダー予測_For_HANS " & ds

    Erase RhY '出力用エリアをクリヤする

    dur = ds
    SQL = "select * from FRICSレーダー予測 where Time ='" & Format(dur, "yyyy/mm/dd hh:nn") & "'"

    ConR.Open SQL, Con_水文, adOpenKeyset, adLockReadOnly

    If ConR.EOF Then
        irc = False
        ConR.Close
        Exit Sub
    End If

    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        m = Int(ConR.Fields("Prediction_Minute").Value / 10 + 0.4)
    Debug.Print "  m=" & Format(m, "##0") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
    Debug.Print "FRICS i=" & Format(i, "000") & " Now_Step=" & Format(Now_Step, "##0")
        For j = 1 To 135
            b = Format(j, "###")
            rw = CSng(ConR.Fields(b).Value)
            If rw > 250 Then
                RYg(m, j) = 0             '流域雨量 未計算等のときは0とする。
            Else
                RYg(m, j) = rw * 0.16667  '流域雨量 mm/hrなので1/6する。
            End If
        Next j
       ConR.MoveNext
    Loop

    irc = True
    ConR.Close


'予測流域平均雨量
    For j = 1 To 5      '5流域
        For m = 1 To 18 '10分ピッチで3時間分
            r = 0
            For k = 1 To R_Ave_Num
                i = R_Ave_N(k)
                r = r + R_Ave(j, k) * RYg(m, i)
            Next k
            RhY(j, m) = r
        Next m
    Next j


    LOG_Out "OUT MDB_FRICSレーダー予測_For_HANS "


End Sub

Sub レーダー雨量作図用流域平均雨量計算()

    Dim i     As Long
    Dim j     As Long
    Dim k     As Long
    Dim m     As Long
    Dim r     As Single

    LOG_Out "  In  レーダー雨量作図用流域平均雨量計算"

'実績雨量
    For i = 1 To Now_Step
        For j = 1 To 5
            r = 0
            For k = 1 To R_Ave_Num
                m = R_Ave_N(k)
                r = r + R_Ave(j, k) * rr(i, m)
            Next k
            RO(j, i) = r
        Next j
    Next i

'予測雨量
    For i = 1 To 3
        For j = 1 To 5
            r = 0
            For k = 1 To R_Ave_Num
                m = R_Ave_N(k)
                r = r + R_Ave(j, k) * RY(i, m)
            Next k
            RO(j, Now_Step + i) = r
        Next j
    Next i

    LOG_Out " Out  レーダー雨量作図用流域平均雨量計算"

End Sub

Sub 流域名読み込み()

    Dim k        As Integer
    Dim buf      As String
    Dim nf       As Integer

    LOG_Out "IN   Sub 流域名読み込み"

    nf = FreeFile
    Open App.Path & "\data\レーダーティーセン.dat" For Input As #nf

    k = 0
    Do Until EOF(nf)
        Line Input #nf, buf
        k = k + 1
        R_T_Name(k) = Trim(Mid(buf, 6, 5))
    Loop

    Close #nf


    LOG_Out "Out  Sub 流域名読み込み"

End Sub

'*************************************************
'検証時のルーティン
'
'現在使用せず
'
'************************************************
Sub レーダー雨量セット()

    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim m               As Integer
    Dim ii              As Integer
    Dim jj              As Integer
    Dim nf              As Integer
    Dim nh              As Integer
    Dim r               As Single
    Dim rw              As Single
    Dim buf             As String
    Dim dw              As Date
    Dim RM(140)         As Single
    Dim pfct            As Integer

    If MAIN.Check3 Then
        pfct = 3
    Else
        pfct = 0
    End If

    LOG_Out "IN   Sub レーダー雨量セット"
    LOG_Out App.Path & "\実績洪水\" & Radar_File & "  読み込み"

    nh = FreeFile
    Open App.Path & "\実績洪水\" & Radar_File For Input As #nh        'レーダーメッシュデータ
    Line Input #nh, buf

    For j = 1 To Now_Step + pfct
        Line Input #nh, buf
        For i = 1 To RRYU '135流域
            rw = 0#
            For k = 1 To 20
                m = R_Meshu(k, i)  'レーダーメッシュの番号
                If m = 0 Then Exit For
                r = CSng(Mid(buf, 17 + (m - 1) * 5, 5))
                rw = rw + r * R_Thissen(k, i)
            Next k
            If rw < 0# Then rw = 0#
            rr(j, i) = rw
        Next i

'ここから予測
        Line Input #nh, buf      '1時間後予測
        For i = 1 To RRYU '135流域
            rw = 0#
            For k = 1 To 20
                m = R_Meshu(k, i) 'レーダーメッシュの番号
                If m = 0 Then Exit For
                r = CSng(Mid(buf, 17 + (m - 1) * 5, 5))
                rw = rw + r * R_Thissen(k, i)
            Next k
            If rw < 0# Then rw = 0#
            RY(1, i) = rw
        Next i
        Line Input #nh, buf      '2時間後予測
        For i = 1 To RRYU '135流域
            rw = 0#
            For k = 1 To 20
                m = R_Meshu(k, i)  'レーダーメッシュの番号
                If m = 0 Then Exit For
                r = CSng(Mid(buf, 17 + (m - 1) * 5, 5))
                rw = rw + r * R_Thissen(k, i)
            Next k
            If rw < 0# Then rw = 0#
            RY(2, i) = rw
        Next i
        Line Input #nh, buf      '3時間後予測
        For i = 1 To RRYU '135流域
            rw = 0#
            For k = 1 To 20
                m = R_Meshu(k, i) 'レーダーメッシュの番号
                If m = 0 Then Exit For
                r = CSng(Mid(buf, 17 + (m - 1) * 5, 5))
                rw = rw + r * R_Thissen(k, i)
            Next k
            If rw < 0# Then rw = 0#
            RY(3, i) = rw
        Next i
    Next j

    Close #nh


End Sub
'
'**************************************************
'
'RSHINKAWA用にレーダー雨量を出力する
'
'検証用予測雨量なし（実績降雨で計算）
'
'
'
'**************************************************
'
Sub レーダー雨量出力_Veri()

    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim ii              As Long
    Dim jj              As Long
    Dim nf              As Long
    Dim k1              As Long
    Dim k2              As Long
    Dim Steps           As Long
    Dim buf             As String
    Dim dw              As Date


    If Verification2.Check1 <> vbChecked Then
        For i = 1 To 3
            For ii = 1 To RRYU
                If rr(Now_Step + i, ii) = 0# Then RY(i, ii) = 0.1
            Next ii
        Next i
        Steps = All_Step
    End If

    jj = Fix((All_Step - 1) / 12) + 1

'実績と予測レーダーデータ出力
    nf = FreeFile
    Open App.Path & "\work\流域平均雨量.dat" For Output As #nf  'レーダー実績流域平均雨量出力

    For ii = 1 To RRYU
        buf = Format(Format(ii, "####0"), "@@@@@") & "     1.E-1    1"
        Print #nf, buf

        For j = 1 To jj
            k1 = (j - 1) * 12 + 1
            k2 = k1 + 11
            If k2 > All_Step Then k2 = All_Step
            buf = ""
            For k = k1 To k2
                If rr(k, ii) > 0# Then
                    buf = buf & Format(Format(rr(k, ii) * 10, "####0"), "@@@@@")
                Else
                    buf = buf & "1.E-0"
                End If
            Next k
            Print #nf, Space(10) & buf
        Next j
    Next ii

    Close #nf

    レーダー雨量作図用流域平均雨量計算

End Sub
'**************************************************
'RSHINKAWA用にレーダー雨量を出力する
'
'2003/10/01 出力形式を変更
'
'
'
'**************************************************
Sub レーダー雨量出力()

    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim k1              As Long
    Dim k2              As Long
    Dim ii              As Long
    Dim jj              As Long
    Dim nf              As Long
    Dim buf             As String
    Dim dw              As Date


    For i = 1 To 3
        For ii = 1 To RRYU
            rr(Now_Step + i, ii) = RY(i, ii)
        Next ii
    Next i

    jj = Fix((All_Step - 1) / 12) + 1

'実績と予測レーダーデータ出力
    nf = FreeFile
    Open App.Path & "\work\流域平均雨量.dat" For Output As #nf  'レーダー実績流域平均雨量出力

    For ii = 1 To RRYU
        buf = Format(Format(ii, "####0"), "@@@@@") & "     1.E-1    1"
        Print #nf, buf

        For j = 1 To jj
            k1 = (j - 1) * 12 + 1
            k2 = k1 + 11
            If k2 > All_Step Then k2 = All_Step
            buf = ""
            For k = k1 To k2
                If rr(k, ii) > 0# Then
                    buf = buf & Format(Format(rr(k, ii) * 10, "####0"), "@@@@@")
                Else
                    buf = buf & "1.E-0"
                End If
            Next k
            Print #nf, Space(10) & buf
        Next j
    Next ii

    Close #nf

    レーダー雨量作図用流域平均雨量計算

End Sub

Sub 基準地点と流域対応を読む()

    Dim i     As Long
    Dim j     As Long
    Dim n     As Long
    Dim buf   As String
    Dim nf    As Integer
    Dim a(5)  As Single
    Dim b     As Single
    Dim c(5)  As Integer
    Dim S     As String

    LOG_Out "  In  基準地点と流域対応を読む"

    nf = FreeFile
    Open App.Path & "\data\基準地点と流域対応.txt" For Input As #nf

    Line Input #nf, buf
    Line Input #nf, buf

    i = 0
    Do Until EOF(nf)
        Line Input #nf, buf
        i = i + 1
        c(1) = 1                                 '下之一色
        c(2) = IIf(Mid(buf, 1, 1) <> " ", 1, 0)  '大治
        c(3) = IIf(Mid(buf, 31, 1) <> " ", 1, 0) '水場川外水位
        c(4) = IIf(Mid(buf, 21, 1) <> " ", 1, 0) '久地野
        c(5) = IIf(Mid(buf, 11, 1) <> " ", 1, 0) '春日
        b = CSng(Mid(buf, 46, 10))
        For j = 1 To 5
            If c(j) > 0 Then
                R_Ave(j, i) = b
                a(j) = a(j) + b
            End If
        Next j
        S = Trim(Mid(buf, 40, 5))
        Find_Rname S, n
        R_Ave_N(i) = n
    Loop
    R_Ave_Num = i
    For i = 1 To R_Ave_Num
        For j = 1 To 5
'            If j = 3 Then Debug.Print " i="; i; "  a(j)="; a(j); "  R_Ave="; R_Ave(j, i); "  R_Ave(j, i) / a(j)="; R_Ave(j, i) / a(j)
            R_Ave(j, i) = R_Ave(j, i) / a(j)
        Next j
    Next i

    Close #nf

    LOG_Out " Out  基準地点と流域対応を読む"

End Sub

'**************************************************************
'気象庁レーダーデータ予測を取得
'
'雨量はmm/Hourで登録されている
'
'2kmメッシュ対応なので使用停止 2007/05/02
'
'**************************************************************
Sub MDB_気象庁レーダー予測(ds As Date, de As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Long
    Dim k      As Long
    Dim buf    As String
    Dim du     As Date
    Dim dl     As Date
    Dim dw     As Date
    Dim Conn   As String
    Dim ConS   As New ADODB.Connection
    Dim ConR   As New ADODB.Recordset
    Dim a
    Dim SQL    As String
    Dim mi     As String
    Dim rw     As Single

    LOG_Out "IN   Sub MDB_気象庁レーダー予測"

'    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & App.Path & "\data\水文.mdb"
'    Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\SHINKAWA\OracleTest\oraDB\Data\水文.mdb"
    Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & 水文MDB
    ConS.ConnectionString = Conn
    ConS.Open

    irc = False

    ReDim RY(3, 140)

    If Err <> 0 Then
        MsgBox "水文.MDBにアクセスできません、水文.MDBの有無を確認してください。" & vbCrLf & _
               "計算できませんのでプロブラムは終了します。", vbExclamation
        End
    End If

    Set ConR.ActiveConnection = ConS

    mi = Fix(Minute(de) / 10) * 10

'１時間後
    SQL = "select * from 気象庁レーダー予測_1 where Time= '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and " & _
          "  Minute = " & mi    '& " order by Time"

    ConR.Open SQL, ConS, adOpenDynamic, adLockOptimistic

    If ConR.EOF Then
        LOG_Out "IN MB_気象庁レーダー予測指定データ取得できず"
        LOG_Out Format(ds, "yyyy/mm/dd hh:nn") & " ～ " & Format(de, "yyyy/mm/dd hh:nn")
        ConR.Close
        Exit Sub
    End If

    dw = ConR.Fields("Time").Value
    k = DateDiff("h", jgd, dw)
'    Debug.Print " k=" & Format(k, "000") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
    If k <> 1 Then
        LOG_Out "IN MDB_気象庁レーダー予測 ここにきてはいけません。"
        LOG_Out " jgd=" & Format(jgd, "yyyy/mm/dd hh:nn")
        LOG_Out "  ds=" & Format(ds, "yyyy/mm/dd hh:nn")
        LOG_Out "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
    End If
    Debug.Print "  dw="; dw; "  ";
    For j = 1 To 135
        a = Format(j, "###")
        RY(k, j) = ConR.Fields(a).Value * 0.1  '流域雨量
        Debug.Print Format(RY(k, j), "##0.0 ");
    Next j
    Debug.Print ""
    ConR.Close

'２～３時間後
    For k = 2 To 3
        dl = DateAdd("h", k, jgd)
        SQL = "select * from 気象庁レーダー予測_2 where Time= '" & Format(dl, "yyyy/mm/dd hh:nn") & "'"

        ConR.Open SQL, ConS, adOpenDynamic, adLockOptimistic

        If ConR.EOF Then
            LOG_Out "IN MB_気象庁レーダー予測_2指定データ取得できず 時刻=" & Format(dl, "yyyy/mm/dd hh:nn")
            ORA_Message_Out "気象庁レーダ雨量受信", "気象庁短時間降雨予測が取得できていません。", 1
            irc = False
            ConR.Close
            Exit Sub
        Else
            For j = 1 To 135
                a = Format(j, "###")
                RY(k, j) = RY(k, j) + ConR.Fields(a).Value * 0.1 '流域雨量
            Next j
            Debug.Print "  dl="; dl; "  ";
            For j = 1 To 135
                Debug.Print Format(RY(k, j), "##0.0 ");
            Next j
            Debug.Print ""
        End If
        irc = True
        ConR.Close
    Next k

End Sub
'**************************************************************
'気象庁レーダーデータ予測を取得
'
'雨量はmm/10で登録されている
'
'2007/05/02 22:37 新規作成
'
'**************************************************************
Sub MDB_気象庁レーダー予測2(ds As Date, de As Date, irc As Boolean)

    Dim i      As Integer
    Dim j      As Integer
    Dim k      As Integer
    Dim n      As Integer
    Dim m      As Integer
    Dim a
    Dim SQL    As String
    Dim d1     As Date
    Dim d2     As Date
    Dim dw     As Date
    Dim d1c    As String
    Dim dsc    As String
    Dim dec    As String

    Dim RM(140)         As Single

    ReDim RY(3, 140)

    dsc = TIMEC(ds)
    dec = TIMEC(de)

    LOG_Out "IN   Sub MDB_気象庁レーダー予測2 " & dsc & "～ " & dec
    JMA_OUT "IN   Sub MDB_気象庁レーダー予測2 " & dsc & "～ " & dec

    n = DateDiff("h", ds, de) + 1
    d1 = DateAdd("n", -50, ds)
    d2 = de
    SQL = "select * from 気象庁レーダー予測_1 where Time between '" & TIMEC(d1) & "' and '" & _
           TIMEC(d2) & "' ORDER BY Time"

    Rec_水文.Open SQL, Con_水文, adOpenDynamic, adLockOptimistic
    d1 = DateAdd("n", -50, ds)
    d2 = d1
    dw = ds
    For k = 1 To n
        Erase RM
        JMA_OUT "                    時間雨量作成 " & TIMEC(dw)
        For m = 1 To 6 '6ステップを足して時間雨量にする
            d1c = TIMEC(d1)
            Rec_水文.Find "Time = '" & d1c & "'"
            If Rec_水文.EOF Then
                LOG_Out "IN MDB_気象庁レーダー予測2 指定データ取得できず"
                LOG_Out d1c
                JMA_OUT "                 時刻雨量取得 " & d1c & " 取得できず"
            Else
                JMA_OUT "                 時刻雨量取得 " & d1c
                For j = 1 To 135
                    a = Format(j, "###")
                    RM(j) = RM(j) + Rec_水文.Fields(a).Value * 0.1  '流域雨量
                Next j
            End If
            Rec_水文.MoveFirst
            d1 = DateAdd("n", 10, d1)
        Next m

        For j = 1 To 135
            RY(k, j) = RM(j)    '流域雨量
        Next j

        d1 = DateAdd("h", k, d2)
        dw = DateAdd("h", 1, dw)

    Next k
    irc = True
    Rec_水文.Close

    LOG_Out "Out  Sub MDB_気象庁レーダー予測2 " & dsc & "～ " & dec
    JMA_OUT "Out  Sub MDB_気象庁レーダー予測2 " & dsc & "～ " & dec

End Sub
'**************************************************************
'気象庁レーダーデータ実績を取得
'
'現在使っていない
'
'**************************************************************
Sub MDB_気象庁レーダー実績(ds As Date, de As Date, dw As Date, irc As Boolean)


    Dim Conn   As String
    Dim ConS   As New ADODB.Connection
    Dim ConR   As New ADODB.Recordset

    Dim i      As Integer
    Dim j      As Integer
    Dim k      As Integer
    Dim n      As Integer
    Dim a
    Dim SQL    As String
    Dim d1     As Date
    Dim d2     As Date

    ReDim rr(500, 140)

    LOG_Out "IN   Sub MDB_気象庁レーダー実績 " & ds & "～ " & de

'    Con = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & App.Path & "\data\水文.mdb"
'    Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=D:\SHINKAWA\OracleTest\oraDB\Data\水文.mdb"
    Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & 水文MDB
    ConS.ConnectionString = Conn
    ConS.Open

    If Err <> 0 Then
        MsgBox "水文.MDBにアクセスできません、水文.MDBの有無を確認してください。" & vbCrLf & _
               "計算できませんのでプロブラムは終了します。", vbExclamation
        End
    End If

    Set ConR.ActiveConnection = ConS

    n = DateDiff("h", ds, de) + 1
    d1 = DateAdd("n", -50, ds)
    d2 = ds
    For k = 1 To n

        ReDim RM(140) As Single

        SQL = "select * from 気象庁レーダー実績 where Time between '" & Format(d1, "yyyy/mm/dd hh:nn") & "' and '" & _
               Format(d2, "yyyy/mm/dd hh:nn") & "' "

        ConR.Open SQL, ConS, adOpenDynamic, adLockOptimistic
        If ConR.EOF Then
            LOG_Out "IN MB_気象庁レーダー実績指定データ取得できず"
            LOG_Out Format(d1, "yyyy/mm/dd hh:nn") & " ～ " & Format(d2, "yyyy/mm/dd hh:nn")
        Else
            Do Until ConR.EOF
                dw = ConR.Fields("Time").Value
                i = DateDiff("h", ds, dw) + 1
'     Debug.Print " i=" & Format(i, "000") & "  dw=" & Format(dw, "yyyy/mm/dd hh:nn")
                For j = 1 To 135
                    a = Format(j, "###")
                    RM(j) = RM(j) + ConR.Fields(a).Value * 0.1  '流域雨量
                Next j
                ConR.MoveNext
            Loop
        End If

'        Debug.Print "  k="; k; "  d2="; d2; "  RM="; RM(1)
        For j = 1 To 135
            rr(k, j) = RM(j)    '流域雨量
        Next j

        ConR.Close
        d1 = DateAdd("h", 1, d1)
        d2 = DateAdd("h", 1, d2)

    Next k
    irc = True
    dw = de

    LOG_Out "Out  Sub MDB_気象庁レーダー実績 " & ds & "～ " & de

End Sub
'
'**************************************************************
'
'気象庁レーダーデータ実績を取得
'このサブルーティンは検証用です。                     ????????2010/03/09
'予測雨量を使わないで実績データのみで計算できるように ????????2010/03/09
'設定なっています、予測雨量は読みません。             ????????2010/03/09
'
'雨量はmm/10minで登録されている
'
'
'**************************************************************
'
Sub MDB_気象庁レーダー実績2(ds As Date, de As Date, dw1 As Date, irc As Boolean)

    Dim i      As Integer
    Dim j      As Integer
    Dim k      As Integer
    Dim n      As Integer
    Dim m      As Integer
    Dim a
    Dim SQL    As String
    Dim d1     As Date
    Dim d2     As Date
    Dim dw     As Date
    Dim d1c    As String
    Dim dsc    As String
    Dim dec    As String

    Dim RM(140)         As Single

    dsc = TIMEC(ds)
    dec = TIMEC(de)

    LOG_Out "IN   Sub MDB_気象庁レーダー実績2 " & dsc & "～ " & dec
    JMA_OUT "IN   Sub MDB_気象庁レーダー実績2 " & dsc & "～ " & dec

    ReDim rr(500, 140)

    n = DateDiff("h", ds, de) + 1
    d1 = DateAdd("n", -50, ds)
    d2 = de
    SQL = "select * from 気象庁レーダー実績 where Time between '" & TIMEC(d1) & "' and '" & _
           TIMEC(d2) & "' ORDER BY Time"

    Rec_水文.Open SQL, Con_水文, adOpenDynamic, adLockOptimistic
    Do
        Debug.Print Rec_水文.Fields("Time").Value
        Rec_水文.MoveNext
    Loop Until Rec_水文.EOF
    Rec_水文.MoveFirst
    JMA_OUT SQL
    d1 = DateAdd("n", -50, ds)
    d2 = d1
    dw = ds
    For k = 1 To n
        Erase RM
        JMA_OUT "                    時間雨量作成 " & TIMEC(dw)
        For m = 1 To 6 '6ステップを足して時間雨量にする
            d1c = TIMEC(d1)
            Rec_水文.Find "Time ='" & d1c & "'"
            If Rec_水文.EOF Then
                LOG_Out "IN MB_気象庁レーダー実績指定データ取得できず"
                LOG_Out d1c
                JMA_OUT "                 時刻雨量取得 " & d1c & " 取得できず"
            Else
                JMA_OUT "                 時刻雨量取得 " & d1c
                For j = 1 To 135
                    a = Format(j, "###")
                    RM(j) = RM(j) + Rec_水文.Fields(a).Value * 0.1  '流域雨量
                Next j
            End If
            Rec_水文.MoveFirst
            d1 = DateAdd("n", 10, d1)
        Next m

        For j = 1 To 135
            rr(k, j) = RM(j)    '流域雨量
        Next j

        d1 = DateAdd("h", k, d2)
        dw = DateAdd("h", 1, dw)

    Next k
    irc = True
    dw1 = de
    Rec_水文.Close

    LOG_Out "Out  Sub MDB_気象庁レーダー実績2 " & dsc & "～ " & dec
    JMA_OUT "Out  Sub MDB_気象庁レーダー実績2 " & dsc & "～ " & dec

End Sub
