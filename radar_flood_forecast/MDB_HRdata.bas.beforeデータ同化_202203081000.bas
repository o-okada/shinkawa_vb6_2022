Attribute VB_Name = "MDB_HRdata"
Option Explicit
Option Base 1
Public wH(6, 25)   As Single
Public DH_Tide     As Single
'
'水文ＤＢよりデータを取り込む
'
'修正記録
'欠測の補填は現時刻だけとした 2004/03/24
'
Sub Data_IN(ds As Date, de As Date, irc As Boolean)

    Dim i      As Long
    Dim j      As Integer
    Dim k      As Integer
    Dim m      As Integer
    Dim b      As String
    Dim du     As Date
    Dim dw     As Date
    Dim dur    As Date
    Dim dwr    As Date
    Dim ConR   As New ADODB.Recordset
    Dim a
    Dim SQL    As String
    Dim mi     As String
    Dim C0     As Single
    Dim C1     As Single
    Dim C2     As Single
    Dim C3     As Single
    Dim ch     As Boolean
    Dim uh     As Boolean
    Dim hw(4)  As Single
    Dim er     As Boolean

    If Err <> 0 Then
        MsgBox "水文.MDBにアクセスできません、水文.MDBの有無を確認してください。" & vbCrLf & _
               "計算できませんのでプロブラムは終了します。", vbExclamation
        End
    End If

    mi = Fix(Minute(de) / 10) * 10

'水位取得
    SQL = "select * from 水位 where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' and Minute = " & mi & " order by Time"
    Short_Break 4
    ConR.Open SQL, Con_水文, adOpenKeyset, adLockReadOnly
    i = 0
    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        If i = 0 Then
            du = dw
        End If
        i = DateDiff("h", du, dw) + 1
        HO(1, i) = ConR.Fields("Tide").Value       'Tide 日光川外水位
        HO(2, i) = 0#                              '洗堰越流量
        HO(3, i) = ConR.Fields("下之一色").Value   '下之一色
        HO(4, i) = ConR.Fields("大治").Value       '大治
        HO(5, i) = ConR.Fields("水場川外").Value   '水場外
        HO(6, i) = ConR.Fields("久地野").Value     '久地野
        HO(7, i) = ConR.Fields("春日").Value       '春日
        ConR.MoveNext
    Loop
    ConR.Close

    If HO(1, Now_Step) < -50# Then
        Tide_Special
        ORA_Message_Out "水位データ受信", "日光川外水位データが欠測しました。天文潮位に直前の実況値との差分を加えて、現況・予測値とします。", 1
    Else
        DH_Tide = 0#
    End If

'予測潮位臨時
'    TidalY dw, C0, C1, C2, C3      '気象庁潮位表から天文潮位を内挿する
    Cal_Tide dw, C0, C1, C2, C3    '60分調から天文潮位を計算する
    If HO(1, Now_Step) < -50# Then
        HO(1, Now_Step) = C0 + DH_Tide
    Else
        DH_Tide = HO(1, Now_Step) - C0
    End If
    HO(1, Now_Step + 1) = C1 + DH_Tide
    HO(1, Now_Step + 2) = C2 + DH_Tide
    HO(1, Now_Step + 3) = C3 + DH_Tide

    If i = 0 Then
'        MsgBox "ローカルDBに水位データがありません。"
        LOG_Out "ローカルDBに水位データがありません。"
        ds = CDate("1900/01/01 01:00")
        de = CDate("1900/01/01 01:00")
        Exit Sub
    End If


'
'実績水位最終データ日付の予測データを
'取りに行く
'

    Set ConR = Nothing

    jsd = du
    js(1) = Year(jsd)
    js(2) = Month(jsd)
    js(3) = Day(jsd)
    js(4) = Hour(jsd)
    js(5) = Minute(jsd)
    js(6) = 0
    jgd = dw
    jg(1) = Year(jgd)
    jg(2) = Month(jgd)
    jg(3) = Day(jgd)
    jg(4) = Hour(jgd)
    jg(5) = Minute(jgd)
    jg(6) = 0
    Now_Step = DateDiff("h", jsd, jgd) + 1
    All_Step = Now_Step + Yosoku_Step

    If Now_Step <= 4 Then
        'LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない！！！"
        '修正開始　2016/09/23　O.OKADA　ここからコメントアウトする。
        '修正理由　計算時刻が常に15分程度遅れているため。
        'Exit Sub
        '修正終了　2016/09/23　O.OKADA　ここまでコメントアウトする。
    End If
    If ds = de Or All_Step < 3 Then
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        LOG_Out "IN  Data_IN  計算に使用する水位データのステップがすくないので計算中止しない？？？"
        '修正開始　2016/09/23　O.OKADA　ここからコメントアウトする。
        '修正理由　計算時刻が常に15分程度遅れているため。
        'Exit Sub
        '修正終了　2016/09/23　O.OKADA　ここまでコメントアウトする。
    End If

'下流端境界条件日光川外水位の補填
    '修正１********************************
    For i = 1 To Now_Step
        If HO(1, i) < -50# Then
            j = 1
            Select Case i
                Case 1
                    HO(1, 1) = 1.5
                Case Is > 1
                     HO(1, i) = HO(1, i - 1)
            End Select
        End If
    Next i

'欠測補填
    er = False
    For i = 1 To 7
        If i <> 2 Then
            ch = False
'            For j = Now_Step - 3 To Now_Step　2004/03/24
            For j = Now_Step To Now_Step
                a = HO(i, j)
                If a < -50# Then
                    ch = True
                    GoTo J1
                End If
            Next j
        End If
    Next i
J1:
    If ch Then
        Pre_水位欠測補填
        For i = 1 To 7
            If i <> 2 Then
                uh = True
'                For j = Now_Step - 3 To Now_Step  2004/03/24
                For j = Now_Step To Now_Step
                    a = HO(i, j)
                    If a < -50# Then
                        uh = True
'                        For k = Now_Step - 3 To Now_Step  2004/03/24
                        For k = Now_Step To Now_Step
'                            m = k - (Now_Step - 3) + 1  2004/03/24
                            m = k - Now_Step + 1
                            hw(m) = HO(i, k)
                        Next k
                        If hw(m) < -50# Then
                            er = True
                            ORA_Message_Out "テレメータ水位受信", Name_H(i) & "の、水位データが欠測しました。洪水予測システムによる結果を用いて水位予測計算を行います。", 1
                        End If
                        Exit For
                    End If
                Next j
            End If
        Next i
    End If
    If HO(1, Now_Step) < -50# Then
        er = True
    End If
    irc = True
    If (AutoDrive.Check6 = vbChecked) And er Then '欠測補填を手入力する
        Load Data_Edit
        Unload Data_Edit
    End If
    If (AutoDrive.Check6 = vbUnchecked) And er Then '欠測なので計算をスキップする
'        irc = False '欠測でも計算するように修正 2004/4/26
'        Exit Sub
    End If

'    Dim nf As Long
'
'    nf = FreeFile
'    open app.Path & "\data\潮位スライド量.dat" for output
'    LOG_Out "IN  Data_IN  潮位スライド量 CX=" & Format(cx, "###0.000")

'    MDB_洗堰 jsd, jgd, er

End Sub
Sub Pre_水位欠測補填()

    Dim ConR        As New ADODB.Recordset
    Dim SQL         As String
    Dim ds          As Date
    Dim de          As Date
    Dim i           As Long
    Dim j           As Long

    ds = DateAdd("h", -4, jgd)
    de = jgd

    SQL = "select * from 水位 where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' order by Time"

    ConR.Open SQL, Con_水文, adOpenKeyset, adLockReadOnly

    j = 1
    Do Until ConR.EOF
        For i = 1 To 6 '６水位観測所
            wH(i, j) = ConR.Fields(i + 1).Value
        Next i
        j = j + 1
        ConR.MoveNext
    Loop

    ConR.Close

End Sub
Sub Tide_Special()

    Dim SQL    As String
    Dim buf    As String
    Dim dw     As Date
    Dim w

    LOG_Out "IN   Tide_Special"

    On Error GoTo ER1

    DH_Tide = 0#

'    MDB_履歴_Connection

    dw = DateAdd("n", -10, jgd)

    If isRAIN = "02" Then
        SQL = "SELECT 日光川外水位 FROM FRICS履歴 WHERE TIME='" & Format(dw, "yyyy/mm/dd hh:nn") & "'"
    Else
        SQL = "SELECT 日光川外水位 FROM 気象庁履歴 WHERE TIME='" & Format(dw, "yyyy/mm/dd hh:nn") & "'"
    End If


    Rec_履歴.Open SQL, Con_履歴, adOpenDynamic, adLockReadOnly

    If Rec_履歴.EOF Then
        DH_Tide = 0#
    Else
        buf = Rec_履歴.Fields(0).Value
        w = Split(buf, ",")
        DH_Tide = w(0)
    End If

    Rec_履歴.Close

'    MDB_履歴_Close

    LOG_Out "OUT  Tide_Special DH_Tide=" & Format(DH_Tide, "###0.000")

    Exit Sub

ER1:
    LOG_Out "OUT  Tide_Special ABend DH_Tide=" & Format(DH_Tide, "###0.000")
    Rec_履歴.Close
    On Error GoTo 0

End Sub


