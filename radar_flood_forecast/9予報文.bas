Attribute VB_Name = "予報文作成"
Option Explicit
Option Base 1
Public Pattan_Now    As Long       '当該パターン
Public Message()     As Pattan     '文言
Public Wng_Last_Time As Long       '既往注意・警戒番号
Public NPat          As Long       '文言の種類数
Type Pattan
    Patn(16)   As Variant
       ' 1=パターン                2=①予報文の種別
       ' 3=②見出し                4=③主文
       ' 5=④注意・警戒情報        6=⑩洪水予測の水位状況
       ' 7=⑫洪水予測の水位状況２  8=⑭洪水予測の水位危険度レベル
       ' 9=③主文表示用           16=予報文種別番号コード
End Type
Public ICat          As Long       '選択したパターン番号
Public Msgz          As Long       '用いる主文番号
Public Const CYUBN_1 = "　　今回の出水は、平成3年9月の台風17・18号に匹敵する規模" & _
                       "と見込まれます。"

Public Const CYUBN_2 = "　　今回の出水は、平成3年9月の台風17・18号を上回る規模と" & _
                       "見込まれます。"

Public Const CYUBN_3 = "　　今回の出水は、平成12年9月の東海豪雨に匹敵する規模と" & _
                       "見込まれます。"

Public Add_Main_Message(20)    As String
Public 主文1                   As String
Public 主文2                   As String

Public 水位                    As six_time
Type six_time
     hm3     As Single '3時間前水位
     hm2     As Single '2時間前水位
     hm1     As Single '1時間前水位
     h       As Single '現在時刻水位
     hy1     As Single '1時間後水位
     hy2     As Single '2時間後水位
     hy3     As Single '3時間後水位
End Type
Sub Cat_Num_Read()

    Dim nf   As Long

    On Error GoTo jump

    nf = FreeFile
    Open App.Path & "\data\Cat.dat" For Input As #nf

    Print #nf, ICat

    Close #nf

    Exit Sub

jump:
    ICat = 1
    On Error GoTo 0

End Sub
Sub Cat_Num_Write()

    Dim nf   As Long

    nf = FreeFile
    Open App.Path & "\data\Cat.dat" For Output As #nf

    Print #nf, Msgz

    Close #nf

End Sub
Sub Disp_Msg(i As Long)

    If i > NPat Then
        i = 1
    End If
    If i < 1 Then
        i = NPat
    End If

'    With 予報文送信
'        .Label7 = Message(i).Patn(2)
'        .Label8 = Message(i).Patn(1)
'        .Text5.Text = Message(i).Patn(9)
'    End With

    Msgz = i

End Sub
Sub Pattan_Add_Lf(w As Variant, ww As Variant, x As Long)

    Dim i    As Long
    Dim L    As Long
    Dim c    As String
    Dim cc   As String
    Dim LF

    If x = 1 Then
        LF = vbCrLf
    Else
        LF = vbLf
    End If

    ww = ""
    cc = ""
    L = Len(w)
    For i = 1 To L
        c = Mid(w, i, 1)
        If c = "%" Or c = LF Then
            cc = cc & LF & "　　"
        Else
            cc = cc & c
        End If
    Next i

    ww = cc

End Sub
Sub パターン文集_Read()

    Dim i     As Long
    Dim j     As Long
    Dim nf    As Long
    Dim buf   As String
    Dim m     As String
    Dim w
    Dim ww
    Dim p     As Long

    LOG_Out "IN    パターン文集_Read"

    nf = FreeFile
    Open App.Path & "\Data\パターンMK.txt" For Input As #nf

'文言の数を調べる
    i = 0
    Do
        Line Input #nf, buf
        i = i + 1
    Loop Until EOF(nf)
    Close #nf
    NPat = i - 2
    ReDim Message(NPat)

    Open App.Path & "\Data\パターンMK.txt" For Input As #nf
    Line Input #nf, buf
    Line Input #nf, buf
    For i = 1 To NPat
        Line Input #nf, buf
        w = Split(buf, vbTab)
        ww = "　　（主文）%新川の清須市水場川外水位水位観測所では、%" & w(4)
        Pattan_Add_Lf ww, w(4), 0
        Pattan_Add_Lf ww, w(9), 1
        p = CLng(w(1))
        For j = 1 To 16
'            If j = 4 Or j = 9 Then
                Message(p).Patn(j) = w(j)           '予報文にかかわる種別
'            Else
'                Message(p).Patn(j) = Trim(w(j))     '予報文にかかわる種別
'            End If
        Next j
        m = "新川　" & Message(p).Patn(3)
        Message(p).Patn(3) = m
        ' 1=パターン番号            2=①予報文の種別
        ' 3=②見出し                4=③主文
        ' 5=④注意・警戒情報        6=⑩洪水予測の水位状況
        ' 7=⑫洪水予測の水位状況２  8=⑭洪水予測の水位危険度レベル
        '16=予報文種別番号コード
    Next i
    Close #nf

    LOG_Out "OUT   パターン文集_Read"

End Sub
'
'
' 1　　新川の水位は１１日１６時４０分現在、次のとおりとなっています。
' 2　　水場川外水位観測所［新川町大字阿原地内］で、３．８０ｍ------以降に追加（レベル2超過）
' 3　　　　　　　　　　　　　　　　　　　　　　-----以降に追加（１時間に４０cmの速さで上昇中）
' 4　　⑩追加 + １１日１９時４０分頃には、---以降に追加 はん濫危険水位に達すると
' 5　　見込まれます。
' 6　　水場川外水位観測所［新川町大字阿原地内］で、６．２０ｍ　---以降に追加（水位危険度レベル４）
' 7　【参考】
' 8　　水場川外水位観測所［清須市新川町大字阿原地内］
' 9　　受け持ち区間
'10　　左右岸とも、庄内川分岐点（ＯＯ市＊＊町）から海（＋＋市？？町）まで
'11　　水位危険度レベル
'12　　■レベル１　水防団待機水位超過   ：２．０ｍ～３．０ｍ
'13　　■レベル２　はん濫注意水位超過   ：３．０ｍ～４．４ｍ
'14　　■レベル３　避難判断水位超過 　　：４．４ｍ～５．２ｍ
'15　　■レベル４　はん濫危険水位超過   ：５．２ｍ～＊．＊ｍ
'16　　■レベル５　はん濫の発生
'17　【問い合わせ先】
'18　　水位関係　　　愛知県尾張建設事務所  維持管理課   電話  ０５２－９６１－４４２１
'
'
'作成される文言例
'
'□□新川水場川外水位観測所（清須市新川町大字阿原地内）では、
'□□新川がさらに増水し、２時間後には、はん濫危険水位に到達する見込みです。市町において避難すべきと判断される場合がありますので、OO市OO地区から＋＋地区では、市町からの避難情報に注意して下さい。
'□【注意・警戒情報】
'□□今回の出水は、平成１２年９月の東海豪雨に匹敵する規模と見込まれます。
'□□また、越水の恐れがありますので厳重な警戒が必要です。
'□【現況・予想】
'□□新川上流域の流域平均雨量
'□□１１日１０時４０分から１１日１６時４０分までの６時間の現況□１００ミリ
'□□１１日１６時４０分から１１日１９時４０分までの３時間の予想□８０ミリ
'□□新川の水位は１１日１６時４０分現在、次のとおりとなっています。
'□□水場川外水位観測所［新川町大字阿原地内］で、３．８０ｍ□（レベル2超過）
'□□□□□□□□□□□□□□□□□（１時間に４０cmの速さで上昇中）
'□□新川の水位は、上昇傾向にあり、１１日１９時４０分頃には、はん濫危険水位に達すると見込まれます。
'□□新川の水位は､上昇傾向にあり､03日14時36分頃には､新川の水位は､上昇傾向にあり､
'□□水場川外水位観測所［新川町大字阿原地内］で、６．２０ｍ□（水位危険度レベル４）
'□【参考】
'□□水場川外水位観測所［清須市新川町大字阿原地内］
'□□受け持ち区間
'□□左右岸とも、庄内川分岐点（ＯＯ市＊＊町）から海（＋＋市？？町）まで
'□□■レベル１□水防団待機水位超過□□□：２．０m～３．０ｍ
'□□■レベル２□はん濫注意水位超過□□□：３．０ｍ～４．４ｍ
'□□■レベル３□避難判断水位超過□□□□：４．４ｍ～５．２ｍ
'□□■レベル４□はん濫危険水位超過□□□：５．２ｍ～＊．＊ｍ
'□□■レベル５□はん濫の発生
'□【問い合わせ先】
'□水位関係□□□愛知県尾張建設事務所□維持管理課□電話□０５２－９６１－４４２１
'
'
'
Sub 主文作成1()

    Dim m     As String
    Dim m1    As String  '【主文】
    Dim m2    As String  '【注意・警戒情報】
    Dim m3    As String  '【現況・予測】
    Dim m4    As String
    Dim m5    As String
    Dim m6    As String
    Dim mw    As String
    Dim dw    As Date
    Dim h0    As Single
    Dim dh    As Single

    h0 = 水位.h
    dh = (h0 - 水位.hm2) * 100#

    主文1 = ""
    主文2 = ""

    m1 = Message(Pattan_Now).Patn(4)

    注意警戒情報 m2

    m3 = vbLf & "　　（現況・予想）"

    主文1 = m1 & m2 & m3


    m3 = "　　新川の水位は" & Format(jgd, "dd月hh時nn分")
    m3 = m3 & "現在、次のとおりとなっています。"
    m3 = m3 & vbLf & "　　水場川外水位観測所［清須市］で、"
    m3 = m3 & Format(h0, "#0.00") & "m"
    水位レベル_Check 水位.h, mw
    m3 = m3 & mw
    水位変動_Check mw, dh
    If InStr(mw, "横") = 0 Then
        m3 = m3 & vbLf & "　　　　　　　　　　　　　　　　　　　　（１時間に"
        m3 = m3 & Format(Abs(dh * 0.5), "##0") & "cmの速さで"
    Else
        m3 = m3 & vbLf & "　　　　　　　　　　　　　　　　　　　　"
    End If
    m3 = m3 & mw
    If Pattan_Now <> 4 Then
        If Pattan_Now <> 14 Then
            m3 = m3 & vbLf & "　　" & Message(Pattan_Now).Patn(11)
        End If
        If Pattan_Now = 1 Then
            dw = DateAdd("h", 2, jgd)
            mw = Format(dw, "dd日hh時nn分頃には、")
        Else
            dw = DateAdd("h", 3, jgd)
            mw = Format(dw, "dd日hh時nn分頃には、")
        End If
        If Pattan_Now <> 14 Then
            m3 = m3 & mw & vbLf & "　　" & Message(Pattan_Now).Patn(13)
            m3 = m3 & "と見込まれます。"
            m4 = vbLf & "　　水場川が水位観測所［清須市］で、"
        End If
        If Pattan_Now = 1 Then
            mw = Format(水位.hy2, "#0.0") & "0m"
            水位レベル_Check 水位.hy2 + 0.04, m
        Else
            mw = Format(水位.hy3 + 0.04, "#0.0") & "0m"
            水位レベル_Check 水位.hy3 + 0.04, m
        End If
        m4 = m4 & mw & m
    Else
        m4 = m3
        m3 = ""
    End If
    If Pattan_Now = 14 Then
        m4 = ""
    End If

    m5 = vbLf & "　　　【参考】"
    m5 = m5 & vbLf & "　　　水場川外水位観測所［清須市阿原］"
    m5 = m5 & vbLf & "　　　受け持ち区間"
    m5 = m5 & vbLf & "　　　左右岸とも、庄内川分岐点から海まで"
    m5 = m5 & vbLf & "　　　水位危険度レベル"
    m5 = m5 & vbLf & "　　　■レベル１　水防団待機水位超過　　：２．０ｍ～３．０ｍ"
    m5 = m5 & vbLf & "　　　■レベル２　はん濫注意水位超過　　：３．０ｍ～４．４ｍ"
    m5 = m5 & vbLf & "　　　■レベル３　避難判断水位超過　　　：４．４ｍ～５．２ｍ"
    m5 = m5 & vbLf & "　　　■レベル４　はん濫危険水位超過　　：５．２ｍ～＊．＊ｍ"
    m5 = m5 & vbLf & "　　　■レベル５　はん濫の発生"
    m5 = m5 & vbLf & "　　　〔問い合わせ先〕"
    m5 = m5 & vbLf & "　　水位関係：愛知県　尾張建設事務所　　維持管理課　電話 052-961-4421"

    主文2 = m3 & m4 & m5

Debug.Print "　　見出し(" & Message(Pattan_Now).Patn(3) & ")"
Debug.Print " "
Debug.Print 主文1
Debug.Print 主文2










End Sub
'
'
' 1　　新川の水位は１１日１６時４０分現在、次のとおりとなっています。
' 2　　水場川外水位観測所［新川町大字阿原地内］で、３．８０ｍ------以降に追加（レベル2超過）
' 3　　　　　　　　　　　　　　　　　　　　　　-----以降に追加（１時間に４０cmの速さで上昇中）
' 4　　⑩追加 + １１日１９時４０分頃には、---以降に追加 はん濫危険水位に達すると
' 5　　見込まれます。
' 6　　水場川外水位観測所［新川町大字阿原地内］で、６．２０ｍ　---以降に追加（水位危険度レベル４）
' 7　【参考】
' 8　　水場川外水位観測所［清須市新川町大字阿原地内］
' 9　　受け持ち区間
'10　　左右岸とも、庄内川分岐点（ＯＯ市＊＊町）から海（＋＋市？？町）まで
'11　　水位危険度レベル
'12　　■レベル１　水防団待機水位超過   ：２．０ｍ～３．０ｍ
'13　　■レベル２　はん濫注意水位超過   ：３．０ｍ～４．４ｍ
'14　　■レベル３　避難判断水位超過 　　：４．４ｍ～５．２ｍ
'15　　■レベル４　はん濫危険水位超過   ：５．２ｍ～＊．＊ｍ
'16　　■レベル５　はん濫の発生
'17　【問い合わせ先】
'18　　水位関係　　　愛知県尾張建設事務所  維持管理課   電話  ０５２－９６１－４４２１
'
'
'作成される文言例
'

'　　はん濫注意水位に到達、水位はさらに上昇するおそれ
'主文1
'　　（主文）
'　　新川の清須市水場川外水位水位観測所では、
'　　はん濫注意水位(レベル2）に達しました。水位はさらに上昇する見込みです。
'　　今後の洪水予報に注意して下さい。
'
'　　（現況・予想）
'主文2
'　　新川の水場川外水位水位観測所〔清須市〕の水位
'　　19月05時50分の現況　3.36m（急上昇中）（水位危険度レベル２）
'　　19日07時50分の予想　4.50m（水位危険度レベル３）
'　　【参考】
'　　　水場川外水位水位観測所〔清須市阿原〕
'　　　はん濫危険水位　5.20m　　　　　　避難判断水位　　4.40m
'　　　はん濫注意水位（警戒水位）3.00m　水防団待機水位　2.00m
'
'　　水位危険度レベル
'　　　■レベル５　はん濫の発生
'　　　■レベル４　はん濫危険水位超過
'　　　■レベル３　避難判断水位超過
'　　　■レベル２　はん濫注意水位超過
'　　　■レベル１　水防団待機水位超過
'
'
'
Sub 主文作成2()

    Dim m     As String
    Dim m1    As String  '【主文】
    Dim m2    As String  '【注意・警戒情報】
    Dim m3    As String  '【現況・予測】
    Dim m4    As String
    Dim m5    As String
    Dim m6    As String
    Dim mw    As String
    Dim dw    As Date
    Dim h0    As Single
    Dim dh    As Single

    h0 = 水位.h
    dh = (h0 - 水位.hm2) * 100#

    主文1 = ""
    主文2 = ""

    m1 = Message(Pattan_Now).Patn(4)

    注意警戒情報 m2

    m3 = vbLf & "　　（現況・予想）"

    主文1 = m1 & m2 & m3


    m3 = "　　新川の水場川外水位水位観測所〔清須市〕の水位" & vbLf
    m3 = m3 & "　　" & Format(jgd, "dd日hh時nn分") & "の現況　" & Format(h0, "#0.00") & "m"
    水位変動_Check mw, dh
    m3 = m3 & mw
    水位レベル_Check 水位.h, mw
    m3 = m3 & mw
    If Pattan_Now <> 4 Then
        If Pattan_Now <> 14 Then
            m3 = m3 & vbLf & "　　"   ' & Message(Pattan_Now).Patn(11)
        End If
        If Pattan_Now = 1 Then
            dw = DateAdd("h", 2, jgd)
            mw = Format(dw, "dd日hh時nn分の予想　")
        Else
            dw = DateAdd("h", 3, jgd)
            mw = Format(dw, "dd日hh時nn分の予想　")
        End If
        If Pattan_Now <> 14 Then
            m3 = m3 & mw
'            m3 = m3 & "と見込まれます。"
'            m4 = vbLf & "　　水場川が水位観測所［清須市］で、"
        End If
        If Pattan_Now = 1 Then
            mw = Format(水位.hy2, "#0.0") & "0m"
            水位レベル_Check 水位.hy2 + 0.04, m
        Else
            mw = Format(水位.hy3 + 0.04, "#0.0") & "0m"
            水位レベル_Check 水位.hy3 + 0.04, m
        End If
        m4 = m4 & mw & m
    Else
        m4 = m3
        m3 = ""
    End If
    If Pattan_Now = 14 Then
        m4 = ""
    End If

    m5 = vbLf & "　　【参考】"
    m5 = m5 & vbLf & "　　　水場川外水位水位観測所〔清須市阿原〕"
    m5 = m5 & vbLf & "　　　はん濫危険水位　5.20m　　　　　　避難判断水位　　4.40m"
    m5 = m5 & vbLf & "　　　はん濫注意水位（警戒水位）3.00m　水防団待機水位　2.00m"
    m5 = m5 & vbLf & "　　　"
    m5 = m5 & vbLf & "　　水位危険度レベル"
    m5 = m5 & vbLf & "　　　■レベル５　はん濫の発生"
    m5 = m5 & vbLf & "　　　■レベル４　はん濫危険水位超過"
    m5 = m5 & vbLf & "　　　■レベル３　避難判断水位超過"
    m5 = m5 & vbLf & "　　　■レベル２　はん濫注意水位超過"
    m5 = m5 & vbLf & "　　　■レベル１　水防団待機水位超過"
    m5 = m5 & vbLf & " "
    m5 = m5 & vbLf & "　〔問い合わせ先〕"
    m5 = m5 & vbLf & "　　水位関係：愛知県　尾張建設事務所　　維持管理課　電話 052-961-4421"


    主文2 = m3 & m4 & m5

Debug.Print "　　見出し(" & Message(Pattan_Now).Patn(3) & ")"
Debug.Print " "
Debug.Print 主文1
Debug.Print 主文2










End Sub

Sub 水位レベル_Check(h As Single, m As String)


    If h < 2# Then
        m = ""
        Exit Sub
    End If

    Select Case h
        Case Is < 3#
            m = "（水位危険度レベル１）"

        Case Is < 4.4
            m = "（水位危険度レベル２）"

        Case Is < 5.2
            m = "（水位危険度レベル３）"

        Case Is < 100#
            m = "（水位危険度レベル４）"

    End Select

End Sub
Sub 水位変動_Check(hg As String, dh As Single)

    hg = ""

    Select Case dh
        Case Is > 30#
            hg = "（急上昇中）"
        Case Is > 10#
            hg = "（上昇中）"
        Case Is > -10#
            hg = "（現在の水位は横ばい）"
        Case Is > -110#
            hg = "（下降中）"
    End Select

End Sub
'
'H0   現在時刻水位水位
'H1   1時間後予測水位
'H2   2時間後予測水位
'H3   3時間後予測水位
'
'm--注意警戒情報
'
'
Sub 注意警戒情報(CYUBN As String)

    Dim Wng    As Long
    Dim HV     As Single
    Dim h0     As Single
    Dim H1     As Single
    Dim H2     As Single
    Dim H3     As Single

    h0 = 水位.h
    H1 = 水位.hy1
    H2 = 水位.hy2
    H3 = 水位.hy3

    If Pattan_Now < 5 Or Pattan_Now > 13 Then
        CYUBN = vbLf
        Exit Sub
    End If

'注意文
    HV = H3 - H2
    Select Case HV
        Case Is < 0.5
            Wng = 1
        Case Is < 1#
            Wng = 2
        Case Is >= 1#
            Wng = 3
    End Select
    If Wng_Last_Time > Wng Then Wng = Wng_Last_Time
    Wng_Last_Time = Wng
    Select Case Wng
        Case 1
            CYUBN = vbLf & "　　（注意事項）" & vbLf & CYUBN_1
        Case 2
            CYUBN = vbLf & "　　（注意事項）" & vbLf & CYUBN_2
        Case 3
            CYUBN = vbLf & "　　（注意事項）" & vbLf & CYUBN_3
    End Select
    If h0 >= 6.2 Or H3 >= 6.2 Then '計画堤防高(T.P 6.2m)を超える
        CYUBN = CYUBN & vbLf & "　　また、越水の恐れがありますので厳重な警戒が必要です。"
    End If

End Sub
'
'2008/03/03現在以下の19データを読む
'
' 1　　新川の水位は１１日１６時４０分現在、次のとおりとなっています。
' 2　　水場川外水位観測所［新川町大字阿原地内］で、３．８０ｍ------以降に追加（レベル2超過）
' 3　　　　　　　　　　　　　　　　　　　　　　-----以降に追加（１時間に４０cmの速さで上昇中）
' 4　　⑩追加 + １１日１９時４０分頃には、---以降に追加 はん濫危険水位に達すると
' 5　　見込まれます。
' 6　　水場川外水位観測所［新川町大字阿原地内］で、６．２０ｍ　---以降に追加（水位危険度レベル４）
' 7　【参考】
' 8　　水場川外水位観測所［清須市新川町大字阿原地内］
' 9　　受け持ち区間
'10　　左右岸とも、庄内川分岐点（ＯＯ市＊＊町）から海（＋＋市？？町）まで
'11　　水位危険度レベル
'12　　■レベル１　水防団待機水位超過   ：２．０ｍ～３．０ｍ
'13　　■レベル２　はん濫注意水位超過   ：３．０ｍ～４．４ｍ
'14　　■レベル３　避難判断水位超過     ：４．４ｍ～５．２ｍ
'15　　■レベル４　はん濫危険水位超過   ：５．２ｍ～＊．＊ｍ
'16　　■レベル５　はん濫の発生
'17　【問い合わせ先】
'18　　水位関係　　　愛知県尾張建設事務所  維持管理課   電話  ０５２－９６１－４４２１
'
'
'作成される文言例
'
'
'
'
'
Sub 追加主文_Read()

    Dim i      As Long
    Dim nf     As Long
    Dim buf    As String
    Dim f      As String

    f = App.Path & "\data\追加主文.txt"

    nf = FreeFile
    Open f For Input As #nf
    For i = 1 To 18
        Line Input #nf, Add_Main_Message(i)
    Next i
        
    Close #nf

End Sub
